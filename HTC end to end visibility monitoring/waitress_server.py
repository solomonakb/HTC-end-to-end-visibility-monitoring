import os
import csv
import logging
from io import StringIO
from datetime import datetime, timedelta
import threading
import time
from flask import Flask, request, jsonify, render_template, Response
from waitress import serve
import database

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'htc_monitor.db')
EXCEL_PATH = os.path.join(BASE_DIR, 'JUN-17-2026.xlsx')
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
NETWORK_SHARE_PATH = r'\\svhqftp03\HCT'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Initialize DB on startup
database.init_db(DB_PATH)

def check_initial_load():
    """Check if DB is empty and load default Excel if it exists."""
    stats = database.get_dashboard_stats(DB_PATH)
    if stats['total_events'] == 0 and os.path.exists(EXCEL_PATH):
        logger.info(f"Database empty. Loading initial data from {EXCEL_PATH}...")
        success, count = database.load_excel(DB_PATH, EXCEL_PATH)
        if success:
            logger.info(f"Successfully loaded {count} events.")
        else:
            logger.error("Failed to load initial data.")

def auto_fetch_reports():
    while True:
        try:
            logger.info("Starting auto-fetch of reports...")
            files = database.scan_share_directory(NETWORK_SHARE_PATH)
            if files:
                latest_file = files[0] # (filename, filepath, timestamp)
                filename, filepath, _ = latest_file
                if not database.is_file_loaded(DB_PATH, filename):
                    logger.info(f"Auto-fetching new report: {filename}")
                    success, count = database.load_excel(DB_PATH, filepath)
                    if success:
                        database.mark_file_loaded(DB_PATH, filename, filepath, count, source='network_share')
                        logger.info(f"Auto-fetch loaded {count} records from {filename}")
                else:
                    logger.info(f"Latest file {filename} already loaded.")
            else:
                logger.info(f"No reports found in {NETWORK_SHARE_PATH} during auto-fetch.")
            
        except Exception as e:
            logger.error(f"Auto-fetch error: {e}")
            
        # Sleep for 2 hours before checking the folder again
        time.sleep(2 * 3600)

def start_scheduler():
    t = threading.Thread(target=auto_fetch_reports, daemon=True)
    t.start()

check_initial_load()
start_scheduler()


# ─── Page Routes ──────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


# ─── API Routes ───────────────────────────────────────────────────────────────

@app.route('/api/dashboard', methods=['GET'])
def dashboard():
    fleet = request.args.get('fleet')
    date_from = request.args.get('date_from')
    date_to = request.args.get('date_to')
    bom = request.args.get('bom')
    part_group = request.args.get('part_group')
    
    try:
        stats = database.get_dashboard_stats(DB_PATH, fleet=fleet, date_from=date_from, date_to=date_to, bom=bom, part_group=part_group)
        return jsonify(stats)
    except Exception as e:
        logger.error(f"Dashboard error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/api/events', methods=['GET'])
def events():
    fleet = request.args.get('fleet')
    event_type = request.args.get('event_type')
    date_from = request.args.get('date_from')
    date_to = request.args.get('date_to')
    search = request.args.get('search')
    bom = request.args.get('bom')
    part_group = request.args.get('part_group')

    try:
        page = int(request.args.get('page', 1))
        per_page = int(request.args.get('per_page', 50))
    except ValueError:
        return jsonify({"error": "Invalid page or per_page parameter"}), 400

    try:
        results = database.get_events(
            DB_PATH, fleet=fleet, event_type=event_type,
            date_from=date_from, date_to=date_to,
            search=search, bom=bom, part_group=part_group,
            page=page, per_page=per_page
        )
        # Frontend expects { events, total, page, per_page }
        return jsonify({
            "events": results["data"],
            "total": results["total"],
            "page": results["page"],
            "per_page": results["per_page"]
        })
    except Exception as e:
        logger.error(f"Events error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/api/alerts/install-remove', methods=['GET'])
def alert_a():
    fleet = request.args.get('fleet')
    date_from = request.args.get('date_from')
    date_to = request.args.get('date_to')
    bom = request.args.get('bom')
    part_group = request.args.get('part_group')

    try:
        rows = database.get_alert_a_events(DB_PATH, fleet=fleet, date_from=date_from, date_to=date_to, bom=bom, part_group=part_group)
        return jsonify({"events": rows, "total": len(rows)})
    except Exception as e:
        logger.error(f"Alert A error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/api/alerts/xxx-sn', methods=['GET'])
def alert_b():
    fleet = request.args.get('fleet')
    date_from = request.args.get('date_from')
    date_to = request.args.get('date_to')
    bom = request.args.get('bom')
    part_group = request.args.get('part_group')

    try:
        rows = database.get_alert_b_events(DB_PATH, fleet=fleet, date_from=date_from, date_to=date_to, bom=bom, part_group=part_group)
        return jsonify({"alerts": rows, "total": len(rows)})
    except Exception as e:
        logger.error(f"Alert B error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/api/alerts/mmc', methods=['GET'])
def alert_mmc():
    """Alert C — Missing Mandatory Component.
    Flags REMOVE events where no reinstall occurred within 7 days."""
    fleet = request.args.get('fleet')
    date_from = request.args.get('date_from')
    date_to = request.args.get('date_to')
    bom = request.args.get('bom')
    part_group = request.args.get('part_group')

    try:
        alerts = database.get_mmc_alerts(DB_PATH, fleet=fleet, date_from=date_from, date_to=date_to, bom=bom, part_group=part_group)
        critical = sum(1 for a in alerts if a.get('mmc_severity') == 'CRITICAL')
        warning = sum(1 for a in alerts if a.get('mmc_severity') == 'WARNING')
        return jsonify({
            "alerts": alerts,
            "total": len(alerts),
            "critical_count": critical,
            "warning_count": warning
        })
    except Exception as e:
        logger.error(f"MMC alert error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/api/dashboard-by-fleet', methods=['GET'])
def dashboard_by_fleet():
    """XXX S/N Alert and Empty Config Slot counts broken down per fleet type,
    for the Dashboard tab's per-fleet KPI cards."""
    date_from = request.args.get('date_from')
    date_to = request.args.get('date_to')
    bom = request.args.get('bom')
    part_group = request.args.get('part_group')

    try:
        breakdown = database.get_fleet_dashboard(DB_PATH, date_from=date_from, date_to=date_to, bom=bom, part_group=part_group)
        return jsonify({"fleets": breakdown})
    except Exception as e:
        logger.error(f"Dashboard by fleet error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/api/fleet-types', methods=['GET'])
def fleet_types():
    try:
        types = database.get_fleet_types(DB_PATH)
        return jsonify({"fleets": types})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/bom-types', methods=['GET'])
def bom_types():
    try:
        types = database.get_bom_types(DB_PATH)
        return jsonify({"boms": types})
    except Exception as e:
        logger.error(f"BOM types error: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/api/part-groups', methods=['GET'])
def part_groups():
    try:
        groups = database.get_part_groups(DB_PATH)
        return jsonify({"groups": groups})
    except Exception as e:
        logger.error(f"Part groups error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/api/export', methods=['GET'])
def export_csv():
    fleet = request.args.get('fleet')
    event_type = request.args.get('event_type')
    date_from = request.args.get('date_from')
    date_to = request.args.get('date_to')
    search = request.args.get('search')
    bom = request.args.get('bom')
    part_group = request.args.get('part_group')

    try:
        results = database.get_events(
            DB_PATH, fleet=fleet, event_type=event_type,
            date_from=date_from, date_to=date_to,
            search=search, bom=bom, part_group=part_group,
            page=1, per_page=1000000
        )

        data = results['data']

        if not data:
            return "No data found", 404

        si = StringIO()
        cw = csv.DictWriter(si, fieldnames=data[0].keys())
        cw.writeheader()
        cw.writerows(data)

        output = si.getvalue()

        return Response(
            output,
            mimetype="text/csv",
            headers={"Content-disposition": "attachment; filename=htc_events_export.csv"}
        )
    except Exception as e:
        logger.error(f"Export error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/api/resolutions', methods=['GET'])
def get_resolutions():
    try:
        results = database.get_resolutions(DB_PATH)
        return jsonify({"resolutions": results})
    except Exception as e:
        logger.error(f"Resolutions error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/api/resolutions', methods=['POST'])
def create_resolution():
    data = request.json
    if not data:
        return jsonify({"error": "Invalid JSON", "success": False}), 400

    # Map frontend field names to database field names
    mapped = {
        "event_barcode": data.get("barcode") or data.get("event_barcode"),
        "alert_date": data.get("alert_date"),
        "aircraft": data.get("aircraft"),
        "config_slot": data.get("config_slot"),
        "part_no": data.get("part_no"),
        "original_sn": data.get("original_sn"),
        "resolved_sn": data.get("resolved_sn"),
        "engineer_responsible": data.get("engineer") or data.get("engineer_responsible"),
        "resolution_date": data.get("resolution_date"),
        "status": data.get("status", "PENDING"),
        "notes": data.get("notes"),
    }

    try:
        result = database.create_resolution(DB_PATH, mapped)
        return jsonify({"success": True, "resolution": result}), 201
    except Exception as e:
        logger.error(f"Create resolution error: {e}")
        return jsonify({"error": str(e), "success": False}), 500


@app.route('/api/resolutions/<int:res_id>', methods=['PUT'])
def update_resolution(res_id):
    data = request.json
    if not data:
        return jsonify({"error": "Invalid JSON", "success": False}), 400

    # Map frontend field names
    mapped = {}
    if "resolved_sn" in data:
        mapped["resolved_sn"] = data["resolved_sn"]
    if "engineer" in data or "engineer_responsible" in data:
        mapped["engineer_responsible"] = data.get("engineer") or data.get("engineer_responsible")
    if "resolution_date" in data:
        mapped["resolution_date"] = data["resolution_date"]
    if "status" in data:
        mapped["status"] = data["status"]
    if "notes" in data:
        mapped["notes"] = data["notes"]

    try:
        result = database.update_resolution(DB_PATH, res_id, mapped)
        if result:
            return jsonify({"success": True, "resolution": result})
        else:
            return jsonify({"error": "Resolution not found", "success": False}), 404
    except Exception as e:
        logger.error(f"Update resolution error: {e}")
        return jsonify({"error": str(e), "success": False}), 500


@app.route('/api/empty-slots', methods=['GET'])
def empty_slots():
    try:
        results = database.get_empty_slots(DB_PATH)
        return jsonify({"slots": results, "total": len(results)})
    except Exception as e:
        logger.error(f"Empty slots error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/api/fetch-reports', methods=['POST'])
def fetch_reports():
    """Scan the network share \\\\svhqftp03\\HCT for new bi-weekly .xlsx report files
    and load any that haven't been imported yet."""
    try:
        files = database.scan_share_directory(NETWORK_SHARE_PATH)
        if not files:
            return jsonify({
                "success": True,
                "share_path": NETWORK_SHARE_PATH,
                "files_loaded": 0,
                "total_records": 0,
                "message": f"No .xlsx or .xls files found at {NETWORK_SHARE_PATH}. Ensure the network share is accessible."
            })

        files_loaded = 0
        total_records = 0
        loaded_details = []

        for filename, filepath, parsed_timestamp in files:
            if database.is_file_loaded(DB_PATH, filename):
                continue  # Already imported

            logger.info(f"Loading new report: {filename} from {filepath}")
            success, count = database.load_excel(DB_PATH, filepath)
            if success:
                database.mark_file_loaded(DB_PATH, filename, filepath, count, source='network_share')
                files_loaded += 1
                total_records += count
                loaded_details.append({
                    "filename": filename, 
                    "records": count,
                    "timestamp": parsed_timestamp.isoformat() if parsed_timestamp else None
                })
                logger.info(f"Loaded {count} records from {filename}")
            else:
                logger.warning(f"Failed to load {filename}")

        return jsonify({
            "success": True,
            "share_path": NETWORK_SHARE_PATH,
            "files_loaded": files_loaded,
            "total_records": total_records,
            "files": loaded_details,
            "message": f"Loaded {files_loaded} file(s) with {total_records} total records."
        })
    except Exception as e:
        logger.error(f"Fetch reports error: {e}")
        return jsonify({"error": str(e), "success": False}), 500

@app.route('/api/scheduler-status', methods=['GET'])
def scheduler_status():
    last_fetch = database.get_last_fetch_time(DB_PATH)
    next_fetch = None
    if last_fetch:
        # Note: the background thread now loops every 2 hours,
        # so we approximate the next fetch by adding 2 hours to the last fetch time
        next_fetch = last_fetch + timedelta(hours=2)
    
    return jsonify({
        "auto_fetch_enabled": True,
        "interval_hours": 2,
        "last_fetch": last_fetch.isoformat() if last_fetch else None,
        "next_fetch": next_fetch.isoformat() if next_fetch else None
    })


@app.route('/api/loaded-files', methods=['GET'])
def loaded_files():
    """Return history of all imported files."""
    try:
        files = database.get_loaded_files(DB_PATH)
        return jsonify({"files": files})
    except Exception as e:
        logger.error(f"Loaded files error: {e}")
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    print('=' * 60)
    print('  HTC End-to-End Visibility Monitoring Dashboard')
    print('  Ethiopian Airlines | MPTC Engineering')
    print('  Running at http://localhost:5000')
    print('=' * 60)
    serve(app, host='0.0.0.0', port=5000)