import sqlite3
import openpyxl
import re
import os
import logging
from datetime import datetime

logger = logging.getLogger(__name__)

def init_db(db_path):
    """Initializes the database schema."""
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS htc_events (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        assembly_cd TEXT,
        nh_assembly_cd TEXT,
        bom_class_cd TEXT,
        config_slot_code TEXT,
        config_slot_name TEXT,
        part_group_cd TEXT,
        part_group_name TEXT,
        event_dt TEXT,
        event_type TEXT,
        status_cd TEXT,
        event_desc TEXT,
        aircraft TEXT,
        inventory_key TEXT,
        inventory TEXT,
        barcode TEXT,
        config_slot TEXT,
        part_no TEXT,
        part_desc TEXT,
        remove_reason TEXT,
        performed_by_user TEXT,
        performed_by_username TEXT,
        performed_by_hr_cd TEXT,
        serial_number TEXT,
        has_xxx_sn INTEGER,
        UNIQUE(barcode, event_dt, event_type, config_slot_code, inventory_key)
    )
    ''')
    
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS resolutions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        event_barcode TEXT,
        alert_date TEXT,
        aircraft TEXT,
        config_slot TEXT,
        part_no TEXT,
        original_sn TEXT,
        resolved_sn TEXT,
        engineer_responsible TEXT,
        resolution_date TEXT,
        status TEXT DEFAULT 'PENDING',
        notes TEXT,
        created_at TEXT,
        updated_at TEXT
    )
    ''')

    cursor.execute('''
    CREATE TABLE IF NOT EXISTS loaded_files (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        filename TEXT UNIQUE,
        filepath TEXT,
        records_loaded INTEGER DEFAULT 0,
        loaded_at TEXT,
        source TEXT DEFAULT 'manual'
    )
    ''')
    
    conn.commit()
    conn.close()


def is_file_loaded(db_path, filename):
    """Check if a file has already been imported."""
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM loaded_files WHERE filename = ?", (filename,))
    count = cursor.fetchone()[0]
    conn.close()
    return count > 0


def mark_file_loaded(db_path, filename, filepath, records_loaded, source='manual'):
    """Record that a file has been imported."""
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    now = datetime.utcnow().isoformat()
    try:
        cursor.execute('''
        INSERT OR REPLACE INTO loaded_files (filename, filepath, records_loaded, loaded_at, source)
        VALUES (?, ?, ?, ?, ?)
        ''', (filename, filepath, records_loaded, now, source))
        conn.commit()
    except Exception as e:
        logger.error(f"Error marking file as loaded: {e}")
    finally:
        conn.close()


def parse_filename_timestamp(filename, filepath):
    """Parse timestamps from filenames."""
    # Pattern 1: MON-DD-YYYY (e.g., JUN-17-2026)
    match = re.search(r'([A-Za-z]{3})-(\d{2})-(\d{4})', filename)
    if match:
        try:
            return datetime.strptime(match.group(0).upper(), '%b-%d-%Y')
        except ValueError:
            pass
            
    # Pattern 2: YYYYMMDDHHMI (e.g., 202607200837)
    match = re.search(r'(\d{12})', filename)
    if match:
        try:
            return datetime.strptime(match.group(1), '%Y%m%d%H%M')
        except ValueError:
            pass
            
    # Pattern 3: YYYY-MM-DD (e.g., 2026-07-17)
    match = re.search(r'(\d{4}-\d{2}-\d{2})', filename)
    if match:
        try:
            return datetime.strptime(match.group(1), '%Y-%m-%d')
        except ValueError:
            pass
            
    # Fallback
    try:
        return datetime.fromtimestamp(os.path.getmtime(filepath))
    except Exception:
        return datetime.min

def scan_share_directory(share_path):
    """Scan a network share or local directory for .xlsx and .xls files.
    Returns a list of (filename, full_path, parsed_timestamp) tuples, sorted latest first.
    """
    results = []
    try:
        if os.path.isdir(share_path):
            for entry in os.listdir(share_path):
                lower_entry = entry.lower()
                if (lower_entry.endswith('.xlsx') or lower_entry.endswith('.xls')) and not entry.startswith('~$'):
                    filepath = os.path.join(share_path, entry)
                    dt = parse_filename_timestamp(entry, filepath)
                    results.append((entry, filepath, dt))
            # Sort by timestamp, latest first
            results.sort(key=lambda x: x[2], reverse=True)
    except PermissionError:
        logger.error(f"Permission denied accessing: {share_path}")
    except FileNotFoundError:
        logger.error(f"Share path not found: {share_path}")
    except Exception as e:
        logger.error(f"Error scanning share: {e}")
    return results

def get_loaded_files(db_path):
    """Return list of all loaded file records."""
    conn = sqlite3.connect(db_path)
    conn.row_factory = _dict_factory
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM loaded_files ORDER BY loaded_at DESC")
    rows = cursor.fetchall()
    conn.close()
    return rows

def get_last_fetch_time(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT MAX(loaded_at) FROM loaded_files WHERE source = 'network_share'")
    result = cursor.fetchone()
    conn.close()
    if result and result[0]:
        try:
            return datetime.fromisoformat(result[0])
        except ValueError:
            return None
    return None

def _extract_sn(inventory_str):
    """Extracts serial number from inventory string and determines if it is XXX-like."""
    if not inventory_str:
        return "", 1
    
    # Example: 'VALVE (PN: 3216000-00, SN: 3216000-00311)'
    sn_match = re.search(r'SN:\s*([^)]+)', inventory_str)
    sn = sn_match.group(1).strip() if sn_match else ""
    
    xxx_values = ['XXX', 'UNKNOWN', 'N/A', 'NONE', '']
    has_xxx_sn = 1 if sn.upper() in xxx_values else 0
    
    return sn, has_xxx_sn

def load_excel(db_path, excel_path):
    """Loads data from the given Excel file into the database."""
    logger.info(f"Loading data from {excel_path} into {db_path}")
    
    rows_iter = None
    if excel_path.lower().endswith('.xls'):
        try:
            import xlrd
            workbook = xlrd.open_workbook(excel_path)
            sheet_name = 'Export Worksheet' if 'Export Worksheet' in workbook.sheet_names() else workbook.sheet_names()[0]
            sheet = workbook.sheet_by_name(sheet_name)
            
            def xlrd_rows():
                for row_idx in range(sheet.nrows):
                    row_vals = []
                    for col_idx in range(sheet.ncols):
                        cell_type = sheet.cell_type(row_idx, col_idx)
                        cell_value = sheet.cell_value(row_idx, col_idx)
                        if cell_type == xlrd.XL_CELL_DATE:
                            try:
                                dt = xlrd.xldate_as_datetime(cell_value, workbook.datemode)
                                row_vals.append(dt)
                            except Exception:
                                row_vals.append(cell_value)
                        else:
                            row_vals.append(cell_value)
                    yield row_vals
            
            rows_iter = xlrd_rows()
            logger.info(f"Using sheet: {sheet_name} (xlrd)")
        except ImportError:
            logger.error("xlrd library is not installed for reading .xls files.")
            return False, 0
        except Exception as e:
            logger.error(f"Failed to open .xls file: {e}")
            return False, 0
    else:
        try:
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            # Try 'Export Worksheet' first, fall back to first sheet
            if 'Export Worksheet' in wb.sheetnames:
                sheet = wb['Export Worksheet']
            else:
                sheet = wb[wb.sheetnames[0]]
                logger.info(f"Using sheet: {wb.sheetnames[0]}")
                
            def openpyxl_rows():
                for row in sheet.rows:
                    yield [cell.value for cell in row]
            
            rows_iter = openpyxl_rows()
        except Exception as e:
            logger.error(f"Failed to open Excel file: {e}")
            return False, 0
    
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Header check
    try:
        header = next(rows_iter)
    except StopIteration:
        logger.error("Empty sheet")
        return False, 0
        
    inserted_count = 0
    for values in rows_iter:
        # Skip empty rows completely
        if not any(values):
            continue
        # Skip header-like rows (e.g. from .xls files where header wasn't consumed)
        if values[0] and str(values[0]).strip().upper() == 'ASSMBL_CD':
            continue
            
        def safe_str(val):
            return str(val) if val is not None else ""
            
        assembly_cd = safe_str(values[0]) if len(values) > 0 else ""
        nh_assembly_cd = safe_str(values[1]) if len(values) > 1 else ""
        bom_class_cd = safe_str(values[2]) if len(values) > 2 else ""
        config_slot_code = safe_str(values[3]) if len(values) > 3 else ""
        config_slot_name = safe_str(values[4]) if len(values) > 4 else ""
        part_group_cd = safe_str(values[5]) if len(values) > 5 else ""
        part_group_name = safe_str(values[6]) if len(values) > 6 else ""
        
        event_dt_raw = values[7] if len(values) > 7 else None
        event_dt = ""
        if isinstance(event_dt_raw, datetime):
            event_dt = event_dt_raw.isoformat()
        elif event_dt_raw:
            event_dt = safe_str(event_dt_raw)
            
        event_type = safe_str(values[8]) if len(values) > 8 else ""
        status_cd = safe_str(values[9]) if len(values) > 9 else ""
        event_desc = safe_str(values[10]) if len(values) > 10 else ""
        aircraft = safe_str(values[11]) if len(values) > 11 else ""
        inventory_key = safe_str(values[12]) if len(values) > 12 else ""
        inventory = safe_str(values[13]) if len(values) > 13 else ""
        barcode = safe_str(values[14]) if len(values) > 14 else ""
        config_slot = safe_str(values[15]) if len(values) > 15 else ""
        part_no = safe_str(values[16]) if len(values) > 16 else ""
        part_desc = safe_str(values[17]) if len(values) > 17 else ""
        remove_reason = safe_str(values[18]) if len(values) > 18 else ""
        performed_by_user = safe_str(values[19]) if len(values) > 19 else ""
        performed_by_username = safe_str(values[20]) if len(values) > 20 else ""
        performed_by_hr_cd = safe_str(values[21]) if len(values) > 21 else ""
        
        sn, has_xxx_sn = _extract_sn(inventory)
        
        try:
            cursor.execute('''
            INSERT INTO htc_events (
                assembly_cd, nh_assembly_cd, bom_class_cd, config_slot_code, config_slot_name,
                part_group_cd, part_group_name, event_dt, event_type, status_cd,
                event_desc, aircraft, inventory_key, inventory, barcode,
                config_slot, part_no, part_desc, remove_reason, performed_by_user,
                performed_by_username, performed_by_hr_cd, serial_number, has_xxx_sn
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                assembly_cd, nh_assembly_cd, bom_class_cd, config_slot_code, config_slot_name,
                part_group_cd, part_group_name, event_dt, event_type, status_cd,
                event_desc, aircraft, inventory_key, inventory, barcode,
                config_slot, part_no, part_desc, remove_reason, performed_by_user,
                performed_by_username, performed_by_hr_cd, sn, has_xxx_sn
            ))
            inserted_count += 1
        except sqlite3.IntegrityError:
            # Skip duplicates based on unique constraint
            pass
            
    conn.commit()
    conn.close()
    return True, inserted_count

def _dict_factory(cursor, row):
    d = {}
    for idx, col in enumerate(cursor.description):
        d[col[0]] = row[idx]
    return d

def get_events(db_path, fleet=None, event_type=None, date_from=None, date_to=None, search=None, bom=None, part_group=None, page=1, per_page=50):
    conn = sqlite3.connect(db_path)
    conn.row_factory = _dict_factory
    cursor = conn.cursor()
    
    query = "SELECT * FROM htc_events WHERE 1=1"
    params = []
    
    if fleet:
        query += " AND assembly_cd = ?"
        params.append(fleet)
    if event_type:
        query += " AND event_type = ?"
        params.append(event_type)
    if date_from:
        query += " AND event_dt >= ?"
        params.append(date_from)
    if date_to:
        query += " AND event_dt <= ?"
        params.append(date_to)
    if bom:
        query += " AND config_slot_code LIKE ?"
        params.append(f"%{bom}%")
    if part_group:
        query += " AND part_group_name LIKE ?"
        params.append(f"%{part_group}%")
    if search:
        query += " AND (barcode LIKE ? OR part_no LIKE ? OR serial_number LIKE ?)"
        params.extend([f"%{search}%", f"%{search}%", f"%{search}%"])
        
    # Count total
    count_query = f"SELECT COUNT(*) as total FROM ({query})"
    cursor.execute(count_query, params)
    total = cursor.fetchone()['total']
    
    # Pagination
    offset = (page - 1) * per_page
    query += " ORDER BY id DESC LIMIT ? OFFSET ?"
    params.extend([per_page, offset])
    
    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()
    
    return {
        "data": rows,
        "total": total,
        "page": page,
        "per_page": per_page
    }

def get_dashboard_stats(db_path, fleet=None, date_from=None, date_to=None, bom=None, part_group=None):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    stats = {
        "total_events": 0,
        "install_count": 0,
        "remove_count": 0,
        "xxx_sn_count": 0,
        "empty_slots_count": 0,
        "fleet_breakdown": {}
    }
    
    # Base query logic for filters
    base_query = " FROM htc_events WHERE 1=1"
    params = []
    
    if fleet:
        base_query += " AND assembly_cd = ?"
        params.append(fleet)
    if date_from:
        base_query += " AND event_dt >= ?"
        params.append(date_from)
    if date_to:
        base_query += " AND event_dt <= ?"
        params.append(date_to)
    if bom:
        base_query += " AND config_slot_code LIKE ?"
        params.append(f"%{bom}%")
    if part_group:
        base_query += " AND part_group_name LIKE ?"
        params.append(f"%{part_group}%")
        
    cursor.execute("SELECT COUNT(*)" + base_query, params)
    stats["total_events"] = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(*)" + base_query + " AND event_type = 'INSTALL'", params)
    stats["install_count"] = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(*)" + base_query + " AND event_type = 'REMOVE'", params)
    stats["remove_count"] = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(*)" + base_query + " AND has_xxx_sn = 1", params)
    stats["xxx_sn_count"] = cursor.fetchone()[0]
    
    # Calculate Empty Slots (MMC) based on the actual logic with filters
    mmc_alerts = get_mmc_alerts(db_path, fleet=fleet, date_from=date_from, date_to=date_to, bom=bom, part_group=part_group)
    stats["empty_slots_count"] = len(mmc_alerts)
    
    cursor.execute("SELECT assembly_cd, COUNT(*) FROM htc_events GROUP BY assembly_cd")
    for row in cursor.fetchall():
        if row[0]:
            stats["fleet_breakdown"][row[0]] = row[1]
            
    conn.close()
    return stats

def get_alert_a_events(db_path, fleet=None, date_from=None, date_to=None, bom=None, part_group=None):
    conn = sqlite3.connect(db_path)
    conn.row_factory = _dict_factory
    cursor = conn.cursor()
    
    query = "SELECT * FROM htc_events WHERE event_type IN ('INSTALL', 'REMOVE')"
    params = []
    
    if fleet:
        query += " AND assembly_cd = ?"
        params.append(fleet)
    if date_from:
        query += " AND event_dt >= ?"
        params.append(date_from)
    if date_to:
        query += " AND event_dt <= ?"
        params.append(date_to)
    if bom:
        query += " AND config_slot_code = ?"
        params.append(bom)
    if part_group:
        query += " AND part_group_name = ?"
        params.append(part_group)
        
    query += " ORDER BY event_dt DESC"
    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()
    return rows

def get_alert_b_events(db_path, fleet=None, date_from=None, date_to=None, bom=None, part_group=None):
    conn = sqlite3.connect(db_path)
    conn.row_factory = _dict_factory
    cursor = conn.cursor()
    
    query = "SELECT * FROM htc_events WHERE has_xxx_sn = 1"
    params = []
    
    if fleet:
        query += " AND assembly_cd = ?"
        params.append(fleet)
    if date_from:
        query += " AND event_dt >= ?"
        params.append(date_from)
    if date_to:
        query += " AND event_dt <= ?"
        params.append(date_to)
    if bom:
        query += " AND config_slot_code LIKE ?"
        params.append(f"%{bom}%")
    if part_group:
        query += " AND part_group_name LIKE ?"
        params.append(f"%{part_group}%")
        
    query += " ORDER BY id DESC"
    
    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()
    return rows


def _parse_event_date(dt_str):
    """Parse event date strings in multiple formats. Returns datetime or None."""
    if not dt_str:
        return None
    # Try DD-MON-YYYY HH:MI:SS (e.g. '02-JUL-2026 15:54:34')
    try:
        return datetime.strptime(dt_str, '%d-%b-%Y %H:%M:%S')
    except (ValueError, TypeError):
        pass
    # Try ISO format (e.g. '2026-07-02T15:54:34')
    try:
        return datetime.fromisoformat(dt_str)
    except (ValueError, TypeError):
        pass
    # Try DD-MON-YYYY (no time)
    try:
        return datetime.strptime(dt_str, '%d-%b-%Y')
    except (ValueError, TypeError):
        pass
    return None


def _get_inv_key_num(inv_key_str):
    if not inv_key_str:
        return 0
    try:
        parts = inv_key_str.split(':')
        if len(parts) > 1:
            return int(parts[1])
        return int(parts[0])
    except:
        return 0

def get_mmc_alerts(db_path, fleet=None, date_from=None, date_to=None, bom=None, part_group=None):
    """Detect Missing Mandatory Component (MMC) alerts.
    
    A removed part must be reinstalled on the same aircraft + config slot
    within 7 days. If no matching INSTALL is found within that window,
    the removal is flagged as MMC.
    """
    conn = sqlite3.connect(db_path)
    conn.row_factory = _dict_factory
    cursor = conn.cursor()
    
    # Get all REMOVE events
    remove_query = "SELECT * FROM htc_events WHERE event_type = 'REMOVE'"
    params = []
    if fleet:
        remove_query += " AND assembly_cd = ?"
        params.append(fleet)
    if date_from:
        remove_query += " AND event_dt >= ?"
        params.append(date_from)
    if date_to:
        remove_query += " AND event_dt <= ?"
        params.append(date_to)
    if bom:
        remove_query += " AND config_slot_code LIKE ?"
        params.append(f"%{bom}%")
    if part_group:
        remove_query += " AND part_group_name LIKE ?"
        params.append(f"%{part_group}%")
    remove_query += " ORDER BY id DESC"
    
    cursor.execute(remove_query, params)
    removals_raw = cursor.fetchall()
    
    # Get all INSTALL events for matching
    cursor.execute("SELECT * FROM htc_events WHERE event_type = 'INSTALL'")
    installs_raw = cursor.fetchall()
    conn.close()
    
    # Parse dates and filter valid ones
    removals = []
    for r in removals_raw:
        dt = _parse_event_date(r.get('event_dt'))
        if dt:
            removals.append((dt, r))
            
    installs = []
    for i in installs_raw:
        dt = _parse_event_date(i.get('event_dt'))
        if dt:
            installs.append((dt, i))
            
    # Sort chronologically (by date, then FG_INVENTORY_KEY as tie-breaker)
    removals.sort(key=lambda x: (x[0], _get_inv_key_num(x[1].get('inventory_key', ''))))
    installs.sort(key=lambda x: (x[0], _get_inv_key_num(x[1].get('inventory_key', ''))))
    
    now = datetime.now()
    mmc_alerts = []
    claimed_installs = set()
    
    for removal_dt, removal in removals:
        aircraft = removal.get('aircraft', '')
        config_slot = removal.get('config_slot_code', '')
        rem_inv_num = _get_inv_key_num(removal.get('inventory_key', ''))
        
        # Find the earliest INSTALL on the same aircraft + config_slot that hasn't been claimed
        matching_install = None
        for inst_dt, inst in installs:
            inst_id = inst.get('id')
            if inst_id in claimed_installs:
                continue
                
            if (inst.get('aircraft', '') == aircraft and 
                inst.get('config_slot_code', '') == config_slot):
                
                inst_inv_num = _get_inv_key_num(inst.get('inventory_key', ''))
                
                # Match if install is strictly after in time
                # OR if time is identical, the install's inventory_key is GREATER THAN OR EQUAL to removal's
                if inst_dt > removal_dt or (inst_dt == removal_dt and inst_inv_num >= rem_inv_num):
                    matching_install = inst
                    claimed_installs.add(inst_id)
                    break  # earliest matching install
        
        days_since = (now - removal_dt).days
        
        if matching_install:
            # There is a reinstall — check if it was within 7 days
            inst_dt = _parse_event_date(matching_install.get('event_dt'))
            days_to_reinstall = (inst_dt - removal_dt).days if inst_dt else 0
            if days_to_reinstall <= 7:
                continue  # Reinstalled on time — no alert
            
            # Reinstalled late
            alert = dict(removal)
            alert['days_since_removal'] = days_since
            alert['days_to_reinstall'] = days_to_reinstall
            alert['mmc_severity'] = 'CRITICAL'
            alert['reinstall_status'] = 'REINSTALLED_LATE'
            alert['reinstall_date'] = matching_install.get('event_dt', '')
            alert['reinstall_sn'] = matching_install.get('serial_number', '')
            alert['_parsed_dt'] = removal_dt
            mmc_alerts.append(alert)
        else:
            # No reinstall found at all
            if days_since >= 3:
                alert = dict(removal)
                alert['days_since_removal'] = days_since
                alert['days_to_reinstall'] = None
                alert['reinstall_date'] = None
                alert['reinstall_sn'] = None
                alert['_parsed_dt'] = removal_dt
                if days_since >= 7:
                    alert['mmc_severity'] = 'CRITICAL'
                    alert['reinstall_status'] = 'OVERDUE'
                else:
                    alert['mmc_severity'] = 'WARNING'
                    alert['reinstall_status'] = 'PENDING'
                mmc_alerts.append(alert)
                
    # Sort newest first for the UI
    mmc_alerts.sort(key=lambda x: x['_parsed_dt'], reverse=True)
    for a in mmc_alerts:
        del a['_parsed_dt']
    
    return mmc_alerts

def get_fleet_types(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT assembly_cd FROM htc_events WHERE assembly_cd IS NOT NULL AND assembly_cd != ''")
    rows = cursor.fetchall()
    conn.close()
    return [r[0] for r in rows]

def get_empty_slots(db_path):
    conn = sqlite3.connect(db_path)
    conn.row_factory = _dict_factory
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM htc_events WHERE event_dt IS NULL OR event_dt = '' ORDER BY id DESC")
    rows = cursor.fetchall()
    conn.close()
    return rows

def get_resolutions(db_path):
    conn = sqlite3.connect(db_path)
    conn.row_factory = _dict_factory
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM resolutions ORDER BY id DESC")
    rows = cursor.fetchall()
    conn.close()
    return rows

def create_resolution(db_path, data):
    conn = sqlite3.connect(db_path)
    conn.row_factory = _dict_factory
    cursor = conn.cursor()
    
    now = datetime.utcnow().isoformat()
    cursor.execute('''
    INSERT INTO resolutions (
        event_barcode, alert_date, aircraft, config_slot, part_no,
        original_sn, resolved_sn, engineer_responsible, resolution_date,
        status, notes, created_at, updated_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        data.get('event_barcode'),
        data.get('alert_date'),
        data.get('aircraft'),
        data.get('config_slot'),
        data.get('part_no'),
        data.get('original_sn'),
        data.get('resolved_sn'),
        data.get('engineer_responsible'),
        data.get('resolution_date'),
        data.get('status', 'PENDING'),
        data.get('notes'),
        now,
        now
    ))
    conn.commit()
    res_id = cursor.lastrowid
    
    cursor.execute("SELECT * FROM resolutions WHERE id = ?", (res_id,))
    row = cursor.fetchone()
    conn.close()
    return row

def update_resolution(db_path, res_id, data):
    conn = sqlite3.connect(db_path)
    conn.row_factory = _dict_factory
    cursor = conn.cursor()
    
    now = datetime.utcnow().isoformat()
    
    set_clauses = []
    params = []
    
    for key, value in data.items():
        if key in ['resolved_sn', 'engineer_responsible', 'resolution_date', 'status', 'notes']:
            set_clauses.append(f"{key} = ?")
            params.append(value)
            
    if set_clauses:
        set_clauses.append("updated_at = ?")
        params.append(now)
        
        query = f"UPDATE resolutions SET {', '.join(set_clauses)} WHERE id = ?"
        params.append(res_id)
        
        cursor.execute(query, params)
        conn.commit()
        
    cursor.execute("SELECT * FROM resolutions WHERE id = ?", (res_id,))
    row = cursor.fetchone()
    conn.close()
    return row

def get_bom_types(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT config_slot_code FROM htc_events WHERE config_slot_code IS NOT NULL AND config_slot_code != '' ORDER BY config_slot_code")
    boms = [row[0] for row in cursor.fetchall()]
    conn.close()
    return boms

def get_part_groups(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT part_group_name FROM htc_events WHERE part_group_name IS NOT NULL AND part_group_name != '' ORDER BY part_group_name")
    groups = [row[0] for row in cursor.fetchall()]
    conn.close()
    return groups
