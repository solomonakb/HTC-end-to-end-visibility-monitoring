import sqlite3

def get_events(db_path, bom=None):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    query = "SELECT COUNT(*) FROM htc_events WHERE 1=1"
    params = []
    
    if bom:
        query += " AND config_slot_code LIKE ?"
        params.append(f"%{bom}%")
        
    cursor.execute(query, params)
    count = cursor.fetchone()[0]
    conn.close()
    return count

print("Lowercase -htc count:", get_events('htc_monitor.db', bom='-htc'))
