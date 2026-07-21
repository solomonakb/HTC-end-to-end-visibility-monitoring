import sqlite3
import pandas as pd

conn = sqlite3.connect('htc_monitor.db')
c = conn.cursor()
c.execute("SELECT event_type, COUNT(*) FROM htc_events WHERE config_slot_code LIKE '%-HTC%' GROUP BY event_type")
print(c.fetchall())
conn.close()
