import sqlite3
import pandas as pd

conn = sqlite3.connect('htc_monitor.db')
df = pd.read_sql_query("SELECT DISTINCT nh_assembly_cd FROM htc_events LIMIT 10", conn)
print("nh_assembly_cd:", df)
df2 = pd.read_sql_query("SELECT config_slot_code, COUNT(*) FROM htc_events WHERE config_slot_code LIKE '%-HTC%' GROUP BY config_slot_code", conn)
print("config_slot_code with -HTC:", df2)
conn.close()
