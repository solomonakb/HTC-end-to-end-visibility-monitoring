import sqlite3
import pandas as pd

conn = sqlite3.connect('htc_monitor.db')
df = pd.read_sql_query("SELECT performed_by_user, performed_by_username, performed_by_hr_cd FROM htc_events LIMIT 10", conn)
print(df)
conn.close()
