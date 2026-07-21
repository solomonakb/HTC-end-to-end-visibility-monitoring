import sqlite3
import pandas as pd

conn = sqlite3.connect('htc_monitor.db')
df = pd.read_sql_query("SELECT assembly_cd, bom_class_cd, config_slot_code, config_slot_name FROM htc_events LIMIT 10", conn)
print(df)
print("\nUnique bom_class_cd:")
df2 = pd.read_sql_query("SELECT DISTINCT bom_class_cd FROM htc_events LIMIT 10", conn)
print(df2)
print("\nUnique config_slot_code:")
df3 = pd.read_sql_query("SELECT DISTINCT config_slot_code FROM htc_events LIMIT 10", conn)
print(df3)
conn.close()
