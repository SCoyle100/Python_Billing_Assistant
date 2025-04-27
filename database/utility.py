import sqlite3

# Connect to the database file (replace 'your_database.db' with your actual file)
conn = sqlite3.connect('invoice.db')
cursor = conn.cursor()

# Retrieve all table names
cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
tables = cursor.fetchall()
print("Tables:", tables)

# Read all rows from each table
for table in tables:
    table_name = table[0]
    print(f"\nContents of table {table_name}:")
    cursor.execute(f"SELECT * FROM {table_name}")
    rows = cursor.fetchall()
    for row in rows:
        print(row)

# Close the connection
conn.close()
