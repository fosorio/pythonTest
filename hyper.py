# Import the required modules
import cx_Oracle # For connecting to Oracle database
import tableauhyperapi as hyper # For creating and manipulating TableDefinition objects
import json # For serializing and deserializing TableDefinition objects

# Define the connection parameters for the Oracle database
oracle_user = "username" # Replace with your username
oracle_password = "password" # Replace with your password
oracle_host = "host" # Replace with your host
oracle_port = "port" # Replace with your port
oracle_sid = "sid" # Replace with your sid

# Define the table name and schema that resides in the Oracle database
table_name = "table_name" # Replace with your table name
table_schema = "table_schema" # Replace with your table schema

# Define the file name for the serialized TableDefinition object
file_name = "table_definition.json" # Replace with your file name

# Define a function to serialize a TableDefinition object to a JSON file
def serialize_table_definition(table_definition, file_name):
    # Convert the TableDefinition object to a dictionary
    table_definition_dict = {
        "table_name": str(table_definition.table_name),
        "persistence": str(table_definition.persistence),
        "columns": [
            {
                "name": str(column.name),
                "type": str(column.type),
                "nullability": str(column.nullability),
                "collation": str(column.collation)
            }
            for column in table_definition.columns
        ]
    }
    # Write the dictionary to a JSON file
    with open(file_name, "w") as file:
        json.dump(table_definition_dict, file, indent=4)

# Define a function to deserialize a TableDefinition object from a JSON file
def deserialize_table_definition(file_name):
    # Read the dictionary from the JSON file
    with open(file_name, "r") as file:
        table_definition_dict = json.load(file)
    # Convert the dictionary to a TableDefinition object
    table_definition = hyper.TableDefinition(
        table_name=hyper.TableName.parse(table_definition_dict["table_name"]),
        persistence=hyper.Persistence[table_definition_dict["persistence"]]
    )
    # Add the columns to the TableDefinition object
    for column in table_definition_dict["columns"]:
        table_definition.add_column(
            name=column["name"],
            type=hyper.SqlType.parse(column["type"]),
            nullability=hyper.Nullability[column["nullability"]],
            collation=column["collation"]
        )
    # Return the TableDefinition object
    return table_definition

# Connect to the Oracle database
oracle_connection = cx_Oracle.connect(
    user=oracle_user,
    password=oracle_password,
    dsn=cx_Oracle.makedsn(oracle_host, oracle_port, sid=oracle_sid)
)

# Create a cursor for executing queries
oracle_cursor = oracle_connection.cursor()

# Query the metadata of the table from the Oracle database
oracle_cursor.execute(f"SELECT COLUMN_NAME, DATA_TYPE, NULLABLE, DATA_LENGTH, DATA_PRECISION, DATA_SCALE FROM ALL_TAB_COLUMNS WHERE TABLE_NAME = '{table_name}' AND OWNER = '{table_schema}'")

# Create a TableDefinition object for the table
table_definition = hyper.TableDefinition(
    table_name=hyper.TableName(table_schema, table_name),
    persistence=hyper.Persistence.PERMANENT # Change this if needed
)

# Loop through the rows of the query result
for row in oracle_cursor:
    # Get the column name, data type, nullability, and collation from the row
    column_name = row[0]
    column_type = row[1]
    column_nullability = row[2]
    column_collation = None # Change this if needed
    # Convert the Oracle data type to the Hyper SQL type
    # This is a simplified mapping and may not cover all cases
    # Refer to the documentation of both Oracle and Hyper for more details
    # [^1^][5] [^2^][1]
    if column_type == "CHAR":
        column_type = hyper.SqlType.char(row[3])
    elif column_type == "VARCHAR2":
        column_type = hyper.SqlType.varchar(row[3])
    elif column_type == "NUMBER":
        column_type = hyper.SqlType.decimal(row[4], row[5])
    elif column_type == "DATE":
        column_type = hyper.SqlType.date()
    elif column_type == "TIMESTAMP":
        column_type = hyper.SqlType.timestamp()
    elif column_type == "CLOB":
        column_type = hyper.SqlType.clob()
    elif column_type == "BLOB":
        column_type = hyper.SqlType.blob()
    else:
        column_type = hyper.SqlType.text() # Use text as a default type
    # Convert the Oracle nullability to the Hyper nullability
    # This is a straightforward mapping
    if column_nullability == "Y":
        column_nullability = hyper.Nullability.NULLABLE
    else:
        column_nullability = hyper.Nullability.NOT_NULLABLE
    # Add the column to the TableDefinition object
    table_definition.add_column(
        name=column_name,
        type=column_type,
        nullability=column_nullability,
        collation=column_collation
    )

# Close the cursor and the connection
oracle_cursor.close()
oracle_connection.close()

# Serialize the TableDefinition object to a JSON file
serialize_table_definition(table_definition, file_name)

# Deserialize the TableDefinition object from the JSON file
# This is for testing purposes, you can comment it out if not needed
table_definition = deserialize_table_definition(file_name)

# Print the TableDefinition object
# This is for testing purposes, you can comment it out if not needed
print(table_definition)



# Import the required modules
import tableauhyperapi as hyper # For creating and manipulating TableDefinition objects
import csv # For reading data from CSV files

# Define the function that takes in the TableDefinition, a csv filename, and produces a hyper file of that csv file
def create_hyper_file_from_csv(table_definition, csv_file_name, hyper_file_name):
    # Create a HyperProcess object
    with hyper.HyperProcess(telemetry=hyper.Telemetry.DO_NOT_SEND_USAGE_DATA_TO_TABLEAU) as hyper_process:
        # Create a Connection object to the hyper file
        with hyper.Connection(endpoint=hyper_process.endpoint, database=hyper_file_name, create_mode=hyper.CreateMode.CREATE_AND_REPLACE) as connection:
            # Create the table in the hyper file using the TableDefinition object
            connection.catalog.create_table(table_definition=table_definition)
            # Open the CSV file for reading
            with open(csv_file_name, "r") as csv_file:
                # Create a CSV reader object
                csv_reader = csv.reader(csv_file)
                # Skip the header row of the CSV file
                next(csv_reader)
                # Loop through the rows of the CSV file
                for row in csv_reader:
                    # Convert the row to a list of values
                    values = [value for value in row]
                    # Insert the values into the table in the hyper file
                    connection.execute_command(command=f"INSERT INTO {table_definition.table_name} VALUES ({', '.join(values)})")
            # Print the number of rows in the table
            print(f"The table {table_definition.table_name} has {connection.execute_scalar_query(query=f'SELECT COUNT(*) FROM {table_definition.table_name}')}")
