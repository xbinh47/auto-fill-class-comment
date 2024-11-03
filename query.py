import json

# Function to read data from a JSON file
def read_data_from_json(filename):
    with open(filename, 'r', encoding='utf-8') as file:
        data = json.load(file)
    return data

# Function to generate updateOne query strings
def generate_update_queries(data):
    queries = []
    for entry in data:
        query = (
            f'db.employee_custom_field_setup.updateOne('
            f'{{ _id: ObjectId("{entry["_id"]}") }}, '
            f'{{ $set: {{ "ja.field_name": "{entry["ja"]["field_name"]}" }} }}'
            f')'
        )
        queries.append(query)
    return queries

# Write queries to a file
def write_queries_to_file(queries, filename):
    with open(filename, 'w', encoding='utf-8') as file:
        for query in queries:
            file.write(query + '\n')

# Read data from JSON file
data = read_data_from_json('data.json')

# Generate queries
queries = generate_update_queries(data)

# Write to file
write_queries_to_file(queries, 'update_queries.txt')

print("Queries have been written to update_queries.txt")