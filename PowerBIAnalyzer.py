import json
import os
import sys
import tempfile
import zipfile
import shutil
import re
from tabulate import tabulate

# Function to process the template file (copy to temp location, rename to .zip, unzip, delete at the end) and get the contents of the DataModelSchema and Layout files
def get_file_contents(template_path):
    temp_folder_path = tempfile.mkdtemp()
    temp_zip_path = os.path.join(temp_folder_path, "temp.zip")

    # Copy the template file to the temporary directory and unzip it
    shutil.copy(template_path, temp_zip_path)
    with zipfile.ZipFile(temp_zip_path, 'r') as zip_ref:
        zip_ref.extractall(temp_folder_path)
    # Get the contents of the DataModelSchema and Layout files
    with open(os.path.join(temp_folder_path, "DataModelSchema"), 'r', encoding='utf-8') as file:
        content = file.read().replace("\x00", "")
        data_model = json.loads(content)
    with open(os.path.join(temp_folder_path, "Report", "Layout"), 'r', encoding='utf-8') as file:
        content = file.read().replace("\x00", "")
        layout = json.loads(content)

    # Delete the temporary directory
    shutil.rmtree(temp_folder_path)

    return data_model, layout

# Get report details into 2 files, make sure they are treated as objects
template_path = sys.argv[1].strip().strip('"')
data_model, layout = get_file_contents(template_path)

# Create variables to store the list of columns and measures used in the report, as well as the various expressions from the files
field_list = []
all_config_details = ""

# Iterate through all tables. Only hidden tables will have an isHidden property, so filtering by its existence is acceptable
for table in data_model['model']['tables']:
    if 'isHidden' not in table or not table['isHidden']:
        # Iterate through all columns and add them to the fieldList collection
        for column in table['columns']:
            if 'isHidden' not in column or not column['isHidden']:
                field_type = column.get('type', 'column')
                expression = column.get('expression', '') if field_type == 'calculated' else ''

                field_list.append({
                    'table': table['name'],
                    'field': column['name'],
                    'type': field_type,
                    'used': False,
                    'expression': expression
                })

        # Iterate through all measures and add them to the fieldList collection
        for measure in table.get('measures', []):
            field_list.append({
                'table': table['name'],
                'field': measure['name'],
                'type': 'measure',
                'used': False,
                'expression': measure['expression']
            })

# Iterate through all visuals and add the configuration details into the allConfigDetails variable
for section in layout['sections']:
    for visual_container in section['visualContainers']:
        config = json.loads(visual_container['config'])
        all_config_details += " " + json.dumps(config['singleVisual'], indent=None)


# Iterate through all fields and mark them as used if they appear anywhere in the configs
for field in field_list:
    ### REGEX PATTERN TO MATCH FIELD IN CONFIGS
    # Table.field
    # 'Table'.field
    # FUNTION(Table.field)
    # FUNCTION('Table'.field)
    # "Property":  "field" <-- two spaces after the colon
    ###
    pattern = f"{re.escape(field['table'])}[']?.{re.escape(field['field'])}[)]?\"|\"Property\":  \"{re.escape(field['field'])}\""
    if re.search(pattern, all_config_details):
        field['used'] = True


# Recursively iterate through all unused fields to see if they are used in any DAX expressions of used fields
unused_fields_count = 0
while True: #unused_fields_count != len([field for field in field_list if not field['used']]):
    change_made = False
    expressions = " ".join([field['expression'] for field in field_list if field['used']])
    for field in [field for field in field_list if not field['used']]:
        ### REGEX PATTERN TO MATCH FIELD IN DAX EXPRESSIONS
        # Table[field]
        # 'Table'[field]
        ###
        pattern = f"{re.escape(field['table'])}[']?\[{re.escape(field['field'])}\]"
        if re.search(pattern, expressions):
            field['used'] = True
            change_made = True
    #unused_fields = [field for field in field_list if not field['used']]
    if (not change_made):
        break

# Print out all unused fields
unused_fields = [field for field in field_list if not field['used']]
table = [[field['table'], field['field'], field['type'], field['used']] for field in unused_fields]

print(tabulate(table, headers=["Table", "Field", "Type", "Used"]))