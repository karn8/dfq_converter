k_field_mapping = {
    'K1001': 'Part_Number',
    'K1002': 'PartTitle',
    'K1003': 'Aggregate',
    'K1005': 'Component',
    'K1008': 'Model',
    'K1086': 'Operation',
    'K1100': 'Plant',
    'K1102': 'Gaging-Station_Name',
    'K1103': 'Cost centre',
}

for k_field, column_name in k_field_mapping.items():
    print(f"{k_field}  {column_name} \n")
    # if k_field in loaded_sheet_1 and loaded_sheet_1[k_field].value is not None:
    #     output_file.write(f'{k_field}/1 {loaded_sheet_1[k_field].value}\n')