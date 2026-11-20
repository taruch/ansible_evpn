#!/usr/bin/python

DOCUMENTATION = '''
---
module: xls_set_row
short_description: Create or update a row in an .xlsx sheet, keyed by a column value
description:
    - Opens the given .xlsx workbook with openpyxl, finds the row in the given
      sheet whose C(key_column) matches C(key_value), and sets the columns
      given in C(values) on that row. If no row matches, a new row is
      appended with the key column populated plus C(values).
    - Column names are read from the sheet's header row (row 1).
    - Saves the workbook back to C(src) in place when a change is made.

requirements:
    - The openpyxl Python module must be installed on the Ansible host:
      pip install openpyxl

options:
    src:
        description: Path to the .xlsx workbook.
        required: true
        type: str
    sheet:
        description: Name of the sheet/tab to update.
        required: true
        type: str
    key_column:
        description: Header name of the column used to find the row.
        required: true
        type: str
    key_value:
        description: Value to match in key_column (or set, when appending a new row).
        required: true
        type: raw
    values:
        description: Mapping of column header to the value to set on the matched/new row.
        required: true
        type: dict

author: Todd Ruch
'''

EXAMPLES = '''
- name: Assign VLAN 1105 in the DataCenter tab
  xls_set_row:
    src: VLAN-Assignments-TEST-for-Ansible.xlsx
    sheet: DataCenter
    key_column: VLAN
    key_value: 1105
    values:
      Assignment: evpn-tenant3
      "IP Block": 172.16.105.0
      "IP Mask": 255.255.255.0
'''

RETURN = '''
row:
    description: The full set of column values on the matched/new row after the update.
    type: dict
    returned: always
'''

from ansible.module_utils.basic import AnsibleModule

import openpyxl


def normalize_number(value):
    """Ansible templates module args as strings by default (e.g. "{{ vlan_id | int }}"
    renders to "1105", not the int 1105), so coerce numeric-looking strings back to
    int/float to match openpyxl's native numeric cell values."""
    if isinstance(value, str):
        try:
            return int(value)
        except ValueError:
            try:
                return float(value)
            except ValueError:
                return value
    return value


def main():
    module = AnsibleModule(
        argument_spec=dict(
            src=dict(type='str', required=True),
            sheet=dict(type='str', required=True),
            key_column=dict(type='str', required=True),
            key_value=dict(type='raw', required=True),
            values=dict(type='dict', required=True),
        ),
        supports_check_mode=True,
    )

    src = module.params['src']
    sheet_name = module.params['sheet']
    key_column = module.params['key_column']
    key_value = normalize_number(module.params['key_value'])
    values = module.params['values']

    try:
        workbook = openpyxl.load_workbook(src)
    except (IOError, OSError) as e:
        module.fail_json(msg="Unable to open %s: %s" % (src, e))

    if sheet_name not in workbook.sheetnames:
        module.fail_json(
            msg="Sheet '%s' not found in %s. Available sheets: %s"
            % (sheet_name, src, workbook.sheetnames)
        )
    sheet = workbook[sheet_name]

    headers = {}
    for col in range(1, sheet.max_column + 1):
        header = sheet.cell(row=1, column=col).value
        if header is not None:
            headers[header] = col

    for column_name in [key_column] + list(values.keys()):
        if column_name not in headers:
            module.fail_json(
                msg="Column '%s' not found in sheet '%s'. Available columns: %s"
                % (column_name, sheet_name, list(headers.keys()))
            )

    key_col_index = headers[key_column]
    match_row = None
    for row in range(2, sheet.max_row + 1):
        if sheet.cell(row=row, column=key_col_index).value == key_value:
            match_row = row
            break

    changed = False
    if match_row is None:
        match_row = sheet.max_row + 1
        sheet.cell(row=match_row, column=key_col_index, value=key_value)
        changed = True

    for column_name, new_value in values.items():
        col_index = headers[column_name]
        if sheet.cell(row=match_row, column=col_index).value != new_value:
            sheet.cell(row=match_row, column=col_index, value=new_value)
            changed = True

    if changed and not module.check_mode:
        try:
            workbook.save(src)
        except (IOError, OSError) as e:
            module.fail_json(msg="Unable to save %s: %s" % (src, e))

    row_result = {
        header: sheet.cell(row=match_row, column=col_index).value
        for header, col_index in headers.items()
    }
    module.exit_json(changed=changed, row=row_result)


if __name__ == '__main__':
    main()
