#!/usr/bin/env python

import os.path
import json
import argparse
import configparser
from openpyxl import load_workbook
from openpyxl.utils import coordinate_from_string, column_index_from_string

config_file = 'xlsx_inventory.cfg'


def main():
    args = parse_args()
    config = load_config()
    try:
        wb = load_workbook(config['xlsx_inventory_file'])
        if 'sheet' in config:
            sheet = wb[config['sheet']]
        else:
            sheet = wb.active
        inventory = sheet_to_inventory(group_by_col=config['group_by_col'], hostname_col=config['hostname_col'], sheet=sheet)
        if args.list:
            print(json.dumps(inventory, indent=4, sort_keys=True))
        elif args.host:
            try:
                print(json.dumps(
                    inventory['_meta']['hostvars'][args.host],
                    indent=4, sort_keys=True))
            except KeyError as e:
                print('\033[91mHost "%s" not Found!\033[0m' % e)
                print(e)
    except FileNotFoundError as e:
        print(
            '\033[91mFile Not Found! Check %s configuration file!'
            ' Is the `xlsx_inventory_file` path setting correct?\033[0m' % config_file)
        print(e)
        exit(1)
    except KeyError as e:
        print(
            '\033[91mKey Error! Check %s configuration file! Is the `sheet` name setting correct?\033[0m' % config_file)
        print(e)
        exit(1)
    exit(0)


def load_config():
    config = configparser.ConfigParser()
    config['DEFAULT'] = {'hostname_col': 'A', 'group_by_col': 'B', 'xlsx_inventory_file': 'inventory.xlsx'}
    config['xlsx_inventory'] = {}
    if os.path.isfile(config_file):
        config.read(config_file)
    else:
        with open(config_file, 'w')as cf:
            config.write(cf)
    return config['xlsx_inventory']


def parse_args():
    arg_parser = argparse.ArgumentParser(description='Excel Spreadsheet Inventory Module')
    group = arg_parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--list', action='store_true', help='List active servers')
    group.add_argument('--host', help='List details about the specified host', default=None)
    return arg_parser.parse_args()


def sheet_to_inventory(group_by_col, hostname_col, sheet):
    if type(group_by_col) is str:
        group_by_col = column_index_from_string(coordinate_from_string(group_by_col + '1')[0]) - 1
    if type(hostname_col) is str:
        hostname_col = column_index_from_string(coordinate_from_string(hostname_col + '1')[0]) - 1

    groups = {
        '_meta': {
            'hostvars': {}
        }
    }
    rows = list(sheet.rows)

    for row in rows[1:]:
        if row[group_by_col].value not in groups.keys():
            groups[row[group_by_col].value] = {
                'hosts': [],
                'vars': {}
            }
        groups[row[group_by_col].value]['hosts'].append(row[hostname_col].value)
        groups['_meta']['hostvars'][row[hostname_col].value] = {}
        for xlsx_head in rows[:1]:
            for idx, var_name in enumerate(xlsx_head):
                if row[idx].value is not None:
                    groups['_meta']['hostvars'][row[0].value][var_name.value.lower().replace(' ', '_')] = row[idx].value

    return groups


if __name__ == "__main__":
    main()
