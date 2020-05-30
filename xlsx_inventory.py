#!/usr/bin/env python

import json
import os
import argparse
import configparser
import six
from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

try:
    FileNotFoundError
except NameError:
    FileNotFoundError = IOError


config_file = "xlsx_inventory.cfg"
default_group = "NO_GROUP"


def find_config_file():
    env_name = "EXCEL_INVENTORY_CONFIG"
    if env_name in os.environ:
        return os.environ[env_name]
    else:
        return config_file


def main():
    args = parse_args()
    if args.config:
        create_config(
            filename=args.file,
            group_by_col=args.group_by_col,
            hostname_col=args.hostname_col,
            sheet=args.sheet,
        )
    config_path = find_config_file()
    config = load_config(config_path)
    try:
        wb = load_workbook(config["xlsx_inventory_file"])
        sheet = wb[config["sheet"]] if "sheet" in config else wb.active
        inventory = sheet_to_inventory(
            group_by_col=config["group_by_col"],
            hostname_col=config["hostname_col"],
            sheet=sheet,
        )
        if args.list:
            print(json.dumps(inventory, indent=4, sort_keys=True, default=str))
        if args.config:
            create_config(
                filename=args.file,
                group_by_col=args.group_by_col,
                hostname_col=args.hostname_col,
                sheet=args.sheet,
            )
        elif args.host:
            try:
                print(
                    json.dumps(
                        inventory["_meta"]["hostvars"][args.host],
                        indent=4,
                        sort_keys=True,
                        default=str,
                    )
                )
            except KeyError as e:
                print('\033[91mHost "%s" not Found!\033[0m' % e)
                print(e)
    except FileNotFoundError as e:
        print(
            "\033[91mFile Not Found! Check %s configuration file!"
            " Is the `xlsx_inventory_file` path setting correct?\033[0m" % config_path
        )
        print(e)
        exit(1)
    except KeyError as e:
        print(
            "\033[91mKey Error! Check %s configuration file! Is the `sheet` name setting correct?\033[0m"
            % config_path
        )
        print(e)
        exit(1)
    exit(0)


def create_config(filename=None, group_by_col=None, hostname_col=None, sheet=None):
    config = configparser.ConfigParser()
    config["xlsx_inventory"] = {}
    if filename is None:
        print("\033[91m--file is required!\033[0m")
        exit(1)
    config["xlsx_inventory"]["xlsx_inventory_file"] = filename
    if group_by_col is not None:
        config["xlsx_inventory"]["group_by_col"] = group_by_col
    if hostname_col is not None:
        config["xlsx_inventory"]["hostname_col"] = hostname_col
    if sheet is not None:
        config["xlsx_inventory"]["sheet"] = sheet
    with open(find_config_file(), "w") as cf:
        config.write(cf)


def load_config(config_path):
    config = configparser.ConfigParser()
    config["DEFAULT"] = {"hostname_col": "A", "group_by_col": "B"}
    if len(config.read(config_path)) > 0:
        return config["xlsx_inventory"]
    else:
        print('\033[91mConfiguration File "%s" not Found!\033[0m' % config_path)
        exit(1)


def parse_args():
    arg_parser = argparse.ArgumentParser(
        description="Excel Spreadsheet Inventory Module"
    )
    group = arg_parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--list", action="store_true", help="List active servers")
    group.add_argument(
        "--host", help="List details about the specified host", default=None
    )
    group.add_argument("--config", action="store_true", help="Create Config File")
    arg_parser.add_argument(
        "--file", default=None, help="Excel Spreadsheet file used by xlsx_inventory.py"
    )
    arg_parser.add_argument(
        "--group-by-col", default=None, help="Column to group hosts by (i.E. `B`)"
    )
    arg_parser.add_argument(
        "--hostname-col", default=None, help="Column containing the hostnames"
    )
    arg_parser.add_argument(
        "--sheet", default=None, help="Name of the Sheet, used by xlsx_inventory.py"
    )
    return arg_parser.parse_args()


def sheet_to_inventory(group_by_col, hostname_col, sheet):
    if isinstance(group_by_col, six.string_types):
        group_by_col = (
            column_index_from_string(coordinate_from_string(group_by_col + "1")[0]) - 1
        )
    if isinstance(hostname_col, six.string_types):
        hostname_col = (
            column_index_from_string(coordinate_from_string(hostname_col + "1")[0]) - 1
        )

    groups = {"_meta": {"hostvars": {}}}
    rows = list(sheet.rows)

    for row in rows[1:]:
        host = row[hostname_col].value
        if host is None:
            continue
        group = row[group_by_col].value
        if group is None:
            group = default_group
        if group not in groups.keys():
            groups[group] = {"hosts": [], "vars": {}}
        groups[group]["hosts"].append(host)
        groups["_meta"]["hostvars"][row[hostname_col].value] = {}
        for xlsx_head in rows[:1]:
            for idx, var_name in enumerate(xlsx_head):
                if var_name.value is None:
                    var_name.value = "xlsx_" + var_name.coordinate
                if row[idx].value is not None:
                    groups["_meta"]["hostvars"][row[0].value][
                        var_name.value.lower().replace(" ", "_")
                    ] = row[idx].value

    return groups


if __name__ == "__main__":
    main()
