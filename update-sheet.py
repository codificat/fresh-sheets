#!/usr/bin/env python

import argparse
import csv
import json
import logging
import os
import string

from googleapiclient import discovery
from google.oauth2.service_account import Credentials

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
logging.basicConfig(format='%(asctime)s : %(levelname)s : %(message)s',
                    level=logging.INFO)


class EnvDefault(argparse.Action):
    def __init__(self, envvar, required=True, default=None, **kwargs):
        if not default and envvar:
            if envvar in os.environ:
                default = os.environ[envvar]
        if required and default:
            required = False
        super(EnvDefault, self).__init__(default=default, required=required,
                                         **kwargs)

    def __call__(self, parser, namespace, values, option_string=None):
        setattr(namespace, self.dest, values)


def connect_to_api(key_file):
    credentials = Credentials.from_service_account_file(
        key_file,
        scopes=SCOPES)
    return discovery.build('sheets', 'v4', credentials=credentials)


def read_csv(path, delimiter):
    logging.info("Reading raw import data from CSV file")
    data = []
    with open(path) as csvfile:
        csvreader = csv.reader(csvfile, delimiter=delimiter)
        for row in csvreader:
            data.append(row)
    return data


def upload_raw_data(service, document_id, doc_range, values):
    logging.info("Uploading raw data")
    body = {
        'values': values
    }

    request = service.spreadsheets().values().update(
        spreadsheetId=document_id, range=doc_range,
        valueInputOption='USER_ENTERED', body=body
    )

    response = request.execute()
    logging.debug(json.dumps(response))


def read_accounts(service, doc_id, doc_range):
    """Returns a dictionary that maps cost centers to the list
    of their account IDs"""
    logging.info("Obtaining account/cost center relationship")
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=doc_id,
                                range=doc_range).execute()
    values = result.get('values', [])
    if not values:
        raise KeyError

    header = values[0]
    acct_col = header.index('Account ID')  # FIXME: hardcoded
    cc_col = header.index('Cost Center')   # FIXME: hardcoded
    accounts = dict()
    for account in values[1:]:
        if not account:  # we stop at an empty row
            break
        acct_id = account[acct_col]
        acct_cc = account[cc_col]
        if acct_cc and acct_id:  # only add complete info
            if acct_cc in accounts.keys():
                accounts[acct_cc].append(acct_id)
            else:
                accounts[acct_cc] = [acct_id]
    return accounts


def values_for(raw_data, acct_id, skip_rows=[]):
    """Extract raw data for acct_id.
    Returns a list of row values in the format expected by
    the spreadsheets API, e.g. [[1], [2], [3]]"""
    header = raw_data[0]
    try:
        acct_col = header.index(acct_id)
    except ValueError:
        logging.warn(f'Could nof find raw data for account {acct_id}')
        return None
    values = [[acct_id]]
    for i, row in enumerate(raw_data[1:]):
        if i in skip_rows:
            continue
        values.append([row[acct_col]])
    return values


def update_sheets(service, doc_id, accounts, raw_data):
    # 'accounts' is a dict: cc => list of account_ids
    logging.info("Updating cost center sheets")
    first_col = 2  # FIXME: hardcoded
    first_row = 3  # FIXME: hardcoded
    max_cols = 20  # FIXME: hardcoded
    data = []
    for cc in accounts:
        cols = (c for c in string.ascii_uppercase[first_col:max_cols])
        data.append({  # Update the cell with the CC name
            "range": f"'{cc}'!B1",  # FIXME: hardcoded
            "values": [[cc]]
        })
        for acct_id in accounts[cc]:
            values = values_for(raw_data, acct_id, skip_rows=[1])
            if not values:
                continue
            col = next(cols)
            range_name = f"'{cc}'!{col}{first_row}:{col}"
            data.append({
                "range": range_name,
                "values": values
            })
        # Clear the values of remaining columns
        for col in cols:
            range_name = f"'{cc}'!{col}{first_row}:{col}"
            data.append({
                "range": range_name,
                "values": 8 * [['']]  # FIXME: hardcoded number of rows
            })
    if data:
        body = {
            'valueInputOption': 'USER_ENTERED',
            'data': data
        }
        logging.debug("Updating sheets: " + str(body))
        request = service.spreadsheets().values().batchUpdate(
            spreadsheetId=doc_id,
            body=body
        )
        response = request.execute()
        logging.debug('Response: ' + json.dumps(response))


def parse_args():
    parser = argparse.ArgumentParser(
        description='fresh-sheets command line arguments'
    )
    parser.add_argument('-k', '--key', type=str, action=EnvDefault,
                        envvar='KEY',
                        help='The file with the Google service account key')
    parser.add_argument('-c', '--csv-file', type=str,
                        action=EnvDefault, envvar='CSV_FILE',
                        help='The path to the CSV data to import')
    parser.add_argument('-t', '--target-id', type=str,
                        action=EnvDefault, envvar='TARGET_ID',
                        help='The UUID of the Google Sheet to update')
    parser.add_argument('-s', '--sheet-name', type=str,
                        action=EnvDefault, envvar='SHEET_NAME',
                        help='Name of the target sheet to import raw data')
    parser.add_argument('-a', '--accounts-doc-id', type=str,
                        action=EnvDefault, envvar='ACCOUNTS_DOC_ID',
                        help='The UUID of the reference Google Spreadsheet')
    parser.add_argument('-i', '--accounts-sheet', type=str,
                        action=EnvDefault, envvar='ACCOUNTS_SHEET',
                        help='ID of the sheet within the reference doc')
    parser.add_argument('--delimiter', type=str, default=',',
                        action=EnvDefault, envvar='DELIMITER',
                        help='The delimiter used to separate column values')
    return parser.parse_args()


def main():
    args = parse_args()
    service = connect_to_api(args.key)
    data = read_csv(args.csv_file, args.delimiter)
    logging.debug('Raw data to import: ' + str(data))
    accounts = read_accounts(
        service,
        args.accounts_doc_id,
        args.accounts_sheet
    )
    logging.debug("Accounts per Cost Center: " + str(accounts))
    upload_raw_data(
        service,
        args.target_id,
        args.sheet_name,
        data
    )
    update_sheets(service, args.target_id, accounts, data)
    logging.info("Finished.")


if __name__ == '__main__':
    main()
