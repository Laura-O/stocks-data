#!/usr/bin/env python3
import os
import sys
import json
import xlsxwriter
from argparse import ArgumentParser
from source.fmp import ProfileFMP

from xlsxwriter.utility import xl_rowcol_to_cell


data_dir = os.path.join(os.path.dirname(__file__), 'data')
if not os.path.isdir(data_dir):
    os.makedirs(data_dir)

def fetch_data_by_symbol(symbol):
    fmp_company = ProfileFMP(symbol)
    
    return {
            'symbol': symbol,
            'profile': fmp_company.profile,
            'rating': fmp_company.rating,
            'income': fmp_company.income,
            }

def load(symbol):
    company = fetch_data_by_symbol(symbol)
    filename = os.path.join(data_dir, symbol + '.json')
    
    with open(filename, 'w') as file:
        json.dump(company, file)

    create_excel(company, symbol)

def create_excel(data, symbol):
    workbook = xlsxwriter.Workbook(data_dir + '/' + symbol + '.xlsx')
    worksheet = workbook.add_worksheet()

    million_format = workbook.add_format({'num_format': ' #,##0.0,, "M";(#,##0.0,, "M")'})
    bold = workbook.add_format({'bold': True})

    worksheet.write(2, 0, 'Q', bold)
    worksheet.write(3, 0, 'Revenue', bold)
    worksheet.write(4, 0, 'Cost of Revenue', bold)
    worksheet.write(5, 0, 'Gross Profit', bold)
    worksheet.write(6, 0, 'Gross Profit Ratio', bold)
    worksheet.write(8, 0, 'Selling and Marketing', bold)
    worksheet.write(9, 0, 'Research & Development', bold)
    worksheet.write(10, 0, 'General & Admin', bold)
    worksheet.write(11, 0, 'Other', bold)
    worksheet.write(12, 0, 'Total OpEx', bold)
    
    row = 1
    col = 1

    for x in reversed(data['income']):
        worksheet.set_column(col, col, 10)
        worksheet.write(row + 1, col, x['date'])
        worksheet.write(row + 2, col, x['revenue'], million_format)
        worksheet.write(row + 3, col, x['costOfRevenue'], million_format)
        worksheet.write(row + 4, col, x['grossProfit'], million_format)
        worksheet.write(row + 5, col, x['grossProfitRatio'])
        worksheet.write(row + 7, col, x['sellingAndMarketingExpenses'], million_format)
        worksheet.write(row + 8, col, x['researchAndDevelopmentExpenses'], million_format)
        worksheet.write(row + 9, col, x['generalAndAdministrativeExpenses'], million_format)
        worksheet.write(row + 10, col, x['otherExpenses'], million_format)
        worksheet.write(row + 11, col, x['operatingExpenses'], million_format)

        cell = xl_rowcol_to_cell(row + 12, col)
        gross_profit = xl_rowcol_to_cell(row + 4, col)
        op_ex = xl_rowcol_to_cell(row + 11, col)
        worksheet.write_dynamic_array_formula('%s' % cell, '=%s-%s' %(gross_profit, op_ex), million_format)

        col += 1

    workbook.close()


parser = ArgumentParser()
subparsers = parser.add_subparsers(dest="action", title='Subcommands')

load_parser = subparsers.add_parser('load', help='laod data')
load_parser.add_argument('symbols', type=str, nargs='*', help='stock symbol')

args = sys.argv[1:]
args = parser.parse_args(args)

if args.action == 'load':
    symbols = args.symbols

    for symbol in symbols:
        print("Loading data for {}.".format(symbol))
        load(symbol)
    sys.exit(0)
else:
    parser.error('Unknown command: ' + repr(args.action))
