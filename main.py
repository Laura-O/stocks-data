#!/usr/bin/env python3
import os
import sys
import json
from argparse import ArgumentParser
from source.fmp import ProfileFMP
from source.create_sheet import create_excel

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

    create_excel(data_dir, company, symbol)

parser = ArgumentParser()
subparsers = parser.add_subparsers(dest="action", title='Subcommands')

load_parser = subparsers.add_parser('load', help='laod data')
load_parser.add_argument('symbols', type=str, nargs='*', help='Stock symbol')

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
