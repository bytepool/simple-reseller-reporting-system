#!/usr/bin/python3

# simple reseller reporting system
# 
# Copyright (C) 2016 - 2019, Aljoscha Lautenbach. 
#
# This program is free software; you can redistribute it
# and/or modify it under the terms of the GNU General
# Public License version 2 as published by the Free
# Software Foundation
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public
# License along with this program; if not, write to the
# Free Software Foundation, Inc., 51 Franklin Street,
# Fifth Floor, Boston, MA 02110-1301 USA.


import openpyxl
from openpyxl.styles.borders import Border, Side

import datetime
import os

import config.srrs_config as config

import lib_srrs_util as util


def main():
    products = ["S03"]
    year = config.target_year
    sales_dir = os.path.join(config.sales_dir_fqn, str(year))
    
    input_files = [f for f in os.listdir(sales_dir) if os.path.isfile(os.path.join(sales_dir, f))]
    print("\nSearching for sales in: " + str(input_files))
    
    for i in range(0, len(input_files)):
        input_files[i] = os.path.join(sales_dir, input_files[i])

    target = 0 #int(input("\nEnter the quarter number (1-4): "))
    
    days = set()
    sheets = []

    print("\nLoading sales data...")
    
    for input_file in input_files:
        wb = openpyxl.load_workbook(filename = input_file, data_only=True)
        sheets.append(wb['Sheet 1'])
    
    print("\nSumming sales...\n")
    
    for product in products:
        item_sum = 0
    
        # row A = date, row B = time, row E = product, row F = variant, row H = amount, row K = price
        for sheet in sheets:
            # iterate over all cells that contain data
            for row in range(config.sales_first_data_row, sheet.max_row + 1):
                date_str = sheet["A" + str(row)].value
                date = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        
                # check if the current row is within the given time window
                if (util.is_target(date, target, days)):
                    cur_product = sheet["E" + str(row)].value
                    if (product in cur_product):
                        #print("product in cur_product:", cur_product)
                        amount = sheet["H" + str(row)].value
                        price = sheet["K" + str(row)].value
                        
                        item_sum = item_sum + price
        
        # computations are done here
        
        print ("In chosen time period, product {:20s} sold for: {:5d} sek.".format(
            product, int(item_sum)))
    
    util.pause()

    
if __name__ == '__main__':
    main()

    
