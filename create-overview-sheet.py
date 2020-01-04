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

import openpyxl
import datetime
import math
import os

import config.srrs_config as config

input_file = os.path.join(config.bookings_dir_fqn, "bookings-{}.xlsx".format(config.target_year))

# The sheet Översikt must already exist, simply copy and past from previous year.
wb = openpyxl.load_workbook(filename = input_file) # fields with formula's contain the formulas
sheet = wb["Översikt"] # must exist already

# start row and end row (rows which contain shelfs)
shelf_start = 3
shelf_end = 36

# two halfs are needed, because at column 26, it starts with AA, AB, etc.
def create_first_half():
    first_week = 3
    last_week = 26
    for week in range(first_week, last_week):
        column = chr(ord("A") + week)
        for row in range(shelf_start, shelf_end):
            coord = column + str(row)
            #print("coord:", coord)
            sheet[coord] = "='V" + '{:02d}'.format(week) + "'!H" + str(row + 1)

def create_second_half():
    first_week = 27
    last_week = 52
    for week in range(first_week, last_week):
        column = chr(ord("A") + week - 26) # minus 26 because we are restarting the alphabet
        for row in range(shelf_start, shelf_end):
            coord = "A" + column + str(row)
            #print("coord:", coord)
            sheet[coord] = "='V" + '{:02d}'.format(week) + "'!H" + str(row + 1)
            
def main():
    create_first_half()
    print("Finished creating first half.")
    create_second_half()
    print("Finished creating second half.")
    wb.save(input_file)
    print("Finished creating the overview sheet.")
    
main()
