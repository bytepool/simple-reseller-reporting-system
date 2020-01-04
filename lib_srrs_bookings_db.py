#/usr/bin/python3

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


class BookingsWorkbook:
    def __init__(self, booking_file, data_only):
        self.wb = openpyxl.load_workbook(filename = booking_file, data_only=True)
        
    def getBookingSheetOfWeekNr(self, weeknr):
        return self.wb["V" + "{:02d}".format(int(weeknr))]

    
class BookingEntry:
    def __init__(self, sheet, row):
        self.sheet = sheet
        self.row_nr = str(row)

        self.nr_items = self.sheet["B" + self.row_nr].value
        self.sales = int(self.sheet["C" + self.row_nr].value)
        self.commission = int(self.sheet["D" + self.row_nr].value)
        self.pay_out = int(self.sheet["E" + self.row_nr].value)
        self.shelf = self.sheet["G" + self.row_nr].value

        self.customer = self.sheet["H" + self.row_nr].value
        
        if self.customer is not None:
            self.customer_nr = int(self.customer.split()[0])
            self.customer_first_name = self.customer.rpartition(' ')[2]
        else:
            self.customer_nr = None
            self.customer_first_name = None
            

    def print_entry(self):
        print("Shelf: %s" % self.shelf)
        print("Customer Nr: %s" % self.customer_nr)
        print("Customer First Name: %s" % self.customer_first_name)

        print("Nr of Items: %s" % self.nr_items)
        print("Sales: %d" % self.sales)
        print("Commission: %d" % self.commission)
        print("Pay Out: %d\n" % self.pay_out)

