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


class CustomerEntry:
    def __init__(self, sheet, row):
        self.sheet = sheet
        self.row_nr = str(row)
        self.customers_first_row = 3 # kund 101 starts here

        self.customer_nr =  int(self.sheet["A" + self.row_nr].value)
        self.name = self.sheet["B" + self.row_nr].value
        self.person_nr = self.sheet["C" + self.row_nr].value
        self.address = self.sheet["D" + self.row_nr].value
        self.phone = self.sheet["E" + self.row_nr].value
        self.email = self.sheet["G" + self.row_nr].value
        self.account_nr = self.sheet["H" + self.row_nr].value

    def print_entry(self):
        print("Customer Nr: %d" % self.customer_nr)
        print("Name: %s" % self.name)
        print("Person Nr: %s" % self.person_nr)
        print("Address: %s" % self.address)
        print("Phone Nr: %s" % self.phone)
        print("Email: %s" % self.email)
        print("Account Nr: %s\n" % self.account_nr)

        
def get_customer_entry(sheet, customer_nr):
    first_row = customer_nr - 98
    # iterate over all cells that contain data
    for row in range(first_row, sheet.max_row + 1):
        try:
            entry = int(sheet["A" + str(row)].value)
            if (entry == customer_nr):
                customer = CustomerEntry(sheet, row)
                return customer
        except: # not a data row
            continue
        
    print  ("Customer nr %d not found!" % customer_nr)

