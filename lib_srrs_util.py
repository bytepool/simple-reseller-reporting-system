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


def pause():
    print("Press 'enter' to close this window.")
    input("")


def is_target_week(date, target_week, days):
    week = date.isocalendar()[1]
    if  (week == target_week):
        days.add(date.strftime("%a %Y-%m-%d"))
        return True
    else:
        return False

    
def is_target_month(date, target_month, days):
    month = date.month
    if  (month == target_month):
        days.add(date.strftime("%a %Y-%m-%d"))
        return True
    else:
        return False

    
def is_target_quarter(date, target_quarter, days):
    month = date.month
    lower_bound = (target_quarter - 1) * 3
    upper_bound = target_quarter * 3
    if  (month > lower_bound and month <= upper_bound):
        days.add(date.strftime("%a %Y-%m-%d"))
        return True
    else:
        return False

    
def is_target(date, target, days):
    return True
    month = date.month
    lower_bound = (target - 1) * 3
    upper_bound = target * 3
    if  (month > lower_bound and month <= upper_bound):
        days.add(date.strftime("%a %Y-%m-%d"))
        return True
    else:
        return False

    
def ask_target_week(year):
    errmsg = "\nERROR: Invalid week number. Please provide a valid week number!"
    
    try:
        week = int(input("\nEnter the week number (year {:s}): ".format(year)))
    except ValueError as e:
        print (errmsg)
        pause()
        exit(1)
        
    if week < 1 or week > 52:
        print(errmsg)
        pause()
        exit(1)
        
    return week
