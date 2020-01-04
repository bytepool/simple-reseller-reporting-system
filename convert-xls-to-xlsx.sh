#!/bin/bash

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



# OS can be set via CLI. Options: "win" or "linux". Default is "win". 
os=${1:-win}

echo -e "\nIf libreoffice (word or excel equivalent) is running, \nyou MUST close it before running this script, or the script will not work!\n"

if [ $os = "win" ]; then

    program="/c/Program Files/LibreOffice/program/soffice.bin"

    if [ ! -f "$program" ]; then
		echo "ERROR: Could not find $program. Make sure LibreOffice is installed and that the given path is correct."
		exit 1
    else
		echo "OK: Found $program..."
    fi
	python="/c/users/melisa/appdata/local/programs/python/python37-32/python.exe"

	
elif [ $os = "linux" ]; then

    program="soffice"

	if [ ! "$($program --version)" ]; then
		echo "ERROR: Could not find $program. Make sure LibreOffice is installed and in PATH."
		exit 1
	else
		echo "OK: Found soffice in PATH..."
	fi
	python="/usr/bin/python3"
fi

if [ ! -f "$python" ]; then
		echo "ERROR: Could not find $python. Make sure python is installed and that the given path is correct."
		exit 1
else
		echo "OK: Found $python..."
fi


# read directories from khconfig.py
sales_dir=$($python -c "import khconfig; print(khconfig.sales_dir_fqn)")
year=$($python -c "import khconfig; print(khconfig.target_year)")

dir="${sales_dir}/${year}"

if [ ! -d "$dir" ]; then
    echo "ERROR: Directory $dir does not exist. Make sure khconfig.py is correctly configured: check sales_dir_fqn and target_year."
    exit 2
else
    echo "OK: Found $dir..."
fi

# make sure *.xls is replaced with nothing if no files with that extension exist,
# instead of using the literal '*.xls' string. 
shopt -s nullglob

tmp_dir="$dir"/conv_tmp
mkdir -p "$tmp_dir"

echo -e "\nStarting conversions.\n"

for f in "$dir"/*.xlsx "$dir"/*.xls; do
    echo "Converting $f..."
    "$program" --headless --convert-to xlsx --outdir "$tmp_dir" "$f"
done

mv "$tmp_dir"/*.xlsx "$dir"
rm -rf "$tmp_dir"

echo -e "\nDone converting files. Press any key to continue. \n"

read -n 1

