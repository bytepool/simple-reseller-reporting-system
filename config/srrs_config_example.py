# simple reseller reporting system

from os import path 

# set targeted year
target_year = 2020

# set directories
#home_dir_fqn = path.expanduser("~")
data_dir_fqn = path.join(".", "data")

sales_dir = "sales_data"
bookings_dir = "bookings"
customer_reg_dir = "customers"
report_dir = "reports"
template_dir = "templates"


# set file names
customer_reg_file = "customers.xlsx"
bookings_file = "bookings-" + str(target_year) + ".xlsx"

report_template_file = "report_template.tex"
email_template_file = "email_template.txt"

# set latex command to be used for report generation
latex_cmd = "pdflatex"
latex_options = "-halt-on-error"

# create fully qualified names to be used in the scripts
sales_dir_fqn = path.join(data_dir_fqn, sales_dir) 
bookings_dir_fqn = path.join(data_dir_fqn, bookings_dir)
customer_reg_dir_fqn = path.join(data_dir_fqn, customer_reg_dir)
report_dir_fqn = path.join(data_dir_fqn, report_dir)
template_dir_fqn = path.join(".", template_dir)

bookings_file_fqn = path.join(bookings_dir_fqn, bookings_file)
customer_reg_file_fqn = path.join(customer_reg_dir_fqn, customer_reg_file)

report_template_file_fqn = path.join(template_dir_fqn, report_template_file)
email_template_file_fqn = path.join(template_dir_fqn, email_template_file)

# emails
smtp_server = "send.foo.com"
smtp_port = 587

email_login_name = "contact@foo.com"
email_subject = "Sales at Foo in week %s, shelf %s"

# set data sheet relevant anchors - WARNING: setting these wrongly will give wrong results!
customer_reg_first_customer_row = 3 # customer 1 starts here
bookings_first_row = 5 # Shelf H01 starts here
sales_first_data_row = 7
