# Adds new customers to the customer database
from faker import Faker


def add_customer(n):
    # Adds new n customers
    f = Faker()
    num_cust = n+1
    # Generate customer details for n customers.
    for customer in range(1, num_cust):
        print customer , ":", f.first_name(), f.last_name()
        print " " , f.phone_number()
        print " ", f.email()
        print " ", f.date_time_this_month()
        print "---------------------"

    # TODO: Add code to write these details to a database

if __name__ == "__main__":
    add_customer(5)
