# import csv
#
# with open('vehicles.csv') as csv_file:
#     csv_reader = csv.reader(csv_file, delimiter=',')
#     line_count = 0
#     for row in csv_reader:
#             print(f'Column names are {"; ".join(row)}')
#             line_count += 1

import pandas
data = pandas.read_csv('vehicles.csv', delimiter=";")
print(data)


import argparse

parser = argparse.ArgumentParser()

parser.add_argument('-k', '--keys', type=str, required=True)
parser.add_argument('-c', '--colored', type=bool, default=True)

args = parser.parse_args()
k = args.x
c = args.y

