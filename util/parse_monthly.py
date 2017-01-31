#parse austin monthly reports

import xlrd
import argparse
from datetime import datetime
import csv
import os

def main():
	print('hello')
	#for each monthly report write two rows
	#date, dwi, narc, other, dt dwi, dt narc, dt other

	parser = argparse.ArgumentParser()
	parser.add_argument('--dir', required=True, type=str)
	parser.add_argument('--out', required=True, type=str)
	args = parser.parse_args()

	all_rows = {}

	with open(args.out,'w') as outfile:
		for filename in os.listdir(args.dir):			
			if (os.path.splitext(os.path.basename(filename))[1] == '.xls'):
				print('parsing: ',filename)
				rows = parse_apd_xls(args.dir+'/'+filename)
				all_rows[rows[0][0]] = rows[0]
				all_rows[rows[1][0]] = rows[1]

		csvwriter = csv.writer(outfile, delimiter=',',quotechar='"', quoting=csv.QUOTE_MINIMAL)

		csvwriter.writerows(all_rows.values())

def parse_apd_xls(filename):
	with xlrd.open_workbook(filename) as workbook:
		city = workbook.sheet_by_name('CityWide')
		dt = workbook.sheet_by_name('George (DT)')

		date  = city.cell(13, 3).value
		prev_date  = city.cell(13, 4).value

		date = datetime.strptime(date,'%b %Y').strftime('%Y-%m-%d')
		prev_date = datetime.strptime(prev_date,'%b %Y').strftime('%Y-%m-%d')

		city_dwi = city.cell(14, 3).value
		city_other = city.cell(18, 3).value
		city_narc = city.cell(16, 3).value

		prev_city_dwi = city.cell(14, 4).value
		prev_city_other = city.cell(18, 4).value
		prev_city_narc = city.cell(16, 4).value

		dt_dwi = dt.cell(14, 3).value
		dt_other = dt.cell(18, 3).value
		dt_narc = dt.cell(16, 3).value

		prev_dt_dwi = dt.cell(14, 4).value
		prev_dt_other = dt.cell(18, 4).value
		prev_dt_narc = dt.cell(16, 4).value

		current_row = [date,city_dwi,city_other,city_narc,dt_dwi,dt_other,dt_narc]
		prev_row = [prev_date,prev_city_dwi,prev_city_other,prev_city_narc,prev_dt_dwi,prev_dt_other,prev_dt_narc]

		return [current_row,prev_row]



if __name__ == '__main__':
	main()
