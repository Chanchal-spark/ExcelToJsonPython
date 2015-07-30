import xlrd
from collections import OrderedDict
import json
import datetime

# Open the workbook and select the first worksheet
wb = xlrd.open_workbook('CardExcel.xlsx')
sh = wb.sheet_by_index(1)

# List to hold dictionaries
card_list = []

# Iterate through each row in worksheet and fetch values into dict
for rownum in range(1, sh.nrows):
	card = OrderedDict()
	row_values = sh.row_values(rownum)


	if row_values[0]!='':
		card['Date']=datetime.datetime(* (xlrd.xldate_as_tuple(row_values[0],wb.datemode))).strftime('%m/%d/%Y')


	else:

		card['Date']=row_values[0]



	card['Report Type'] = row_values[1]

	card['Scenario'] = row_values[2]

	card['Actual_Model'] = row_values[3]

	card['Asset_Group'] = row_values[4]

	card['Asset_Class'] = row_values[5]

	if row_values[6]!='':
		card['Loss Forecast Q']=datetime.datetime(* (xlrd.xldate_as_tuple(row_values[6],wb.datemode))).strftime('%m/%d/%Y')


	else:
		card['Loss Forecast Q']=row_values[6]




	card['Alternate_Time_Period'] = row_values[7]

	card['Loss Forecast $'] = row_values[8]

	card['ALLL Forecast $'] = row_values[9]

	card['Provision Forecast $'] = row_values[10]

	card['NPL Forecast $'] = row_values[11]

	card['Reported Balance $'] = row_values[12]




	card_list.append(card)

# Serialize the list of dicts to JSON
j = json.dumps(card_list)

# Write to file
with open('CardExcel_chanchal.json', 'w') as f:
    f.write(j)
