import csv
import json

csvfile = open('CardExcel.csv', 'r')
jsonfile = open('CardExcel.json', 'w')

fieldnames = ("Date","Report Type","Scenario","Actual_Model","Asset_Group","Asset_Class","Loss Forecast Q", "Alternate_Time_Period" ,"Loss Forecast $","ALLL Forecast $","Provision Forecast $","NPL Forecast $","Reported Balance $")


reader = csv.DictReader( csvfile, fieldnames)
for row in reader:
    json.dump(row, jsonfile)
    jsonfile.write('\n')