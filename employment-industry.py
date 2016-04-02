import requests
import json
from openpyxl import Workbook

# API pull data below taken from sample at http://www.bls.gov/developers/api_python.htm#python2
headers = {'Content-type': 'application/json'}

SeriesID_list = ['SMS1100000600000001', 
'SMS11000006000000001',
'SMS11000001000000001',
'SMS11000001100000001',
'SMS11000001110000001',
'SMS11000001120000001',
'SMS11000001130000001',
'SMS11000001140000001',
'SMS11000001150000001',
'SMS11000002000000001',
'SMS11000002100000001',
'SMS11000002110000001',
'SMS11000002120000001',
'SMS11000002130000001',
'SMS11000002200000001',
'SMS11000002300000001',
'SMS11000002360000001',
'SMS11000002370000001',
'SMS11000002380000001',
'SMS11000003000000001',
'SMS11000003100000001',
'SMS11000003110000001',
'SMS11000003120000001',
'SMS11000003130000001',
'SMS11000003140000001',
'SMS11000003150000001',
'SMS11000003160000001',
'SMS11000003200000001',
'SMS11000003210000001',
'SMS11000003220000001',
'SMS11000003230000001',
'SMS11000003240000001',
'SMS11000003250000001',
'SMS11000003260000001',
'SMS11000003270000001',
'SMS11000003300000001',
'SMS11000003310000001',
'SMS11000003320000001',
'SMS11000003330000001',
'SMS11000003340000001',
'SMS11000003350000001',
'SMS11000003360000001',
'SMS11000003370000001',
'SMS11000003390000001',
'SMS11000004000000001',
'SMS11000004200000001',
'SMS11000004230000001',
'SMS11000004240000001',
'SMS11000004250000001',
'SMS11000004400000001',
'SMS11000004410000001',
'SMS11000004420000001',
'SMS11000004430000001',
'SMS11000004440000001',
'SMS11000004450000001',
'SMS11000004460000001',
'SMS11000004470000001',
'SMS11000004480000001',
'SMS11000004500000001',
'SMS11000004510000001',
'SMS11000004520000001',
'SMS11000004530000001',
'SMS11000004540000001',
'SMS11000004800000001',
'SMS11000004810000001',
'SMS11000004820000001',
'SMS11000004830000001',
'SMS11000004840000001',
'SMS11000004850000001',
'SMS11000004860000001',
'SMS11000004870000001',
'SMS11000004880000001',
'SMS11000004900000001',
'SMS11000004910000001',
'SMS11000004920000001',
'SMS11000004930000001',
'SMS11000005000000001',
'SMS11000005100000001',
'SMS11000005110000001',
'SMS11000005120000001',
'SMS11000005150000001',
'SMS11000005160000001',
'SMS11000005170000001',
'SMS11000005180000001',
'SMS11000005190000001',
'SMS11000005200000001',
'SMS11000005210000001',
'SMS11000005220000001',
'SMS11000005230000001',
'SMS11000005240000001',
'SMS11000005250000001',
'SMS11000005300000001',
'SMS11000005310000001',
'SMS11000005320000001',
'SMS11000005330000001',
'SMS11000005400000001',
'SMS11000005500000001',
'SMS11000005500000001',
'SMS11000005600000001',
'SMS11000005610000001',
'SMS11000005620000001',
'SMS11000006000000001',
'SMS11000006100000001',
'SMS11000006200000001',
'SMS11000006210000001',
'SMS11000006220000001',
'SMS11000006230000001',
'SMS11000006240000001',
'SMS11000006500000001',
'SMS11000007000000001',
'SMS11000007100000001',
'SMS11000007110000001',
'SMS11000007120000001',
'SMS11000007130000001',
'SMS11000007200000001',
'SMS11000007210000001',
'SMS11000007220000001',
'SMS11000008100000001',
'SMS11000008110000001',
'SMS11000008120000001',
'SMS11000008130000001',
'SMS11000008140000001'
] 

Industries_list = ['Goods-Producing Industries', 
'Mining and Logging', 
'Agriculture, Forestry, Fishing and Hunting', 
'Crop Production', 
'Animal Production', 
'Forestry and Logging', 
'Fishing, Hunting and Trapping', 
'Support Activities for Agriculture and Forestry', 
'Construction', 
'Mining, Quarrying, and Oil and Gas Extraction', 
'Oil and Gas Extraction', 
'Mining', 
'Support Activities for Mining', 
'Utilities', 
'Construction', 
'Construction of Buildings', 
'Heavy and Civil Engineering Construction', 
'Specialty Trade Contractors', 
'Manufacturing', 
'Manufacturing', 
'Food Manufacturing', 
'Beverage and Tobacco Product Manufacturing', 
'Textile Mills', 
'Textile Product Mills', 
'Apparel Manufacturing', 
'Leather and Allied Product Manufacturing', 
'Manufacturing', 
'Wood Product Manufacturing', 
'Paper Manufacturing', 
'Printing and Related Support Activities', 
'Petroleum and Coal Products Manufacturing', 
'Chemical Manufacturing', 
'Plastics and Rubber Products Manufacturing', 
'Nonmetallic Mineral Product Manufacturing', 
'Manufacturing', 
'Primary Metal Manufacturing', 
'Fabricated Metal Product Manufacturing', 
'Machinery Manufacturing', 
'Computer and Electronic Product Manufacturing', 
'Electrical Equipment, Appliance, and Component Manufacturing', 
'Transportation Equipment Manufacturing', 
'Furniture and Related Product Manufacturing', 
'Miscellaneous Manufacturing', 
'Trade, Transportation, and Utilities', 
'Wholesale Trade', 
'Merchant Wholesalers, Durable Goods', 
'Merchant Wholesalers, Nondurable Goods', 
'Wholesale Electronic Markets and Agents and Brokers', 
'Retail Trade', 
'Motor Vehicle and Parts Dealers', 
'Furniture and Home Furnishings Stores', 
'Electronics and Appliance Stores', 
'Building Material and Garden Equipment and Supplies Dealers', 
'Food and Beverage Stores', 
'Health and Personal Care Stores', 
'Gasoline Stations', 
'Clothing and Clothing Accessories Stores', 
'Retail Trade', 
'Sporting Goods, Hobby, Book, and Music Stores', 
'General Merchandise Stores', 
'Miscellaneous Store Retailers', 
'Nonstore Retailers', 
'Transportation and Warehousing', 
'Air Transportation', 
'Rail Transportation', 
'Water Transportation', 
'Truck Transportation', 
'Transit and Ground Passenger Transportation', 
'Pipeline Transportation', 
'Scenic and Sightseeing Transportation', 
'Support Activities for Transportation', 
'Transportation and Warehousing', 
'Postal Service', 
'Couriers and Messengers', 
'Warehousing and Storage', 
'Information', 
'Information', 
'Publishing Industries', 
'Motion Picture and Sound Recording Industries', 
'Broadcasting', 
'Internet Publishing and Broadcasting', 
'Telecommunications', 
'Data Processing, Hosting, and Related Services', 
'Other Information Services', 
'Finance and Insurance', 
'Monetary Authorities - Central Bank', 
'Credit Intermediation and Related Activities', 
'Securities, Commodity Contracts, and Other Financial Investments and Related Activities', 
'Insurance Carriers and Related Activities', 
'Funds, Trusts, and Other Financial Vehicles', 
'Real Estate and Rental and Leasing', 
'Real Estate', 
'Rental and Leasing Services', 
'Lessors of Nonfinancial Intangible Assets', 
'Professional, Scientific, and Technical Services', 
'Financial Activities', 
'Management of Companies and Enterprises', 
'Administrative and Support and Waste Management and Remediation Services', 
'Administrative and Support Services', 
'Waste Management and Remediation Services', 
'Professional and Business Services', 
'Educational Services', 
'Health Care and Social Assistance', 
'Ambulatory Health Care Services', 
'Hospitals', 
'Nursing and Residential Care Facilities', 
'Social Assistance', 
'Education and Health Services', 
'Leisure and Hospitality', 
'Arts, Entertainment, and Recreation', 
'Performing Arts, Spectator Sports, and Related Industries', 
'Museums, Historical Sites, and Similar Institutions', 
'Amusement, Gambling, and Recreation Industries', 
'Accommodation and Food Services', 
'Accommodation', 
'Food Services and Drinking Places', 
'Other Services', 
'Repair and Maintenance', 
'Personal and Laundry Services', 
'Religious, Grantmaking, Civic, Professional, and Similar Organizations', 
'Private Households']

dictionary = dict(zip(SeriesID_list, Industries_list))

#add series ID here, make a list 
data = json.dumps({"seriesid": [dictionary], "startyear":"2000", "endyear":"2015", 'registrationKey':"8623bfd2de5f4d17b4baca7d01c14843"}) 
#this is JW'skey
p = requests.post('http://api.bls.gov/publicAPI/v2/timeseries/data/', data=data, headers=headers)

json_data = json.loads(p.text)

# for series in json_data['Results']['series']:
#     # Data is a list of dicts
#     for record in series['data']:
#         print series['seriesID']
#         print record['year']
#         print record['periodName']
#         print record['value']
#         print "\n"

# # Create a workbook with a sheet named ForTableau
wb = Workbook(guess_types=True)
ForTableau = wb.active
ForTableau.title = "ForTableau"

# Write the column headers
ForTableau['A1'] = 'Date'
ForTableau['B1'] = 'Total Employment'
ForTableau['C1'] = 'Industry'
ForTableau['D1'] = 'Series ID'

for x in range(len(SeriesID_list)):
	for series in json_data['Results']['series']:
		for i, record in enumerate(series['data']):

			a = ForTableau.cell(row = record+2, column = 1)
			a.value = record['periodName']+'-01-'+record['year']

			b = ForTableau.cell(row = record+2, column = 2)
			b.value = float(record['value'])*1000

		# c = ForTableau.cell(row = i+2, column = 3)
		# c.value = Industries_list

		# d = ForTableau.cell(row = i+2, column = 4)
		# d.value = SeriesID_list

  # elif series['seriesID'] == DCTotalPrivateEmployment:
  #   for i, record in enumerate(series['data']):
  #       c= ForTableau.cell(row = i+2, column = 3)
  #       c.value = float(record['value'])*1000

wb.save('Employment_by_Industry.xlsx')
