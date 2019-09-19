import pandas as pd
from pandas import ExcelWriter

# Open the quotation file
excel_file = 'Box-fans.xlsx'
df = pd.read_excel(excel_file)

# Create layout columns
col = ['Type', 'Size', 'Motor type', 'Motor Power', 'Inlet',
'Power Switch', 'Pressure Switch', 'Roof', 'Outlet Rain Prot.', 'VFD',
'Filter', 'Double Skin Panels']

# Create dictionaries to map the keywords
dict_type = {
	'KPB': ['kpb', 'kp'],
	'KVBF': ['kvb', 'kvbf', 'kvb-f', 'kvbf/f']
}

dict_size = {
	9: ['09/09', '9/9', '09-09', '09/09'],
	10: ['10/10', '10-10'],
	12: ['12/12', '12-12'],
	15: ['15/15', '15-15'],
	18: ['18/18', '18-18'],
	20: ['20/20', '20-20'],
	22: ['22/22', '22-22'],
	25: ['25/25', '25-25'],
	30: ['30/28', '30-28']
}

dict_speed = {
	'One speed': ['1 speed', '1speed', '1 vites', '1vites', '1v'],
	'Two speeds': ['2 speed', '2speed', '2 vites', '2vites', '2v']
}

dict_motor = {
	0.25: ['0.25 kw', '0.25kw', '0,25 kw', '0,25kw'],
	0.37: ['0.37 kw', '0.37kw', '0,37 kw', '0,37kw'],
	0.55: ['0.55 kw', '0.55kw', '0,55 kw', '0,55kw'],
	0.75: ['0.75 kw', '0.75kw', '0,75 kw', '0,75kw'],
	1.1: ['1.1 kw', '1.1kw', '1,1 kw', '1,1kw'],
	1.5: ['1.5 kw', '1.5kw', '1,5 kw', '1,5kw'],
	2.2: ['2.2 kw', '2.2kw', '2,2 kw', '2,2kw'],
	3: ['3 kw', '3kw'],
	4: ['4 kw', '4kw'],
	5.5: ['5.5 kw', '5.5kw', '5,5 kw', '5,5kw'],
	7.5: ['7.5 kw', '7.5kw', '7,5 kw', '7,5kw'],
	11: ['11 kw', '11kw'],
	15: ['15 kw', '15kw'],
	18.5: ['18.5 kw', '18.5kw', '18,5 kw', '18,5kw'],
	22: ['22 kw', '22kw']
}

dict_inlet = {
	'Inlet spigot': ['inlet spig', '.spig', '. spig'],
	'Inlet rain prot.': ['in. rain', 'in.rain', 'inlet rain'],
	'Damper inlet': ['damper', 'dumper', 'regist'],
	'Blind panel inlet': ['blind']
}


dict_power_switch = {
	1: ['interr', 'mount.int', 'mount. int']
}

dict_pressure = {
	1: ['press', 'préss', 'diff']
}

dict_roof = {
	1: ['toit', 'roof']
}

dict_outlet_rain = {
	1: ['outlet rain', 'out.rain', 'out. rain', 'visie']
}

dict_vfd = {
	1: ['freq', 'conv']
}

dict_filter = {
	1: ['filt']
}

dict_double_skin = {
	1: ['doubl', 'skin']
}

# Create a list with all the dictionaries to loop
dicts = [dict_type, dict_size, dict_speed, dict_motor, dict_inlet,
dict_power_switch, dict_pressure, dict_roof, dict_outlet_rain, dict_vfd, dict_filter, dict_double_skin]
new_val, new_list = [], []

# Function to get the translation from quotation to template
for row in df['Model']:
	for dict in dicts:
		# Reset the value to ensure the pass for each dictionary
		x = ''
		new_val.append(x)
		for key, value in sorted(dict.items()):		
			# Only use that dictionary if there is no value found
			if x == '':
				for item in value:
					if item in row.lower():
						# Add the key found to the list
						new_val.pop()
						new_val.append(key)
						# Once found the value, stop the loop
						x = 1
						break
			# If the value if found, stop looping in that dictionary
			else:
				break
	new_list.append(new_val)
	new_val = []

# Check the standard configuration of the KPB
for item in new_list:
	type = item[0]
	inlet = item [4]
	if (type == 'KPB') & (inlet == ''):
		item[4] = 'Inlet spigot'

# Create dataframe
result = pd.DataFrame(new_list, columns = col)

# Add Qty and reorder the columns
result['Qty'] = df['Qty']
col.insert(5, 'Qty')
result = result[col]

# Export to Excel
name = 'Box-fans result.xlsx'
writer = pd.ExcelWriter(name)
result.to_excel(writer, index = False)
writer.save()