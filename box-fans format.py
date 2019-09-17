import pandas as pd

excel_file = 'Box-fans.xlsx'
df = pd.read_excel(excel_file)

columns = ['Type', 'Size', 'Motor type', 'Motor Power', 'Qty',
'Power Switch', 'Pressure Switch', 'Roof', 'Outlet Rain Prot.', 'VFD',
'Filter', 'Double Skin Panels']

dict_type = {
	'KPB': ['kpb', 'kp'],
	'KVBF': ['kvb', 'kvbf', 'kvb-f', 'kvbf/f']
}

dict_size = {
	'9': ['09/09', '9/9', '09-09', '09/09'],
	'10': ['10/10', '10-10'],
	'12': ['12/12', '12-12'],
	'15': ['15/15', '15-15'],
	'18': ['18/18', '18-18'],
	'20': ['20/20', '20-20'],
	'22': ['22/22', '22-22'],
	'25': ['25/25', '25-25'],
	'30': ['30/28', '30-28']
}

dict_speed = {
	'One speed': ['1 speed', '1speed'],
	'Two speeds': ['2 speeds', '2speeds']
}

dict_motor = {
	'0.25': ['0.25 kw', '0.25kw', '0,25 kw', '0,25kw'],
	'0.37': ['0.37 kw', '0.37kw', '0,37 kw', '0,37kw'],
	'0.55': ['0.55 kw', '0.55kw', '0,55 kw', '0,55kw'],
	'0.75': ['0.75 kw', '0.75kw', '0,75 kw', '0,75kw'],
	'1.1': ['1.1 kw', '1.1kw', '1,1 kw', '1,1kw'],
	'1.5': ['1.5 kw', '1.5kw', '1,5 kw', '1,5kw'],
	'2.2': ['2.2 kw', '2.2kw', '2,2 kw', '2,2kw'],
	'3': ['3 kw', '3kw'],
	'4': ['4 kw', '4kw'],
	'5.5': ['5.5 kw', '5.5kw', '5,5 kw', '5,5kw'],
	'7.5': ['7.5 kw', '7.5kw', '7,5 kw', '7,5kw'],
	'11': ['11 kw', '11kw'],
	'15': ['15 kw', '15kw'],
	'18.5': ['18.5 kw', '18.5kw', '18,5 kw', '18,5kw'],
	'22': ['22 kw', '22kw']
}

dict_power_switch = {
	'1': ['interr']
}

dict_filter = {
	'1': ['filt']
}

new_val = []
row_01 = 'KPB 25/25-11 kw-517 rpm-1 Speed + Filtre 25/25 + Interrup. KPB 1V >7,5<=11kw'
dicts = [dict_type, dict_size, dict_speed, dict_motor, dict_power_switch, dict_filter]

for dict in dicts:
	# Reset the value to ensure the pass for each dictionary
	x = 0
	for key, value in dict.items():		
		# Only use that dictionary if there is no value found
		if x == 0:
			for item in value:
				print('\n')
				print('Checking', item, 'at', key)
				if item in row_01.lower():
					# Add the key found to the list
					print('Item found!', key)
					new_val.append(key)
					# Once found the value, stop the loop
					x = 1
					break
		# If the value if found, stop looping in that dictionary
		else:
			break

print(new_val)