import pandas as pd

excel_file = 'Box-fans.xlsx'
df = pd.read_excel(excel_file)

columns = ['Type', 'Size', 'Motor type', 'Motor Power', 'Qty',
'Power Switch', 'Pressure Switch', 'Roof', 'Outlet Rain Prot.', 'VFD',
'Filter', 'Double Skin Panels']

dict_type = {
	'KPB': ['KPB', 'KP'],
	'KVBF': ['KVBF', 'KVB-F', 'KVB/F']
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
	'One speed': ['1 speed', '1 Speed'],
	'Two speeds': ['2 speeds', '2 Speeds']
}

dict_motor = {
	'0.25': ['0.25 KW', '0.25 kW', '0.25KW', '0.25kW', '0,25 KW', '0,25 kW', '0,25KW', '0,25kW'],
	'0.37': ['0.37 KW', '0.37 kW', '0.37KW', '0.37kW', '0,37 KW', '0,37 kW', '0,37KW', '0,37kW'],
	'0.55': ['0.55 KW', '0.55 kW', '0.55KW', '0.55kW', '0,55 KW', '0,55 kW', '0,55KW', '0,55kW'],
	'0.75': ['0.75 KW', '0.75 kW', '0.75KW', '0.75kW', '0,75 KW', '0,75 kW', '0,75KW', '0,75kW'],
	'1.1': ['1.1 KW', '1.1 kW', '1.1KW', '1.1kW', '1,1 KW', '1,1 kW', '1,1KW', '1,1kW'],
	'1.5': ['1.5 KW', '1.5 kW', '1.5KW', '1.5kW', '1,5 KW', '1,5 kW', '1,5KW', '1,5kW'],
	'2.2': ['2.2 KW', '2.2 kW', '2.2KW', '2.2kW', '2,2 KW', '2,2 kW', '2,2KW', '2,2kW'],
	'3': ['3 KW', '3 kW', '3KW', '3kW'],
	'4': ['4 KW', '4 kW', '4KW', '4kW'],
	'5.5': ['5.5 KW', '5.5 kW', '5.5KW', '5.5kW', '5,5 KW', '5,5 kW', '5,5KW', '5,5kW'],
	'7.5': ['7.5 KW', '7.5 kW', '7.5KW', '7.5kW', '7,5 KW', '7,5 kW', '7,5KW', '7,5kW'],
	'11': ['11 KW', '11 kW', '11KW', '11kW'],
	'15': ['15 KW', '15 kW', '15KW', '15kW'],
	'18.5': ['18.5 KW', '18.5 kW', '18.5KW', '18.5kW', '18,5 KW', '18,5 kW', '18,5KW', '18,5kW'],
	'22': ['22 KW', '22 kW', '22KW', '22kW'],
}

dict_power_switch = {
	'1': ['Interr', 'interr']
}

dict_filter = {
	'1': ['Filt', 'filt']
}

row_01 = 'KPB 25/25-11 KW-517 rpm-1 Speed + Filtre 25/25 + Interrup. KPB 1V >7,5<=11Kw'
dicts = [dict_type, dict_size, dict_speed, dict_motor, dict_power_switch, dict_filter]

for dict in dicts:
	for main_field in dict:
		list_main_field = dict[main_field]
		value = '0'
		for item in list_main_field:
			if item in row_01:
				print('\n')
				print('Main field is', main_field)
				break


#print(keys['Type'])

#for row in df.Model:
	#print(row)