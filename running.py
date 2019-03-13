import os
import sys
import json
import time

from models import PExcelModel, GTASSModel
from datetime import datetime

def getKonsorsiumName(i):
	konsorsium_list = ['KAI']
	return konsorsium_list[i-1]

print('**************************************')
print('* Konsorsium yang di pilih           *')
print('**************************************')
print('* 1. KAI                             *')
print('**************************************')

konsorsium_choice = input('> ')
konsorsium_choice = int(konsorsium_choice)

if not konsorsium_choice:
	print('result : Konsorsium harus di pilih tidak boleh kosong !')
	sys.exit()

if konsorsium_choice not in [1]:
	print('result : Konsorsium yang di pilih tidak ada !')
	sys.exit()

konsorsium_name = getKonsorsiumName(konsorsium_choice)

if not konsorsium_name:
	print('result : Konsorsium yang di pilih kosong !')
	sys.exit()

file_name = input('Nama File : \n')

if not file_name:
	print('result : Nama file tidak boleh kosong !')
	sys.exit()

file_dir = 'file/' + file_name

is_file_exists = True if os.path.exists(file_dir) else False
if is_file_exists == False:
	print('result : Nama file (' + konsorsium_name + ') tidak ada !')
	sys.exit()

pexcel = PExcelModel(file_dir)
record = pexcel.record
title = pexcel.title

if not title:
	print('result : format tidak mendukung !')
	sys.exit()

if title[1] != 'Date' or title[4] != 'Booking Code' or title[15] != 'Ticket Number':
	print('result : Format title tidak mendukung !');
	sys.exit() 

if title[14] != 'Airline':
	print('result : format title tidak mendukung !')
	sys.exit();

print('|'.join(title))
for list in record:
	data = []
	for i in title:
		data.append(str(list[i]))
	print('|'.join(data))

confirm = input('Simpan data ? (Y/N)\n')

if confirm == 'Y' or confirm == 'y':
	config = open('config/gtass.txt', 'r')
	config = config.readlines(0)
	params = json.loads(config[0])
	
	gtass = GTASSModel(params['url'], params['username'], params['password'])
	gtass_login = gtass.login(is_force=False)
	customer = gtass.getCustomer()
	
	customer_code = input('Masukan Code Customer (AGENT VERSA OR MITRA)\n')
	if not customer_code:
		print('result : Code Customer kosong !')
		gtass.logout()
		sys.exit()
	
	if len(customer_code) != 8:
		print('result : Code Customer harus sesuai (8 karakter) !')
		gtass.logout()
		sys.exit()
	
	customer_data = gtass.getCustomerData(customer_code)
	
	if not customer_data or 'code' not in customer_data:
		print('result : Code Customer tidak ada !')
		gtass.logout()
		sys.exit()
	
	supplier_code = input('Masukan Code Supplier\n')
	if not supplier_code:
		print('result : Code Supplier kosong !')
		gtass.logout()
		sys.exit()
	
	if len(supplier_code) != 6:
		print('result : Code Supplier harus sesuai (6 karakter) !')
		gtass.logout()
		sys.exit()
	
	supplier_data = gtass.getSupplierData(supplier_code)
	
	if not supplier_data or 'code' not in supplier_data:
		print('result : Code Supplier tidak ada !')
		gtass.logout()
		sys.exit()
	
	for list in record:
		result = ''
		
		data = []
		for i in title:
			data.append(str(list[i]))
		
		valid = False
		
		if konsorsium_choice == 1:
			if list['Airline'] != 'Trigana API':
				result = ('|'.join(data)) + '(' + konsorsium_name + ') Must Trigana API in column airline'
				time.sleep(5)
			else:
				valid = True
		
		if valid == True:
			air_code = ticket_three_code = ''
			flight_code = class_code = ''
			is_domestic = 1
			oneway = 1
			qty = 1
			
			if konsorsium_choice == 1:
				air_code = 'A00019'
				ticket_three_code = '000'
			
			route_list = list['Route'].split('-')
			time_depart = datetime.fromisoformat(str(list['Time Depart'])).timestamp()
			time_depart = int(time_depart)
			time_arrive = datetime.fromisoformat(str(list['Time Arrive'])).timestamp()
			time_arrive = int(time_arrive)
			city_depart = route_list[0]
			city_arrive = route_list[1]
			contact_list = list['Contact'].split('.')
			contact_title = contact_list[0].upper()
			contact_name = contact_list[1].strip()
			
			# date_time = datetime.fromtimestamp(time_depart)
			# date_time.strftime('%Y-%m-%d %H:%i:%s')
			
			params = {}
			params['booking_code'] = list['Booking Code']
			params['booking_code'] = params['booking_code'][:7]
			params['booking_date'] = datetime.fromisoformat(str(list['Booking Date'])).timestamp()
			params['booking_date'] = int(params['booking_date'])
			params['issued_date'] = datetime.fromisoformat(str(list['Date'])).timestamp()
			params['issued_date'] = int(params['issued_date'])
			params['oneway'] = oneway
			params['qty'] = qty
			params['is_domestic'] = is_domestic
			params['air_code'] = air_code
			params['contact_title'] = contact_title
			params['contact_name'] = contact_name
			params['ticket_three_code'] = ticket_three_code
			params['ticket_number'] = (list['Ticket Number'].lstrip(ticket_three_code)).strip()
			params['supplier_data'] = supplier_data
			
			# INCLUDE TICKET SCHEDULE
			params['schedule'] = {}
			params['schedule']['time_depart'] = time_depart
			params['schedule']['time_arrive'] = time_arrive
			params['schedule']['city_depart'] = city_depart
			params['schedule']['city_arrive'] = city_arrive
			params['schedule']['class_code'] = class_code
			params['schedule']['class_type'] = ''
			params['schedule']['flight_code'] = flight_code
			
			# INCLUDE TICKET PRICE
			params['fare'] = {}
			params['fare']['basic'] = list['Basic']
			params['fare']['tax'] = list['Tax']
			params['fare']['total'] = list['Publish']
			# jika ada real_nta = Real NTA maka dapat komisi di GTASS
			params['fare']['real_nta'] = list['Real NTA']
			
			remark1 = params['booking_code'] + ' ' + konsorsium_name + ' ' + params['ticket_number']
			remark1 = remark1[:100]
			
			is_already_invoice = True
			is_already_invoice = gtass.isAlreadyInvoice(params['issued_date'], remark1)
			
			if is_already_invoice == True:
				result = ('|'.join(data)) + ' Is Already Invoice'
				time.sleep(5)
			else:
				is_already_tiket = True
				is_already_tiket = gtass.isAlreadyResTicketTrain(params['issued_date'], params['booking_code'])
				
				if is_already_tiket == True:
					result = ('|'.join(data)) + ' Is Already Ticket'
					time.sleep(5)
				else:
					gtass.addReservationTicketTrain(params)
					result = ('|'.join(data)) + ' Done'
					time.sleep(5)
		
		print(result)
	
	print('running')
	gtass.logout()
else :
	sys.exit()
