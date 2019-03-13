from datetime import datetime
from datetime import time

from urllib.request import urlopen
from bs4 import BeautifulSoup

import xlrd
import requests

import json


class PExcelModel():
	def __init__(self, file_dir):
		workbook = xlrd.open_workbook(file_dir)
		worksheet = workbook.sheet_by_index(0)
		
		num_rows = worksheet.nrows
		num_cols = worksheet.ncols
		
		title = []
		for curr_row in range(0, 1):
			for curr_col in range(0, num_cols, 1):
				data = worksheet.cell_value(curr_row, curr_col)
				title.append(data)
		self.title = title
		
		record = []
		for curr_row in range(1, num_rows, 1):
			column = {}
			for curr_col in range(0, num_cols, 1):
				data = worksheet.cell_value(curr_row, curr_col)
				
				if title[curr_col] in ['Date', 'Time', 'Booking Date', 'Time Depart', 'Time Arrive']:
					if title[curr_col] == 'Time':
						data = int(data * 24 * 3600)
						data = time(data//3600, (data%3600)//60, data%60)
					else:
						data = int(data) if type(data) is float else data
						data = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + data - 2)
					column[title[curr_col]] = data
				else:
					data = int(data) if type(data) is float else data
					column[title[curr_col]] = data
			record.append(column)
		
		self.record = record
	
	def record(self):
		return self.record
	
	def title(self):
		return self.title

class GTASSModel():
	def __init__(self, url, username, password):
		self.request = requests.Session()
		self.url = url
		self.username = username
		self.password = password
		
	def login(self, is_force=False):
		try:
			force = '/force' if is_force == True else ''
			self.username = self.username + force
			
			response = self.request.post(self.url + '/login', data = {'username': self.username, 'password': self.password})
			result = response.text
			
			bs = BeautifulSoup(result, 'html.parser')
			body = bs.find('li')
			body = str(body)
			body = body[(body.find(">") + 1):]
			matches = body[:body.find("<")]
			matches = matches.lower()
			matches = matches.replace('.', '')
			
			file_name = 'GTASSLogin.html'
			
			if matches == 'this user is logged on':
				self.login(is_force=True)
				file_name = 'GTASSForceLogin.html'
			
			fo = open('log/' + file_name, 'w')
			fo.write(result)
			fo.close()
		except:
			result = 'login failed !'
			
			fo = open('log/GTASSLogin.html', 'w')
			fo.write(result)
			fo.close()
		
		return
	
	def logout(self):
		response = self.request.get(self.url + '/logout', data = {})
		result = response.text
		
		fo = open('log/GTASSLogout.html', 'w')
		fo.write(result)
		fo.close()
		
		return
	
	def getCustomer(self):
		response = self.request.post(self.url + '/api/customer/list', data = {'search': '', 'take': 50, 'skip': 0, 'page': 1, 'pageSize': 50})
		result = response.text
		result = json.dumps(result)
		
		fo = open('log/GTASSCustomerList.html', 'w')
		fo.write(result)
		fo.close()
		
		return
	
	def getCustomerData(self, customer_code):
		response = self.request.post(self.url + '/api/customer/list', data = {'search': customer_code, 'take': 50, 'skip': 0, 'page': 1, 'pageSize': 50})
		result = response.text
		
		fo = open('log/GTASSCustomerData.html', 'w')
		fo.write(json.dumps(result))
		fo.close()
		
		res = json.loads(result)
		return res['data'][0]
	
	def getSupplierData(self, supplier_code):
		response = self.request.post(self.url + '/api/supplier/list', data = {'search': supplier_code, 'take': 50, 'skip': 0, 'page': 1, 'pageSize': 50})
		result = response.text
		
		fo = open('log/GTASSSupplierData.html', 'w')
		fo.write(json.dumps(result))
		fo.close()
		
		res = json.loads(result)
		return res['data'][0]
	
	def isAlreadyInvoice(self, date, remark1):
		date_time = datetime.fromtimestamp(date)
		date_time = date_time.strftime('%d %b %Y')
		
		response = self.request.post(self.url + '/api/supplier/list', data = {'search': date_time, 'take': 10, 'skip': 0, 'page': 1, 'pageSize': 100, 'voidStatus': False})
		result = response.text
		
		fo = open('log/GTASSInvoiceList.html', 'w')
		fo.write(json.dumps(result))
		fo.close()
		
		res = json.loads(result)
		for data in res['data']:
			if data['remark1'] == remark1:
				return True
		
		return False
	
	def isAlreadyResTicketTrain(self, date, booking_code):
		date_time = datetime.fromtimestamp(date)
		date_time = date_time.strftime('%d %b %Y')
		
		response = self.request.post(self.url + '/api/tickettrain-trans/list', data = {'search': date_time, 'take': 10, 'skip': 0, 'page': 1, 'pageSize': 100, 'voidStatus': False})
		result = response.text
		
		fo = open('log/GTASSReservationTicketTrainList.html', 'w')
		fo.write(json.dumps(result))
		fo.close()
		
		res = json.loads(result)
		for data in res['data']:
			if data['bookCode'] == booking_code:
				return True
		
		return False
	
	def addReservationTicketTrain(self, res):
		response = self.request.post(self.url + '/api/tickettrain-trans/list', data = {'search': '', 'take': 50, 'skip': 0, 'page': 1, 'pageSize': 50})
		result = response.text
		
		fo = open('log/GTASSReservationTicketTrainList.html', 'w')
		fo.write(json.dumps(result))
		fo.close()
		
		booking_date = datetime.fromtimestamp(res['booking_date'])
		booking_date = booking_date.strftime('%Y-%m-%d')
		
		issued_date = datetime.fromtimestamp(res['issued_date'])
		issued_date = issued_date.strftime('%Y-%m-%d')
		
		json_data = {
			'act': 'add',
			'adtQty': res['qty'],
			'airCode': res['air_code'],
			'bookBy': 'M0003',
			'bookCode': res['booking_code'],
			'bookDate': booking_date,
			'chdQty': 0,
			'currCode': 'IDR',
			'dom': 1,
			'infQty': 0,
			'issuedBy': 'M0003',
			'issuedDate': issued_date,
			'locationId': 1,
			'oneWay': res['oneway'],
			'referral': 'tickettrain-trans',
			'supCode': res['supplier_data']['code'],
			'tourCode': '',
			'type': 'DS'
		}
		print(json_data)
		
		response = self.request.post(self.url + '/api/tickettrain-trans/update?act=add', json = json_data)
		result = response.text
		
		fo = open('log/GTASSReservationTicketTrainAdd.html', 'w')
		fo.write(json.dumps(result))
		fo.close()
		
		_res = json.loads(result)
		if _res['success'] != True:
			fo = open('log/gtass_error.txt', 'a')
			fo.write('input tiket' + _res['message'])
			fo.close()
		else:
			fo = open('log/gtass_success.txt', 'a')
			fo.write('input tiket' + _res['message'])
			fo.close()
		
		ticket_code = _res['data']['id'];
		
		time_depart = datetime.fromtimestamp(res['schedule']['time_depart'])
		time_depart = time_depart.strftime('%H:%M')
		
		flight_date = datetime.fromtimestamp(res['schedule']['time_depart'])
		flight_date = flight_date.strftime('%Y-%m-%d')
		
		time_arrive = datetime.fromtimestamp(res['schedule']['time_arrive'])
		time_arrive = time_arrive.strftime('%H:%M')
		
		json_data = {
			'act': 'add',
			'depAirport': res['schedule']['city_depart'],
			'arrAirport': res['schedule']['city_arrive'],
			'depTime': time_depart,
			'arrTime': time_arrive,
			'classCode': res['schedule']['class_code'],
			'classType': res['schedule']['class_type'],
			'flightDate': flight_date,
			'flightNo': res['schedule']['flight_code'],
			'id': ticket_code,
			'idc': ticket_code,
			'locId': 1,
			'referral': 'tickettrain-trans'
		}
		print(json_data)
		
		response = self.request.post(self.url + '/api/tickettrain-trans/updateroute?act=add', json = json_data)
		result = response.text
		
		fo = open('log/GTASSReservationTicketTrainSchedule.html', 'w')
		fo.write(json.dumps(result))
		fo.close()
		
		_res = json.loads(result)
		if _res['success'] != True:
			fo = open('log/gtass_error.txt', 'a')
			fo.write('input schedule ' + _res['message'])
			fo.close()
		else:
			fo = open('log/gtass_success.txt', 'a')
			fo.write('input schedule ' + _res['message'])
			fo.close()
		
		json_data = {
			'act': 'add',
			'id': ticket_code,
			'basicFare': res['fare']['basic'],
			'iwjr': res['fare']['tax'],
			'publish': res['fare']['total'],
			'nta': res['fare']['real_nta'], #111000-5000
			'agentCom': res['fare']['total'] - res['fare']['real_nta'],
			'bsp': 0,
			'disc': 0,
			'discP': 0,
			'extraDisc': 0,
			'extraDiscP': 0,
			'fs': 0,
			'incentive': 0,
			'insurance': 0,
			'issueFee': 0,
			'agentComP': 0,
			'airportTax': 0,
			'ppn': 0,
			'ppnP': 0,
			'profit': 0,
			'serFee': 0,
			'disc': res['fare']['total'] - res['fare']['real_nta'], #lawan komisi
			'allowEditPurchase': True,
			'allowEditSales': True,
			'allowShowSales': True,
			'threeCode': res['ticket_three_code'], #990 => Lion, 
			'tickNo': res['ticket_three_code'] + res['ticket_number'], #by ticket
			'title': res['contact_title'], #by contact
			'firstName': res['contact_name'], #by contact
			'lastName': '',
			'locId': 1,
			'partTickNo': res['ticket_number'],
			'paxType': 'A',
			'referral': 'tickettrain-trans',
			'tourCode': ''
		}
		print(json_data)
		
		response = self.request.post(self.url + '/api/tickettrain-trans/updatedetail?act=add', json = json_data)
		result = response.text
		
		fo = open('log/GTASSReservationTicketTrainPrice.html', 'w')
		fo.write(json.dumps(result))
		fo.close()
		
		_res = json.loads(result)
		if _res['success'] != True:
			fo = open('log/gtass_error.txt', 'a')
			fo.write('input tiket ' + _res['message'])
			fo.close()
		else:
			fo = open('log/gtass_success.txt', 'a')
			fo.write('input tiket ' + _res['message'])
			fo.close()
	
	def addInvoice(self, res, customer_data, remark1):
		response = self.request.post(self.url + '/api/invoice/list', data = {'search': '', 'take': 50, 'skip': 0, 'page': 1, 'pageSize': 50})
		result = response.text
		
		fo = open('log/GTASSInvoiceGeneralList.html', 'w')
		fo.write(json.dumps(result))
		fo.close()
		
		issued_date = datetime.fromtimestamp(res['issued_date'])
		issued_date = issued_date.strftime('%Y-%m-%d')
		
		json_data = {
			'act': 'add', #MANDATORY
			'attn': '', #MANDATORY
			'code': '<AUTO>', #MANDATORY
			'currCode': 'IDR', #MANDATORY
			'custCode': customer_data['code'], #COA SUB AGENT VERSA
			'custName': customer_data['name'], #COA SUB AGENT VERSA
			'custPhone': customer_data['phone'], #COA SUB AGENT VERSA
			'date': issued_date,
			'dueDate': issued_date, #2018-08-20 DATE ISSUED
			'deposit': 0,
			'extraDisc': 0, #MANDATORY
			'isDpSubAgent': True, #CHOICE BY DEPOSIT AGENT
			'isWG': False, #CHOICE BY FIT
			'locationId': 1, #PUSAT
			'paxPaid': 0,
			'pph23': 0,
			'ppn': 0, #MANDATORY
			'prodCode': 'TICKD', #MANDATORY
			'rate': 1, #MANDATORY
			'referral': 'invoice', #MANDATORY
			'remark1': remark1, #DESKRIPSI
			'salesBy': 'M0003', #PUTUT
			'stamp': 0, #MANDATORY
			'taxNo': 0, #OPTIONAL
			'tourCode': '' #MANDATORY
		}
		
		response = self.request.post(self.url + '/api/tickettrain-trans/updatedetail?act=add', json = json_data)
		result = response.text
		
		fo = open('log/GTASSInvoiceGeneralAdd.html', 'w')
		fo.write(json.dumps(result))
		fo.close()
		
		_res = json.loads(result)
		if _res['success'] != True:
			fo = open('log/gtass_error.txt', 'a')
			fo.write('input invoice ' + _res['message'])
			fo.close()
		else:
			fo = open('log/gtass_success.txt', 'a')
			fo.write('input invoice ' + _res['message'])
			fo.close()
			
		invoice_code = _res['data']['code']
		created_date = _res['data']['createdDate']
		changed_date = _res['data']['changedDate']
		created_by = _res['data']['createdBy']
		changed_by = _res['data']['changedBy']
		
		response = self.request.post(self.url + '/api/ticket-trans/uninv-lists', data = {'curr': 'IDR', 'prodType': 'TICKD', 'searchBy': 'pnr', 'search': res['booking_code']})
		result = response.text
		
		fo = open('log/GTASSInvoiceGeneralSearch.html', 'w')
		fo.write(json.dumps(result))
		fo.close()
		
		_res = json.loads(result)
		ticket_id = _res[0]['id']
		
		json_data = {
			'act': 'add',
			'code': invoice_code,
			'dChecked': [ticket_id], #307
			'date': issued_date, #2018-08-21
			'locId': 1,
			'referral': 'invoice',
			'tourCode': '',
			'type': 'tick'
		}
		
		response = self.request.post(self.url + '/api/tickettrain-trans/updatedetail?act=add', json = json_data)
		result = response.text
		
		fo = open('log/GTASSInvoiceGeneralIncludeTicket.html', 'w')
		fo.write(json.dumps(result))
		fo.close()
		
		_res = json.loads(result)
		if _res['success'] != True:
			fo = open('log/gtass_error.txt', 'a')
			fo.write('include ticket ' + _res['message'])
			fo.close()
		else:
			fo = open('log/gtass_success.txt', 'a')
			fo.write('include ticket ' + _res['message'])
			fo.close()
