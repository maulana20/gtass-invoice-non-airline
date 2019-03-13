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
	
	def getSupplierData(self, supplier_code):
		response = self.request.post(self.url + '/api/supplier/list', data = {'search': 'S00008', 'take': 50, 'skip': 0, 'page': 1, 'pageSize': 50})
		result = response.text
		
		fo = open('log/GTASSupplierData.html', 'w')
		fo.write(json.dumps(result))
		fo.close()
		
		res = json.loads(result)
		return res['data'][0]
	
	def isAlreadyInvoice(self, date, remark1):
		date_time = datetime.fromtimestamp(date)
		date_time = date_time.strftime('%d %b %Y')
		print(date_time)

		response = self.request.post(self.url + '/api/supplier/list', data = {'search': date_time, 'take': 10, 'skip': 0, 'page': 1, 'pageSize': 100, 'voidStatus': False})
		result = response.text
		
		fo = open('log/GTASSReservationTicketList.html', 'w')
		fo.write(json.dumps(result))
		fo.close()
		
		res = json.loads(result)
		for data in res['data']:
			if data['remark1'] == remark1:
				return True
		
		return False
