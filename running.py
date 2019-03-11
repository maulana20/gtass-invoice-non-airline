import os
import sys

from models import PExcelModel

file_name = input('Nama File : \n')

if not file_name:
	print('result : Nama file tidak boleh kosong !')
	sys.exit()

file_dir = 'file/' + file_name

is_file_exists = True if os.path.exists(file_dir) else False
if is_file_exists == False:
	print('result : Nama file tidak ada !')
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
for data in record:
	_data = []
	for i in title:
		_data.append(str(data[i]))
	print('|'.join(_data))

confirm = input('Simpan data ? (Y/N)\n')

if confirm == 'Y' or confirm == 'y':
	print('running')
else :
	print('not running')
