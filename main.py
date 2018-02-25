import math, sys, os, string, csv
import ConfigParser
import json
import flask
import httplib2

import time
from datetime import date
from datetime import timedelta

import google_auth
from apiclient import discovery

#flow = google_auth_oauthlib.flow.Flow.from_client_secrets_file(os.path.join('client_secret.json'), scopes=['https://www.googleapis.com/auth/spreadsheets'])
#flow.redirect_uri = 'http://localhost:8080'
#authorization_url, state = flow.authorization_url(access_type='offline', include_granted_scopes='true')

#Set up our filepaths
main_dir = sys.path[0]

#Define a dictionary for tenants, we'll need this later
income = {}
payments = {}
expenditures = []

#This is a class used to hold data from 
class account():

	def __init__(self, name, accounts, incometype):
		self.name = name
		self.type = incometype
		
		self.accounts = []
		for account in accounts:
			self.accounts.append(account)
		
		self.cat_totals = {}
		for cat in self.type.categories.iterkeys():
			self.cat_totals[cat] = 0.0
		
		self.cat_totals['Unknown'] = 0.0
		
		#set up the first 'active period'
		self.ap = 0
		
		self.periods = {}
		self.periods[self.ap] = {}
		
		for cat in self.type.categories.iterkeys():
			self.periods[self.ap][cat] = 0.0
		
		self.periods[self.ap]['Unknown'] = 0.0

	#Compare all of the category keys to the comments string to categorize the payment
	def check(self, comments, ammount):
		category = 'Unknown'
		for cat in self.type.categories.iterkeys():
			for key in self.type.categories[cat]:
				if comments.find(key) > 0:
					category = cat
					break
		
		self.cat_totals[category] += ammount
		self.periods[self.ap][category] += ammount
		
	def new_period(self):
		self.ap +=1
		self.periods[self.ap] = {}
		
		for cat in self.type.categories.iterkeys():
			self.periods[self.ap][cat] = 0.0
		
		self.periods[self.ap]['Unknown'] = 0.0
		
class incometype():
	
	def __init__(self, name, keywords):
		self.name = name
		self.total = 0
		self.categories = {}
		
		for key in keywords.iterkeys():
			self.categories[key] = keywords[key]

class expenditure():

	def __init__(self, name, accounts):
		self.name = name
		
		self.accounts = []
		for account in accounts:
			self.accounts.append(account)
		
		self.total = 0.0
		
		#set up the first 'active period'
		self.ap = 0
		
		self.periods = {}
		self.periods[self.ap] = 0.0
		
	def new_period(self):
		self.ap +=1
		self.periods[self.ap] = 0.0
		
class transaction_unknown():
	
	def __init__(self):
	
		self.transactions = {}
		self.total = 0.0
	
	def sort(self, account, ammount):
		
		if account in self.transactions.iterkeys():
			self.transactions[account] += ammount
		else:
			self.transactions[account] = ammount
		
		self.total += ammount
		
#Parse the settings ini and create our tennants
configfile = open(os.path.join('settings.json'))
config = json.loads(configfile.read())

#Establish our spreadsheet ID for posting to google drive
spreadsheet = config['SpreadsheetID']

for type in config['IncomeTypes'].iterkeys():
	income[type] = incometype(type, config['IncomeTypes'][type])

unknown_income = transaction_unknown()

print 'Income Types'
print income

for payment in config['Payments'].iterkeys():
	dict = config['Payments'][payment]
	payments[payment] = account(payment, dict['Accounts'], income[dict['Type']])
	
unknown_expenditure = transaction_unknown()

print 'Payments:'
print payments

for ex in config['ExpenditureTypes'].iterkeys():
	dict = config['ExpenditureTypes'][ex]
	expenditures.append(expenditure(ex, dict['Accounts']))
	
print 'ExpenditureTypes:'
for item in expenditures:
	print item.name

#The main loop for the import step, go over every line in the csv and figure out what to do with them
def parse(input):
	csvfile = csv.reader(input)
	
	initialized = False
	next_date = None
	period_count = 0
	
	for row in csvfile:
		if initialized == False:
			initialized = True
			#Could set up an iterator here to dynamically detect data sorting, but we use a specific format for input
			continue
		
		#Time management, gets the date of the current transaction and compares to the time step
		ds = row[6].split('/')
		current = date(int(ds[2]), int(ds[1]), int(ds[0]))
		if next_date == None:
			next_date = current - timedelta(days=7)
			
		#If we're past the time step, start a new period for all defined payments
		elif current <= next_date:
			for payment in payments.itervalues():
				payment.new_period()
			for ex in expenditures:
				ex.new_period()
				
			next_date = current - timedelta(days=7)
			period_count += 1
			
			print 'week '+str(period_count)+', next week is:'
			print next_date
		
		#Is this income? If so sort it through our income filters
		if row[0] == 'Deposit' or row[0] == 'Bill Payment' or row[0] == 'Direct Credit':
			#Match a payment up to a given tenant
			entry = 'bork'
			for payment in payments.itervalues():
				for account in payment.accounts:
					if row[1] == account:
						entry = payment
						break
				if entry != 'bork':
					break
			
			#If we have a matching account, sort the payment
			if entry != 'bork':
				notes = str(row[2]+' '+row[3]+' '+row[4]).lower()
				entry.check(notes, float(row[5]))
			else:
				unknown_income.sort(row[1], float(row[5]))
		
		#Is this an expenditure? If so sort it through our expenditure filters
		else:
			entry = 'bork'
			for ex in expenditures:
				for account in ex.accounts:
					if account in row[1]:
						entry = ex
						break
				if entry != 'bork':
					break
					
			if entry != 'bork':
				entry.total += float(row[5])
				entry.periods[entry.ap] += float(row[5])
			else:
				unknown_expenditure.sort(row[1], float(row[5]))
			
			
	
	#Set up headers in our spreadsheet
	range = 'Sheet1!A1:E'
	row = 1
	values = [['Unknown Income','Ammount']]
	
	for account in unknown_income.transactions.iterkeys():
		subvalues = []
		subvalues.append(account)
		subvalues.append(unknown_income.transactions[account])
		values.append(subvalues)
	
	values.append(['TOTAL', str(unknown_income.total)])
	
	#Line break, in spreadsheet
	values.append([])

	range = 'Sheet1!A'+str(row)+':E'
	output_data(range, values)
	row += len(values)
	
	for item in income.iterkeys():
		#Set up major income type header
		values = [[item]]
		
		subvalues = []
		subvalues.append('Account')
		for cat in income[item].categories.iterkeys():
			subvalues.append(cat)
		subvalues.append('Unknown')
		values.append(subvalues)
		
		#Go through and print out all subype data
		for subtype in payments.itervalues():
			subvalues = []
			if subtype.type == income[item]:
				subvalues.append(subtype.name)
				for cat in income[item].categories.iterkeys():
					subvalues.append(str(subtype.cat_totals[cat]))
				subvalues.append(str(subtype.cat_totals['Unknown']))
				values.append(subvalues)
				print 'added an account:'+subtype.name
		
		i = 0
		while i <= period_count:
			
			#Line break, in spreadsheet
			values.append([])
			values.append(['Week '+str(i)])
			
			for subtype in payments.itervalues():
				subvalues = []
				if subtype.type == income[item]:
					subvalues.append(subtype.name)
					for cat in income[item].categories.iterkeys():
						subvalues.append(str(subtype.periods[i][cat]))
					subvalues.append(str(subtype.periods[i]['Unknown']))
					values.append(subvalues)
					
			i += 1
		
		#Line break, in spreadsheet
		values.append([])
				
		range = 'Sheet1!A'+str(row)+':E'
		output_data(range, values)
		row += len(values)
		
	#EXPENDITURE OUTPUT
		
	#Set up headers in our spreadsheet
	range = 'Sheet1!H1:M'
	row = 1
	values = [['Unknown Expenditure','Ammount']]
	
	for account in unknown_expenditure.transactions.iterkeys():
		subvalues = []
		subvalues.append(account)
		subvalues.append(unknown_expenditure.transactions[account])
		values.append(subvalues)
	
	values.append(['TOTAL', str(unknown_expenditure.total)])
	
	#Line break, in spreadsheet
	values.append([])
	
	values.append(['Known Expenditures'])
	
	subvalues = []
	for ex in expenditures:
		subvalues.append(ex.name)
	values.append(subvalues)
	subvalues = []
	for ex in expenditures:
		subvalues.append(ex.total)
	values.append(subvalues)
	
	#Do weekly expenditures
	i = 0
	while i <= period_count:
		
		#Line break, in spreadsheet
		values.append([])
		values.append(['Week '+str(i)])
		
		subvalues = []
		for ex in expenditures:
			subvalues.append(ex.name)
		values.append(subvalues)
		subvalues = []
		for ex in expenditures:
			subvalues.append(ex.periods[i])
		values.append(subvalues)
				
		i += 1
	
	#Output the whole expenditures block
	range = 'Sheet1!H'+str(row)+':M'
	output_data(range, values)
	row += len(values)
	
	
		

def output_data(target_range, data):
	credentials = google_auth.get_credentials()
	http = credentials.authorize(httplib2.Http())
	discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?' 'version=v4')
	service = discovery.build('sheets', 'v4', http=http, discoveryServiceUrl=discoveryUrl)
	
	values = data
	body = {'values':values}

	spreadsheet_id = spreadsheet
	range_name = target_range
	value_input_option = 'RAW'
	insert_data_option = 'OVERWRITE'
	
	result = service.spreadsheets().values().update(
	spreadsheetId=spreadsheet_id, range=range_name,valueInputOption=value_input_option, body=body
	).execute()
	
	print('{0} cells updated.'.format(result.get('updatedCells')));
	
	#result = service.spreadsheets().values().get(spreadsheetId=spreadsheetId, range=rangeName).execute()
	#insert = service.spreadsheets().values().append(spreadsheetId=spreadsheetId, range=rangeName, valueInputOption=value_input_option, insertDataOption=insert_data_option, body=value_range_body)
	#insert.execute()
	#values = result.get('values', [])

#	if not values:
#		print('No data found.')
#	else:
#		print('Name, Major:')
#		for row in values:
			# Print columns A and E, which correspond to indices 0 and 4.
#			print('%s, %s' % (row[0], row[4]))

			
#Wait for the user to provide a csv path for us to use
data = raw_input('Give me a CSV:')
assert os.path.exists(data), "I did not find the file at, "+str(data)
file = open(data,'r+')
parse(file)
file.close()