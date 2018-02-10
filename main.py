import math, sys, os, string, csv
import ConfigParser
import json
import flask
import httplib2

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

	#Compare all of the category keys to the comments string to categorize the payment
	def check(self, comments, ammount):
		category = 'Unknown'
		for cat in self.type.categories.iterkeys():
			for key in self.type.categories[cat]:
				if comments.find(key) > 0:
					category = cat
					break
		
		self.cat_totals[category] += ammount
				
		
class incometype():
	
	def __init__(self, name, keywords):
		self.name = name
		self.total = 0
		self.categories = {}
		
		for key in keywords.iterkeys():
			self.categories[key] = keywords[key]
		
#Parse the settings ini and create our tennants
configfile = open(os.path.join('settings.json'))
config = json.loads(configfile.read())

#Establish our spreadsheet ID for posting to google drive
spreadsheet = config['SpreadsheetID']

for type in config['IncomeTypes'].iterkeys():
	income[type] = incometype(type, config['IncomeTypes'][type])

print 'Income Types'
print income

for payment in config['Payments'].iterkeys():
	dict = config['Payments'][payment]
	payments[payment] = account(payment, dict['Accounts'], income[dict['Type']])

print 'Payments'
print payments

#The main loop for the import step, go over every line in the csv and figure out what to do with them
def parse(input):
	csvfile = csv.reader(input)
	
	initialized = False
	
	for row in csvfile:
		if initialized == False:
			initialized = True
			#Could set up an iterator here to dynamically detect data sorting, but we use a specific format for input
			pass
		
		#For now, all we care about is the deposits
		if row[0] == 'Deposit' or row[0] == 'Bill Payment' or row[0] == 'Direct Credit':
			#Match a payment up to a given tenant
			entry = 'bork'
			for payment in payments.itervalues():
				for account in payment.accounts:
					if row[1] == account:
						entry = payment
						break
			
			#If we have a matching account, sort the payment
			if entry != 'bork':
				notes = str(row[2]+' '+row[3]+' '+row[4]).lower()
				entry.check(notes, float(row[5]))
	
	#Set up headers in our spreadsheet
	range = 'Sheet1!A1:E'
	row = 1
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
				for total in subtype.cat_totals.itervalues():
					subvalues.append(str(total))
				values.append(subvalues)
				print 'added an account'
				print subvalues
		
		#Line break, in spreadsheet
		values.append([])
				
		range = 'Sheet1!A'+str(row)+':E'
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