#!/usr/local/bin/python3 
#
# This script will get the membership list from Leader and Clerk Resources
# and will then check a Google Sheet to determine when the member last spoke.
# A prioritized list of members will be  updated on a Google Sheet for
# potential speakers
#

from __future__ import print_function
import httplib2
import os
import pathlib
import datetime

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

from lcr import API as LCR


try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Sacrament Speakers'


# constant
never = datetime.date(1900, 1, 1)
 

class google_sheet:
    def __init__(self):
        self.credentials = self.get_credentials()
        self.http = self.credentials.authorize(httplib2.Http())
        self.discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
        self.service = discovery.build('sheets', 'v4', http=self.http,
                              discoveryServiceUrl=self.discoveryUrl)

        self.spreadsheetId = '1F1tNtKisqUmovpgPi-ZaD8IvicewqS6YgIvOYpeTWKg'  #MASTER SPREADSHEET
        #self.spreadsheetId = '1se6pNb-PzyGhfeZRqrzZZGMYv2_7nypOO35vzFBeKL8'   #JUSTIN'S COPY

    def get_credentials(self):
        """Gets valid user credentials from storage.

        If nothing has been stored, or if the stored credentials are invalid,
        the OAuth2 flow is completed to obtain the new credentials.

        Returns:
            Credentials, the obtained credential.
        """
        home_dir = os.path.expanduser('~')
        credential_dir = os.path.join(home_dir, '.credentials')
        if not os.path.exists(credential_dir):
            os.makedirs(credential_dir)
        credential_path = os.path.join(credential_dir,
                                       'sheets.googleapis.com-sacrament_speakers.json')

        store = Storage(credential_path)
        credentials = store.get()
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
            flow.user_agent = APPLICATION_NAME
            if flags:
                credentials = tools.run_flow(flow, store, flags)
            else: # Needed only for compatibility with Python 2.6
                credentials = tools.run(flow, store)
            print('Storing credentials to ' + credential_path)
        return credentials

    def clear_potential_speakers(self, group_name):
        # Clear out the previous Potential Adult Speakers
        if group_name == 'Adult' or group_name == 'Youth':
            rangeName = 'Potential ' + group_name + ' Speakers!A2:C'
            
            result = self.service.spreadsheets().values().clear(
                spreadsheetId=self.spreadsheetId, range=rangeName, body={}).execute()
        else:
            print('ERROR: provided invalid group to clear_potential_speakers(): ' + group_name)
         
    def get_speakers_and_dates(self, group_name):
        speakers = {}
        if group_name == 'Adult' or group_name == 'Youth':
            rangeName = 'Sacrament ' + group_name + ' Speaker!A3:D'
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheetId, range=rangeName).execute()
            values = result.get('values', [])
                        
            if not values:
                print('No data found.')
            else:
                if group_name == 'Adult':
                    date_col = 2
                else:
                    date_col = 3
                
                for row in values:
                    if len(row) > date_col:
                        month, day, year = row[date_col].split('/')
                        name = row[0]
                        last_talk_date = datetime.date(int(year), int(month), int(day))
                        speakers[name] = last_talk_date
        else:
            print('ERROR: provided invalid group to get_speakers_and_dates(): ' + group_name)
            
        return speakers
    
    def write_potential_speakers(self, potential, memberinfo, group_name):
        if group_name == 'Adult' or group_name == 'Youth':
            rangeName = 'Potential ' + group_name + ' Speakers!A2:C'
            cell_values = []
            for name, date in potential:
                if name in memberinfo:
                    info = memberinfo[name]
                    if date == never:
                        date = 'NEVER'

                    row = [str(name), str(info['phone']), str(date)]
                    cell_values.append(row)
            request_body = {
                'values' : cell_values
            }
            request = self.service.spreadsheets().values().update(
                spreadsheetId=self.spreadsheetId, range=rangeName,
                valueInputOption='USER_ENTERED', body=request_body)
            response = request.execute()
        else:
            print('ERROR: provided invalid group to write_potential_speakers(): ' + group_name)
            
def main():

    lds_username = ''
    lds_password = ''
    lds_unit_number = ''

    from pathlib import Path
    home = str(Path.home())
    
    file = open(home + '/.lds', "r")
    for line in file:
        if line.startswith('LDS_USER'):
            lds_username = line.replace('LDS_USER=', '').strip()
        elif line.startswith('LDS_PASSWORD'):
            lds_password = line.replace('LDS_PASSWORD=', '').strip()
        elif line.startswith('UNIT_NUMBER'):
            lds_unit_number = line.replace('UNIT_NUMBER=', '').strip()

    # download the ward list and store it in a dictionary using their
    # name as the key.  this will make it easier to retrieve later
    lcr = LCR(lds_username, lds_password, lds_unit_number)
    members_list=lcr.member_list()
    members = {}
    for m in members_list:
        name = m['name']
        members[name] = m

    sheet = google_sheet()
    # clear out the google sheet for our potential speaker list
    sheet.clear_potential_speakers('Adult')
    sheet.clear_potential_speakers('Youth')

    # get the list of speakers and the dates they gave talks. 
    adult_speakers = sheet.get_speakers_and_dates('Adult')
    youth_speakers = sheet.get_speakers_and_dates('Youth')
    
    # remove any speakers that are no longer in the ward
    adult_speakers = {key: value for key, value in adult_speakers.items()
                if key in members}
    youth_speakers = {key: value for key, value in youth_speakers.items()
                if key in members}

    # The blacklist allows us to filter out members from the potential speaker
    # list like those who are in-active.
    blacklist = []
    blackfile = open('blacklist.txt', 'r')
    for i in blackfile:
        blacklist.append(i.strip())

    # now we need to iterate through the membership directory to find people who
    # have never given a talk
    for name, memberinfo in members.items():
        if name not in blacklist:
            # if they are an adult, and they're not in our list of speakers,
            # then add them to the potential speakers list and mark them as
            # never having spoken.
            if memberinfo['isAdult']:
                if name not in adult_speakers:
                    adult_speakers[name] = never
            elif memberinfo['actualAge'] > 12 and memberinfo['actualAge'] < 18:
                if name not in youth_speakers:
                    youth_speakers[name] = never

    # sort the potential speaker list first by date of last time they gave
    # a talk, and then alphabetical after that
    adult_potential_speaker_list = sorted(adult_speakers.items(), key=lambda x: (x[1], x[0]))
    youth_potential_speaker_list = sorted(youth_speakers.items(), key=lambda x: (x[1], x[0]))

    # write that list of potential speakers to the google sheet
    sheet.write_potential_speakers(adult_potential_speaker_list, members, 'Adult')
    sheet.write_potential_speakers(youth_potential_speaker_list, members, 'Youth')
    
      
if __name__ == '__main__':
    main()
