from __future__ import print_function

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import dateutil.parser as parser
import base64
import email

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']


def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=creds)
        results = service.users().messages().list(userId='me', labelIds = ["CATEGORY_PERSONAL", "UNREAD"]).execute()
        messages = results.get('messages', [])
        final_list = []
        subject_list = []
        filtered_sub_list = []

        mssg_list = results['messages']
        print ("Total unread messages in Personal: ", str(len(mssg_list)))
        for mssg in mssg_list:
            temp_dict = {}
            m_id = mssg['id']
            message = service.users().messages().get(userId='me', id=m_id).execute()
            payld = message['payload']
            headr = payld['headers']
            #print(headr)

            for one in headr: # getting the Subject
                if one['name'] == 'Subject':
                    msg_subject = one['value']
                    #temp_dict['Subject'] = msg_subject
                    subject_list.append(msg_subject)
                    #print(msg_subject)
                    if "Dasani" in msg_subject:
                        temp_dict[m_id] = msg_subject
                        filtered_sub_list.append(temp_dict)
            else:
                pass

        print(filtered_sub_list)
        '''
            for two in headr: # getting the date
                if two['name'] == 'Date':
                    msg_date = two['value']
                    date_parse = (parser.parse(msg_date))
                    m_date = (date_parse.date())
                    temp_dict['Date'] = str(m_date)
            else:
                pass

            for three in headr: # getting the Sender
                if three['name'] == 'From':
                    msg_from = three['value']
                    temp_dict['Sender'] = msg_from
            else:
                pass

            temp_dict['Snippet'] = message['snippet'] # fetching message snippet

            try:
        
                # Fetching message body
                mssg_parts = payld['parts'] # fetching the message parts
                part_one  = mssg_parts[0] # fetching first element of the part 
                part_body = part_one['body'] # fetching body of the message
                part_data = part_body['data'] # fetching data from the body
                clean_one = part_data.replace("-","+") # decoding from Base64 to UTF-8
                clean_one = clean_one.replace("_","/") # decoding from Base64 to UTF-8
                clean_two = base64.b64decode (bytes(clean_one, 'UTF-8')) # decoding from Base64 to UTF-8
                soup = BeautifulSoup(clean_two , "lxml" )
                mssg_body = soup.body()
                # mssg_body is a readible form of message body
                # depending on the end user's requirements, it can be further cleaned 
                # using regex, beautiful soup, or any other method
                temp_dict['Message_body'] = mssg_body
            except:
                pass
            print (temp_dict)
            final_list.append(temp_dict) # This will create a dictonary item in the final list

        print ("Total messaged retrived: ", str(len(final_list)))

        '''
        '''
        for mssg in mssg_list:
            m_id = mssg['id']
            message = service.users().messages().get(userId='me', id=m_id, format='raw').execute()
            msg_str = base64.urlsafe_b64decode(message['raw'].encode('ASCII'))
            mime_msg = email.message_from_bytes(msg_str)
            print(mime_msg)

        '''

    except HttpError as error:
        # TODO(developer) - Handle errors from gmail API.
        print(f'An error occurred: {error}')


if __name__ == '__main__':
    main()