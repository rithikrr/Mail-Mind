# Import the required libraries
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle
import os.path
import base64
import email
from scrapy.http import TextResponse
from scrapy.selector import Selector
import pandas as pd

# Define the SCOPES. If modifying it, delete the token.pickle file.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def getEmails():
    # ... Same code for getting credentials and connecting to the Gmail API ...
    creds = None
  
    # The file ztoken.pickle contains the user access token.
    # If credentials are not available or are invalid, ask the user to log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(r'C:\Users\thari\OneDrive\Desktop\InboxIQ\Main\credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
  
        # Save the access token in token.pickle file for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    # Connect to the Gmail API
    service = build('gmail', 'v1', credentials=creds)

    # Request a list of all the messages
    result = service.users().messages().list(userId='me').execute()

    # messages is a list of dictionaries where each dictionary contains a message id.
    messages = result.get('messages')

    # iterate through all the messages
    email_details_list = []
    for msg in messages:
        # Get the message from its id
        txt = service.users().messages().get(userId='me', id=msg['id']).execute()

        # Use try-except to avoid any Errors
        try:
            # Get value of 'payload' from dictionary 'txt'
            payload = txt['payload']
            headers = payload['headers']

            # Look for Subject and Sender Email in the headers
            for d in headers:
                if d['name'] == 'Subject':
                    subject = d['value']
                if d['name'] == 'From':
                    sender = d['value']

            # The Body of the message is in Encrypted format. So, we have to decode it.
            # Get the data and decode it with base 64 decoder.
            parts = payload.get('parts')[0]
            data = parts['body']['data']
            data = data.replace("-", "+").replace("_", "/")
            decoded_data = base64.b64decode(data)

            # Parse the decoded data with Scrapy's TextResponse
            response = TextResponse(url='dummy_url', body=decoded_data)
            body = response.xpath('//body//text()').get()

            email_details_list.append({
                'Subject': subject,
                'From': sender,
                'Message': body.strip(),  # Strip any leading/trailing spaces or newlines
            })


        except:
            pass
    df = pd.DataFrame(email_details_list)

    # Save the DataFrame to an Excel file
    df.to_excel('email_details.xlsx', index=False)

getEmails()
