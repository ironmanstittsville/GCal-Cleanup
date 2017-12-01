# =========================================================================
# Use this to cleanup your Google Calendar.
#   I developed this little script based heavily on the Google Calendar
#   API quickstart example python code.
# I used this because my Google Calendar had become frightfully polluted
#   with events that had been sync-ed into it over years by other apps
#   but they became independent events that ran across the next 20 years!
#   Crazy.
# This script will use the "LIST_OF_RECURRING_EVENTS" and will search the
#   next 1000 future events in your Google Calendar and delete any
#   matching ones for you.
#
# I ran this on an AMD PC running Windows 10.
# I ran this within the PyCharm IDE for Python (Community Edition)
# I also installed the Google Calendar API before I began. I found
#   instructions here: https://developers.google.com/google-apps/calendar/quickstart/python
# The Python version that I'm running is: Python 3.5.2
# =========================================================================

# -------------- Imports --------------
import httplib2
import os
import pprint
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
import datetime

# -------------- Structures --------------
# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/GCal-Cleanup.json
SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'GCal-Cleanup.json'
APPLICATION_NAME = 'GCal-Cleanup'

# This list contains the names of the events that you want to find and
#   delete from your Google Calendar.
LIST_OF_RECURRING_EVENTS = (
    "Oil change for the Malibu",
    "Alcatel Payday",
    "Global Knowledge Payday",
    "Poker w/ Cinnabar guys.",
    "Squish w/ Jack",
    "Beavers Colony B.",
    "Cubs: Coyote Pack @ Guradian Angels",
    "CSI is on! Ch 7",
    "Clean the furnace filter!",
    " Walk for Hospice at May Court."
)

# -------------- The Appln --------------
def get_credentials():
    """Gets valid user credentials from storage.
    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.
    Returns:
        Credentials, the obtained credential.
    """
    cur_dir = os.path.curdir
    credential_dir = os.path.join(cur_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'GCal-Cleanup.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def print_list(aList):
    pp = pprint.PrettyPrinter(indent=4)
    print("Looking for the following matching events:")
    pp.pprint(aList)
    print("")

def print_events(aList):
    pp = pprint.PrettyPrinter(indent=4)
    for item in aList:
        pp.pprint("  {}".format(item))
    print("")

def main():
    """ The above list of rascally recurring events that are separate events due to
        a sync issue.
        This script will modify the user calendar to remove many of the upcoming
        instances.
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    now = datetime.datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
    print_list(LIST_OF_RECURRING_EVENTS)
    eventsResult = service.events().list(
        calendarId='primary', timeMin=now, maxResults=6, singleEvents=True,
        orderBy='startTime', showDeleted=False).execute()
    events = eventsResult.get('items', [])

    if not events:
        print('No upcoming events found.')

    list_of_deleted_events = []
    counter = 0
    for event in events:
        if event['summary'] in LIST_OF_RECURRING_EVENTS:
            start = event['start'].get('dateTime', event['start'].get('date'))
            list_of_deleted_events.append("DELETED: {} - {}".format(start,event['summary']))
            service.events().delete(calendarId='primary',eventId=event['id']).execute()
            counter += 1

    print("Deleting these events:")
    if counter == 0:
        print("  No matching events found.")
    else:
        print_events(list_of_deleted_events)

if __name__ == '__main__':
    main()