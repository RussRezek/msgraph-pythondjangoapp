# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

from datetime import datetime, timedelta

from dateutil import parser, tz
from django.http import HttpResponseRedirect
from django.shortcuts import render
from django.urls import reverse
from graph_connector_app.auth_helper import (get_sign_in_flow, get_token,
                                  get_token_from_code, get_token_for_app,
                                  remove_user_and_token, store_user)
from graph_connector_app.graph_helper import (create_event, get_calendar_events,
                                   get_file_data, get_filelist,
                                   get_iana_from_windows, get_user,
                                   get_worksheets)
from graph_connector_app.sqlalchemy_models import sql_models as sm

#SET DRIVE AND DIRECTORY LIST

drive = '/drives/b!j7GCe-Two0-73phJMu9XqviV7CunZMpEm0xZVyOP271WIaR6JXsIQKeBR7P0r-37/'
directory_list = ['root:/IT Solutions/2023-2024 Data/Data Reports - Broward:/children'
               ,'root:/IT Solutions/2023-2024 Data/Data Reports - Charlotte-Mecklenburg:/children'
               ,'root:/IT Solutions/2023-2024 Data/Data Reports - Delaware:/children'
               ,'root:/IT Solutions/2023-2024 Data/Data Reports - New York:/children'
               ,'root:/IT Solutions/2023-2024 Data/Data Reports - Palm Beach:/children'
               ,'root:/IT Solutions/2023-2024 Data/Data Reports - Philadelphia:/children' ]
               
#directory_list = ['root:/IT Solutions/Data Reports - Philadelphia:/children']



def initialize_context(request):
    context = {}

    # Check for any errors in the session
    error = request.session.pop('flash_error', None)

    if error is not None:
        context['errors'] = []
        context['errors'].append(error)

    # Check for user in the session
    context['user'] = request.session.get('user', {'is_authenticated': False})
    return context

def home(request):
    context = initialize_context(request)

    return render(request, 'graph_connector_app/home.html', context)

def sign_in(request):
    # Get the sign-in flow
    flow = get_sign_in_flow()
    # Save the expected flow so we can use it in the callback
    request.session['auth_flow'] = flow

    # Redirect to the Azure sign-in page
    return HttpResponseRedirect(flow['auth_uri'])

def callback(request):
    # Make the token request
    result = get_token_from_code(request)

    #Get the user's profile
    user = get_user(result['access_token'])

    # Store user
    store_user(request, user)
    return HttpResponseRedirect(reverse('home'))

def sign_out(request):
    # Clear out the user and token
    remove_user_and_token(request)

    return HttpResponseRedirect(reverse('home'))

def calendar(request):
    context = initialize_context(request)
    user = context['user']
    if not user['is_authenticated']:
        return HttpResponseRedirect(reverse('signin'))

    # Load the user's time zone
    # Microsoft Graph can return the user's time zone as either
    # a Windows time zone name or an IANA time zone identifier
    # Python datetime requires IANA, so convert Windows to IANA
    time_zone = get_iana_from_windows(user['timeZone'])
    tz_info = tz.gettz(time_zone)

    # Get midnight today in user's time zone
    today = datetime.now(tz_info).replace(
        hour=0,
        minute=0,
        second=0,
        microsecond=0)

    # Based on today, get the start of the week (Sunday)
    if today.weekday() != 6:
        start = today - timedelta(days=today.isoweekday())
    else:
        start = today

    end = start + timedelta(days=7)

    token = get_token(request)

    events = get_calendar_events(
        token,
        start.isoformat(timespec='seconds'),
        end.isoformat(timespec='seconds'),
        user['timeZone'])

    if events:
        # Convert the ISO 8601 date times to a datetime object
        # This allows the Django template to format the value nicely
        for event in events['value']:
            event['start']['dateTime'] = parser.parse(event['start']['dateTime'])
            event['end']['dateTime'] = parser.parse(event['end']['dateTime'])

        context['events'] = events['value']

    return render(request, 'graph_connector_app/calendar.html', context)

def ai_files(request):
    context = initialize_context(request)
    #user = context['user']
    #if not user['is_authenticated']:
    #    return HttpResponseRedirect(reverse('signin'))

    #SET THE DRIVE AND DIRECTORY FOLDERS THAT WE WILL NEED

    #token = get_token(request)
    token = get_token_for_app(request)

    for directory in directory_list:
        files = get_filelist(token,drive,directory)

        if files:
            for file in files['value']:
                created_date = parser.parse(file['createdDateTime'])
                file['createdDateTime'] = created_date.strftime('%Y-%m-%d %H:%M:%S')
                file['ParentDirectory'] = file['parentReference']['path'].rsplit('/', 1)[-1]
                worksheet_data = get_worksheets(token,drive,file['id'])
                file['WorksheetName'] = worksheet_data['WorksheetName']
            if 'ai_files' in context:
                for new_file in files['value']:
                    context['ai_files'].append(new_file)
            else:
                context['ai_files'] = files['value']


    return render(request, 'graph_connector_app/ai_files.html', context)

def file_data(request, file_id, worksheet_name):
    context = initialize_context(request)

    token = get_token_for_app(request)

    context['file_data'] = get_file_data(token,drive,file_id,worksheet_name)['values']
    context['file_name'] = request.GET['file_name']

    return render(request, 'graph_connector_app/file_data.html', context)

def get_all_math_iready(request):
    context = initialize_context(request)

    token = get_token_for_app(request)
    file_info_list = []
    file_dict = {}

    for directory in directory_list:
        files = get_filelist(token,drive,directory)

        if files:
            for file in files['value']:
                file_dict['FileName'] = file['name']
                file_dict['id'] = file['id']
                created_date = parser.parse(file['createdDateTime'])
                file['createdDateTime'] = created_date.strftime('%Y-%m-%d %H:%M:%S')
                file_dict['createdDateTime'] = created_date.strftime('%Y-%m-%d %H:%M:%S')
                file['ParentDirectory'] = file['parentReference']['path'].rsplit('/', 1)[-1]
                file_dict['ParentDirectory'] = file['parentReference']['path'].rsplit('/', 1)[-1]
                worksheet_data = get_worksheets(token,drive,file['id'])
                file['WorksheetName'] = worksheet_data['WorksheetName']
                file_dict['WorksheetName'] = worksheet_data['WorksheetName']

                if '_math' in file_dict['FileName'].lower():
                    file_info_list.append(file_dict.copy())

    file_data_tab =[]
    for file in file_info_list:
        if 'file_data' not in context:
            file_data_tab = get_file_data(token,drive,file['id'],file['WorksheetName'])['values']
            for row in file_data_tab:
                row.insert(0, 'Math iReady')
                row.insert(0, file['ParentDirectory'])
            context['file_data'] = file_data_tab
        else:
            file_data_tab = get_file_data(token,drive,file['id'],file['WorksheetName'])['values']
            for row in file_data_tab:
                row.insert(0, 'Math iReady')
                row.insert(0, file['ParentDirectory'])
            context['file_data'] += file_data_tab
    for idx, record in enumerate (context['file_data']):
        if not isinstance(record[5], str): #StudentGrade
            record[5] = str(record[5])
        if isinstance(record[45], str): #DiagnosticGain
            record[45] = None
        if isinstance(record[48], str): #[PercentProgresstoAnnualTypicalGrowth(%)]
            record[48] = None
        if isinstance(record[49], str): #[PercentProgresstoAnnualStretchGrowth(%)]
            record[49] = None
    for idx, record in enumerate (context['file_data']):
        if record[4] == 'Student ID':
            context['file_data'].pop(idx)

    db = sm.db

    trans = db.connection.begin()
    db.connection.execute("TRUNCATE TABLE AI.MathIreadyStaging")
    trans.commit()

    #for row in context['file_data']:
    for row in context['file_data']:
        db.connection.execute(sm.MathiReady.__table__.insert().values(row))

    lp = sm.LoadProduction()
    lp.load_production_tables()

    return render(request, 'graph_connector_app/file_data.html', context)

def get_all_reading_iready(request):
    context = initialize_context(request)

    token = get_token_for_app(request)
    file_info_list = []
    file_dict = {}

    for directory in directory_list:
        files = get_filelist(token,drive,directory)

        if files:
            for file in files['value']:
                file_dict['FileName'] = file['name']
                file_dict['id'] = file['id']
                created_date = parser.parse(file['createdDateTime'])
                file['createdDateTime'] = created_date.strftime('%Y-%m-%d %H:%M:%S')
                file_dict['createdDateTime'] = created_date.strftime('%Y-%m-%d %H:%M:%S')
                file['ParentDirectory'] = file['parentReference']['path'].rsplit('/', 1)[-1]
                file_dict['ParentDirectory'] = file['parentReference']['path'].rsplit('/', 1)[-1]
                worksheet_data = get_worksheets(token,drive,file['id'])
                file['WorksheetName'] = worksheet_data['WorksheetName']
                file_dict['WorksheetName'] = worksheet_data['WorksheetName']

                if '_ela' in file_dict['FileName'].lower():
                    file_info_list.append(file_dict.copy())

    file_data_tab =[]
    for file in file_info_list:
        if 'file_data' not in context:
            file_data_tab = get_file_data(token,drive,file['id'],file['WorksheetName'])['values']
            for row in file_data_tab:
                row.insert(0, 'ELA iReady')
                row.insert(0, file['ParentDirectory'])
            context['file_data'] = file_data_tab
        else:
            file_data_tab = get_file_data(token,drive,file['id'],file['WorksheetName'])['values']
            for row in file_data_tab:
                row.insert(0, 'ELA iReady')
                row.insert(0, file['ParentDirectory'])
            context['file_data'] += file_data_tab
    for idx, record in enumerate (context['file_data']):
        if not isinstance(record[5], str): #StudentGrade
            record[5] = str(record[5])
        if isinstance(record[54], str): #DiagnosticGain
            record[54] = None
        if isinstance(record[57], str): #[PercentProgresstoAnnualTypicalGrowth(%)]
            record[57] = None
        if isinstance(record[58], str): #[PercentProgresstoAnnualStretchGrowth(%)]
            record[58] = None
    for idx, record in enumerate (context['file_data']):
        if record[4] == 'Student ID':
            context['file_data'].pop(idx)

    db = sm.db

    trans = db.connection.begin()
    db.connection.execute("TRUNCATE TABLE AI.ReadingIreadyStaging")
    trans.commit()


    #for row in context['file_data']:
    for row in context['file_data']:
        db.connection.execute(sm.ReadingiReady.__table__.insert().values(row))

    lp = sm.LoadProduction()
    lp.load_production_tables()

    return render(request, 'graph_connector_app/file_data.html', context)

def get_all_eligibility(request):
    context = initialize_context(request)

    token = get_token_for_app(request)
    file_info_list = []
    file_dict = {}

    for directory in directory_list:
        files = get_filelist(token,drive,directory)

        if files:
            for file in files['value']:
                file_dict['FileName'] = file['name']
                file_dict['id'] = file['id']
                created_date = parser.parse(file['createdDateTime'])
                file['createdDateTime'] = created_date.strftime('%Y-%m-%d %H:%M:%S')
                file_dict['createdDateTime'] = created_date.strftime('%Y-%m-%d %H:%M:%S')
                file['ParentDirectory'] = file['parentReference']['path'].rsplit('/', 1)[-1]
                file_dict['ParentDirectory'] = file['parentReference']['path'].rsplit('/', 1)[-1]
                worksheet_data = get_worksheets(token,drive,file['id'])
                file['WorksheetName'] = worksheet_data['WorksheetName']
                file_dict['WorksheetName'] = worksheet_data['WorksheetName']

                if 'eligibility' in file_dict['FileName'].lower():
                    file_info_list.append(file_dict.copy())

    file_data_tab =[]
    for file in file_info_list:
        if 'file_data' not in context:
            file_data_tab = get_file_data(token,drive,file['id'],file['WorksheetName'])['values']
            for row in file_data_tab:
                row.insert(0, 'Eligibilty')
                row.insert(0, file['ParentDirectory'])
            context['file_data'] = file_data_tab
        else:
            file_data_tab = get_file_data(token,drive,file['id'],file['WorksheetName'])['values']
            for row in file_data_tab:
                row.insert(0, 'Eligibility')
                row.insert(0, file['ParentDirectory'])
            context['file_data'] += file_data_tab
    for idx, record in enumerate (context['file_data']):
        if not isinstance(record[6], str): #StudentGrade
            record[6] = str(record[6])
    for idx, record in enumerate (context['file_data']):
        if record[2] == 'ReferralStatus':
            context['file_data'].pop(idx)

    db = sm.db

    trans = db.connection.begin()
    db.connection.execute("TRUNCATE TABLE AI.EligibilityStaging")
    trans.commit() 

    #for row in context['file_data']:
    for row in context['file_data']:
        db.connection.execute(sm.Eligibility.__table__.insert().values(row))

    lp = sm.LoadProduction()
    lp.load_production_tables()


    return render(request, 'graph_connector_app/file_data.html', context)

def new_event(request):
    context = initialize_context(request)
    user = context['user']
    if not user['is_authenticated']:
        return HttpResponseRedirect(reverse('signin'))

    if request.method == 'POST':
        # Validate the form values
        # Required values
        if (not request.POST['ev-subject']) or \
            (not request.POST['ev-start']) or \
            (not request.POST['ev-end']):
            context['errors'] = [
                {
                    'message': 'Invalid values',
                    'debug': 'The subject, start, and end fields are required.'
                }
            ]
            return render(request, 'graph_connector_app/newevent.html', context)

        attendees = None
        if request.POST['ev-attendees']:
            attendees = request.POST['ev-attendees'].split(';')

        # Create the event
        token = get_token(request)

        create_event(
          token,
          request.POST['ev-subject'],
          request.POST['ev-start'],
          request.POST['ev-end'],
          attendees,
          request.POST['ev-body'],
          user['timeZone'])

        # Redirect back to calendar view
        return HttpResponseRedirect(reverse('calendar'))
    else:
        # Render the form
        return render(request, 'graph_connector_app/newevent.html', context)
