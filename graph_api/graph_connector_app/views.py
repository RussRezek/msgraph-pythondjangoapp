# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

from datetime import datetime, timedelta

from dateutil import parser, tz
from django.http import HttpResponseRedirect, JsonResponse
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

"""directory_list = ['root:/IT Solutions/2023-2024 Data/Data Reports - Broward:/children'
               ,'root:/IT Solutions/2023-2024 Data/Data Reports - Charlotte-Mecklenburg:/children'
               ,'root:/IT Solutions/2023-2024 Data/Data Reports - Delaware:/children'
               ,'root:/IT Solutions/2023-2024 Data/Data Reports - New York:/children'
               ,'root:/IT Solutions/2023-2024 Data/Data Reports - Palm Beach:/children'
               ,'root:/IT Solutions/2023-2024 Data/Data Reports - Philadelphia:/children' ]
               ,'root:/IT Solutions/2023-2024 Data/Data Reports - Dallas:/children'
               ,root:/IT Solutions/2023-2024 Data/Data Reports - Fort Worth:/children'
               ,root:/IT Solutions/2023-2024 Data/Data Reports - New York:/children'
               
"""            
directory_list = ['root:/IT Solutions/2024-2025 Data/New York:/children']

#runs on https://localhost:8000



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

def get_picker(request):
    context = initialize_context(request)


    #SET THE DRIVE AND DIRECTORY FOLDERS THAT WE WILL NEED

    #token = get_token(request)
    token = get_token_for_app(request)

    directory_list = ['root:/IT Solutions:/children']


    for directory in directory_list:
        years = get_filelist(token,drive,directory)

        if years:
            for index, year in enumerate(years['value']):
                created_date = parser.parse(year['createdDateTime'])
                year['createdDateTime'] = created_date.strftime('%Y-%m-%d %H:%M:%S')
                year['ParentDirectory'] = year['parentReference']['path'].rsplit('/', 1)[-1]
                year['AcademicYear'] = year['name']
                #RWR 2025-01-28 REMOVE LIOTA FOLDER:
                if year['name'] == 'LIOTA':
                    years['value'].pop(index)
            if 'ai_years' in context:
                for new_year in years['value']:
                    context['ai_years'].append(new_year)
            else:
                context['ai_years'] = years['value']

    return render(request, 'graph_connector_app/ai_folderpicker.html', context)

def get_districts(request):
    #context = initialize_context(request)

    ai_year = request.GET.get('ai_year_selection')

    directory = 'root:/IT Solutions/' + ai_year + ':/children'

    token = get_token_for_app(request)

    file_data = get_filelist(token,drive,directory)

    district_list = []
    
    for index, district in enumerate(file_data['value']):
        district_list.append(district['name'])

    response_data = {
        "districts" : district_list
    }    
  
    return JsonResponse(response_data)



def ai_files(request):
    context = initialize_context(request)
    #user = context['user']
    #if not user['is_authenticated']:
    #    return HttpResponseRedirect(reverse('signin'))

    #SET THE DRIVE AND DIRECTORY FOLDERS THAT WE WILL NEED

    #token = get_token(request)
    token = get_token_for_app(request)

    directory_list = []
    directory_list.append('root:/IT Solutions/' + request.POST.get('year') + '/' + request.POST.get('district') + ':/children')

    for directory in directory_list:
        files = get_filelist(token,drive,directory)

        if files:
            for file in files['value']:
                created_date = parser.parse(file['createdDateTime'])
                modified_date = parser.parse(file['lastModifiedDateTime'])
                file['createdDateTime'] = created_date.strftime('%Y-%m-%d %H:%M')
                file['lastModifiedDateTime'] = modified_date.strftime('%Y-%m-%d %H:%M')
                file['ParentDirectory'] = file['parentReference']['path'].rsplit('/', 1)[-1]
                file['AcademicYear'] = file['parentReference']['path'].rsplit('/', 2)[-2][:9]
                worksheet_data = get_worksheets(token,drive,file['id'])
                file['WorksheetName'] = worksheet_data['WorksheetName']
            if 'ai_files' in context:
                for new_file in files['value']:
                    context['ai_files'].append(new_file)
            else:
                context['ai_files'] = files['value']

    context['ai_directory_path'] = directory_list[0]

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
    directory_list = []
    file_info_list = []
    file_dict = {}

    directory_list.append(request.POST.get('directory_path'))

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
                file['AcademicYear'] = file['parentReference']['path'].rsplit('/', 2)[-2][:9]
                file_dict['AcademicYear'] = file['parentReference']['path'].rsplit('/', 2)[-2][:9]                
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
        if isinstance(record[24], str): #StartDate
            record[24] = None        
        if isinstance(record[25], str): #CompletionDate
            record[25] = None        
        if isinstance(record[29], str): #Duration(min)
            record[29] = None        
        if isinstance(record[32], str): #OverallScaleScore
            record[32] = None        
        if isinstance(record[35], str): #Percentile
            record[35] = None        
        if isinstance(record[36], str): #Grouping
            record[36] = None        
        if isinstance(record[39], str): #NumberandOperationsScaleScore
            record[39] = None        
        if isinstance(record[42], str): #AlgebraandAlgebraicThinkingScaleScore
            record[42] = None        
        if isinstance(record[45], str): #MeasurementandDataScaleScore
            record[45] = None        
        if isinstance(record[48], str): #GeometryScaleScore
            record[48] = None
        if isinstance(record[51], str): #DiagnosticGain
            record[51] = None
        if isinstance(record[52], str): #AnnualTypicalGrowthMeasure
            record[52] = None
        if isinstance(record[53], str): #AnnualStretchGrowthMeasure
            record[53] = None
        if isinstance(record[54], str): #[PercentProgresstoAnnualTypicalGrowth(%)]
            record[54] = None
        if isinstance(record[55], str): #[PercentProgresstoAnnualStretchGrowth(%)]
            record[55] = None
        if isinstance(record[56], str): #[MidOnGradeLevelScaleScore]
            record[56] = None
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

    #lp = sm.LoadProduction()
    #lp.load_production_tables()

    return render(request, 'graph_connector_app/file_data.html', context)

def get_all_reading_iready(request):
    context = initialize_context(request)

    token = get_token_for_app(request)
    directory_list = []
    file_info_list = []
    file_dict = {}

    directory_list.append(request.POST.get('directory_path'))

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
                file['AcademicYear'] = file['parentReference']['path'].rsplit('/', 2)[-2][:9]
                file_dict['AcademicYear'] = file['parentReference']['path'].rsplit('/', 2)[-2][:9]
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
        if isinstance(record[29], str): #Duration(min)
            record[29] = None
        if isinstance(record[32], str): #OverallScaleScore
            record[32] = None    
        if isinstance(record[35], str): #Percentile
            record[35] = None    
        if isinstance(record[36], str): #Grouping
            record[36] = None    
        if isinstance(record[39], str): #PhonologicalAwarenessScaleScore
            record[39] = None    
        if isinstance(record[42], str): #PhonicsScaleScore
            record[42] = None    
        if isinstance(record[45], str): #High-FrequencyWordsScaleScore
            record[45] = None    
        if isinstance(record[48], str): #VocabularyScaleScore
            record[48] = None    
        if isinstance(record[51], str): #ReadingComprehension:OverallScaleScore
            record[51] = None    
        if isinstance(record[54], str): #ReadingComprehension:LiteratureScaleScore
            record[54] = None    
        if isinstance(record[57], str): #ReadingComprehension:InformationalTextScaleScore
            record[57] = None    
        if isinstance(record[60], str): #DiagnosticGain
            record[60] = None
        if isinstance(record[63], str): #[PercentProgresstoAnnualTypicalGrowth(%)]
            record[63] = None
        if isinstance(record[64], str): #[PercentProgresstoAnnualStretchGrowth(%)]
            record[64] = None
        if isinstance(record[65], str): #MidOnGradeLevelScaleScore
            record[65] = None    

    for idx, record in reversed (list (enumerate (context['file_data']))):
        if record[4] == 'Student ID':
            context['file_data'].pop(idx)            

    db = sm.db

    trans = db.connection.begin()
    db.connection.execute("TRUNCATE TABLE AI.ReadingIreadyStaging")
    trans.commit()


    #for row in context['file_data']:
    for row in context['file_data']:
        db.connection.execute(sm.ReadingiReady.__table__.insert().values(row))

    #lp = sm.LoadProduction()
    #lp.load_production_tables()

    return render(request, 'graph_connector_app/file_data.html', context)

def get_all_eligibility(request):
    context = initialize_context(request)

    token = get_token_for_app(request)
    directory_list = []
    file_info_list = []
    file_dict = {}

    directory_list.append(request.POST.get('directory_path'))

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
                file['AcademicYear'] = file['parentReference']['path'].rsplit('/', 2)[-2][:9]
                file_dict['AcademicYear'] = file['parentReference']['path'].rsplit('/', 2)[-2][:9]
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
                row.insert(0, file['AcademicYear'])
                row.insert(0, 'Eligibility')
                row.insert(0, file['ParentDirectory'])
            context['file_data'] = file_data_tab
        else:
            file_data_tab = get_file_data(token,drive,file['id'],file['WorksheetName'])['values']
            for row in file_data_tab:
                row.insert(0, file['AcademicYear'])
                row.insert(0, 'Eligibility')
                row.insert(0, file['ParentDirectory'])
            context['file_data'] += file_data_tab
    for idx, record in enumerate (context['file_data']):
        if not isinstance(record[7], str): #StudentGrade
            record[7] = str(record[7])
    for idx, record in reversed (list (enumerate (context['file_data']))):
        if record[3] == 'School Name':
            context['file_data'].pop(idx)
    for idx, record in reversed (list (enumerate (context['file_data']))):    
        if record[6] == '':
            context['file_data'].pop(idx)


    db = sm.db

    trans = db.connection.begin()
    db.connection.execute("TRUNCATE TABLE AI.EligibilityStaging")
    trans.commit() 

    #for row in context['file_data']:
    for row in context['file_data']:
        db.connection.execute(sm.Eligibility.__table__.insert().values(row))

 
    return render(request, 'graph_connector_app/file_data.html', context)

def load_tables(request):

    context = initialize_context(request)
    db = sm.db

    lp = sm.LoadProduction()
    lp.load_production_tables()

    return render(request, 'graph_connector_app/home.html',context)

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
