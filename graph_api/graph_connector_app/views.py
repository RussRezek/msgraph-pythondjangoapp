# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

from datetime import datetime, timedelta
from django.contrib import messages
from dateutil import parser, tz
from django.http import HttpResponseRedirect, JsonResponse
from django.shortcuts import render
from django.urls import reverse
from graph_main_app import settings
from graph_connector_app.auth_helper import (get_sign_in_flow, get_token,
                                  get_token_from_code, get_token_for_app,
                                  remove_user_and_token, store_user)
from graph_connector_app.graph_helper import (create_event, get_calendar_events,
                                   get_file_data, get_filelist,
                                   get_iana_from_windows, get_user,
                                   get_worksheets, get_sharepoint_users_via_graph)
from graph_connector_app.sqlalchemy_models import sql_models as sm
from datetime import date, timedelta
import logging
import traceback

logger = logging.getLogger(__name__)

#SET DRIVE AND DIRECTORY LIST FOR AI FILES

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
help_desk_email = 'Russell.Rezek@learnbehavioral.com'

#runs on https://localhost:8000

def excel_date_to_python(excel_date):
    """Convert Excel serial date to Python date object"""
    if isinstance(excel_date, (int, float)) and excel_date > 0:
        excel_epoch = date(1899, 12, 30)
        return excel_epoch + timedelta(days=int(excel_date))
    return None


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
    """
    Retrieve and aggregate all Math iReady assessment files for a selected directory.

    This view is typically invoked by an AJAX POST request from the UI when the user
    selects a specific academic year and district directory that contains Math iReady
    files. It enumerates all files in the provided directory path using Microsoft Graph,
    gathers worksheet information, and prepares normalized metadata needed to render
    a combined Math iReady view.

    Args:
        request (HttpRequest): The incoming Django request. It is expected to contain
            a ``directory_path`` key in ``request.POST`` that identifies the
            OneDrive/SharePoint folder to scan for Math iReady files.

    Returns:
        HttpResponse: A rendered HTML response that uses the populated context
        (including Math iReady file metadata and worksheet information) to display
        aggregated Math iReady data to the user.

    Raises:
        ValueError: If a data row length does not match the number of columns
        Exception: Re-raised database insertion errors after logging

    Process:
        1. Initialize the base template context via ``initialize_context`` and set
           the ``file_source`` label.
        2. Acquire an application token via ``get_token_for_app``.
        3. Build a list of directory paths starting from the submitted
           ``directory_path``.
        4. For each directory, call ``get_filelist`` with the current drive and
           directory to retrieve file metadata from Microsoft Graph.
        5. For each file, normalize date fields, infer the parent directory and
           academic year, and call ``get_worksheets`` to fetch worksheet details.
        6. Accumulate file and worksheet information into intermediate collections
           (such as ``file_info_list`` and ``file_dict``) for downstream processing.
        7. Update the template context with the aggregated Math iReady data and
           render the appropriate results page.
    """

    context = initialize_context(request)
    file_source = 'Math iReady'

    try: 
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
                for idx, row in enumerate (file_data_tab):
                    if idx == 0:
                        row.insert(0, 'Subject')
                        row.insert(0, 'District')
                    else:
                        row.insert(0, file_source)
                        row.insert(0, file['ParentDirectory'])
                context['file_data'] = file_data_tab
            else:
                file_data_tab = get_file_data(token,drive,file['id'],file['WorksheetName'])['values']
                for idx, row in enumerate (file_data_tab):
                    if idx == 0:
                        row.insert(0, 'Subject')
                        row.insert(0, 'District')
                    else:
                        row.insert(0, file_source)
                        row.insert(0, file['ParentDirectory'])
                context['file_data'] += file_data_tab

        # First pass: Remove header rows
        excel_headers = None
        for idx, record in reversed(list(enumerate(context['file_data']))):
            if record[4] == 'Student ID':
                # Get Column Headers from Excel File
                excel_headers = [item.replace(' ', '') if isinstance(item, str) else item for item in record]
                context['file_data'].pop(idx)
                continue

        # Check if headers were found
        if excel_headers is None:
            error_msg = f"{file_source} Error: Header row not found in Excel file. Expected to find 'Student ID' in column 5 of the header row. Contact {help_desk_email}."
            logger.error(error_msg)
            messages.error(request, error_msg)
            return render(request, 'graph_connector_app/file_data.html', context)
        
        # Convert data to a dictionary using Excel Headers as Keys
        columns = excel_headers

         # Validate Excel headers match database columns
        db_columns = [c.name for c in sm.MathiReady.__table__.columns]
        extra_columns = [col for col in excel_headers if col not in db_columns]
        
        if extra_columns:
            error_msg = f"{file_source} column mismatch: "
            if extra_columns:
                error_msg += f"Incorrect column name in Excel: {', '.join(extra_columns)}. "
            error_msg += f"Contact {help_desk_email}."
            logger.error(error_msg)
            messages.error(request, error_msg)
            return render(request, 'graph_connector_app/file_data.html', context)
    
        #columns = [c.name for c in sm.MathiReady.__table__.columns]
        records = []
        for row in context['file_data']:
            if len(row) != len(columns):
                raise ValueError(f"{file_source}: Row length {len(row)} does not match columns length {len(columns)}")
            records.append(dict(zip(columns, row)))



        # Second pass: Convert data types
        for record in records:
            if not isinstance(record['StudentGrade'], str): #StudentGrade
                record['StudentGrade'] = str(record['StudentGrade'])
            if isinstance(record['StartDate'], str): #StartDate
                record['StartDate'] = None  
            else:
                record['StartDate'] = excel_date_to_python(record['StartDate']) 
            if isinstance(record['CompletionDate'], str): #CompletionDate
                record['CompletionDate'] = None  
            else:
                record['CompletionDate'] = excel_date_to_python(record['CompletionDate'])      
            if isinstance(record['Duration(min)'], str): #Duration(min)
                record['Duration(min)'] = None        
            if isinstance(record['OverallScaleScore'], str): #OverallScaleScore
                record['OverallScaleScore'] = None        
            if isinstance(record['Percentile'], str): #Percentile
                record['Percentile'] = None        
            if isinstance(record['Grouping'], str): #Grouping
                record['Grouping'] = None        
            if isinstance(record['NumberandOperationsScaleScore'], str): #NumberandOperationsScaleScore
                record['NumberandOperationsScaleScore'] = None        
            if isinstance(record['AlgebraandAlgebraicThinkingScaleScore'], str): #AlgebraandAlgebraicThinkingScaleScore
                record['AlgebraandAlgebraicThinkingScaleScore'] = None        
            if isinstance(record['MeasurementandDataScaleScore'], str): #MeasurementandDataScaleScore
                record['MeasurementandDataScaleScore'] = None        
            if isinstance(record['GeometryScaleScore'], str): #GeometryScaleScore
                record['GeometryScaleScore'] = None
            if isinstance(record['DiagnosticGain'], str): #DiagnosticGain
                record['DiagnosticGain'] = None
            if isinstance(record['AnnualTypicalGrowthMeasure'], str): #AnnualTypicalGrowthMeasure
                record['AnnualTypicalGrowthMeasure'] = None
            if isinstance(record['AnnualStretchGrowthMeasure'], str): #AnnualStretchGrowthMeasure
                record['AnnualStretchGrowthMeasure'] = None
            if isinstance(record['PercentProgresstoAnnualTypicalGrowth(%)'], str): #[PercentProgresstoAnnualTypicalGrowth(%)]
                record['PercentProgresstoAnnualTypicalGrowth(%)'] = None
            if isinstance(record['PercentProgresstoAnnualStretchGrowth(%)'], str): #[PercentProgresstoAnnualStretchGrowth(%)]
                record['PercentProgresstoAnnualStretchGrowth(%)'] = None
            if isinstance(record['MidOnGradeLevelScaleScore'], str): #[MidOnGradeLevelScaleScore]
                record['MidOnGradeLevelScaleScore'] = None
        
        db = sm.db

        try:
            trans = db.connection.begin()
            db.connection.execute("TRUNCATE TABLE AI.MathIreadyStaging")
            trans.commit()
        except Exception as e:
            logger.error(f"Error truncating {file_source} staging table: {type(e).__name__}: {str(e)}")
            raise e

        if records:
            try:
                trans = db.connection.begin()
                db.connection.execute(sm.MathiReady.__table__.insert(), records)
                trans.commit()
                messages.success(request, f'{file_source} Data Successfully Imported!')
            except Exception as e:
                logger.error(f"Error inserting {file_source} records: {type(e).__name__}: {str(e)}")
                messages.error(request, f'Something went wrong importing {file_source} data: {str(e)}. Contact {help_desk_email}.')
                raise

        lp = sm.LoadProduction()
        lp.load_production_tables(2)

    except Exception as e:
        logger.error(f"Error loading production tables: {type(e).__name__}: {str(e)}")
        messages.error(request, f'Something went wrong loading production tables: {str(e)}. Contact {help_desk_email}.')

    return render(request, 'graph_connector_app/file_data.html', context)

def get_all_reading_iready(request):
    """
    Retrieve and process iReady reading assessment data from Microsoft Graph API.
    
    This function fetches Reading iReady Excel files from a specified directory,
    extracts worksheet data, validates and transforms the data types, and loads
    the processed records into the ReadingIreadyStaging database table.
    
    Args:
        request: HttpRequest object containing POST data with 'directory_path'
        
    Returns:
        HttpResponse: Renders 'graph_connector_app/file_data.html' with context
                     containing processed file_data and other context information
                     
    Raises:
        ValueError: If a data row length does not match the number of columns
        Exception: Re-raised database insertion errors after logging
        
    Process:
        1. Retrieves authentication token and initializes context
        2. Fetches file list from specified directory via Graph API
        3. Filters for files containing '_ela' in filename
        4. Extracts and processes worksheet data (dates, file metadata)
        5. Removes header rows and normalizes column names
        6. Converts data types (dates, numeric fields)
        7. Truncates staging table and bulk inserts processed records
        8. Triggers production table load pipeline
    """
    context = initialize_context(request)
    file_source = 'Reading iReady'

    try:
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
                for idx, row in enumerate(file_data_tab):
                    if idx == 0:
                        row.insert(0, 'Subject')
                        row.insert(0, 'District')
                    else:
                        row.insert(0, file_source)
                        row.insert(0, file['ParentDirectory'])  
                context['file_data'] = file_data_tab
            else:
                file_data_tab = get_file_data(token,drive,file['id'],file['WorksheetName'])['values']
                for idx, row in enumerate(file_data_tab):
                    if idx == 0:
                        row.insert(0, 'Subject')
                        row.insert(0, 'District')
                    else:
                        row.insert(0, file_source)
                        row.insert(0, file['ParentDirectory'])  
                context['file_data'] += file_data_tab

        # First pass: Remove header rows
        excel_headers = None
        for idx, record in reversed(list(enumerate(context['file_data']))):
            if record[4] == 'Student ID':
                # Get Column Headers from Excel File
                excel_headers = [item.replace(' ', '') if isinstance(item, str) else item for item in record]
                context['file_data'].pop(idx)
                continue

        # Check if headers were found
        if excel_headers is None:
            error_msg = f"{file_source} Error: Header row not found in Excel file. Expected to find 'Student ID' in column 5 of the header row. Contact {help_desk_email}."
            logger.error(error_msg)
            messages.error(request, error_msg)
            return render(request, 'graph_connector_app/file_data.html', context)

        # Convert data to a dictionary using Excel Headers as Keys
        columns = excel_headers

        # Validate Excel headers match database columns
        db_columns = [c.name for c in sm.ReadingiReady.__table__.columns]
        extra_columns = [col for col in excel_headers if col not in db_columns]
        
        if extra_columns:
            error_msg = f"{file_source} column mismatch: "
            if extra_columns:
                error_msg += f"Incorrect column name in Excel: {', '.join(extra_columns)}. "
            error_msg += f"Contact {help_desk_email}."
            logger.error(error_msg)
            messages.error(request, error_msg)
            return render(request, 'graph_connector_app/file_data.html', context)

        #columns = [c.name for c in sm.ReadingiReady.__table__.columns]
        records = []
        for row in context['file_data']:
            if len(row) != len(columns):
                raise ValueError(f"{file_source}: Row length {len(row)} does not match columns length {len(columns)}")
            records.append(dict(zip(columns, row)))
        
        for record in records:
            if not isinstance(record['StudentGrade'], str): #StudentGrade
                record['StudentGrade'] = str(record['StudentGrade'])
            if isinstance(record['StartDate'], str): #StartDate
                record['StartDate'] = None  
            else:
                record['StartDate'] = excel_date_to_python(record['StartDate']) 
            if isinstance(record['CompletionDate'], str): #CompletionDate
                record['CompletionDate'] = None  
            else:
                record['CompletionDate'] = excel_date_to_python(record['CompletionDate'])      
            if isinstance(record['Duration(min)'], str) or record['Duration(min)'] == '': #Duration(min)
                record['Duration(min)'] = None
            if isinstance(record['OverallScaleScore'], str) or record['OverallScaleScore'] == '': #OverallScaleScore
                record['OverallScaleScore'] = None    
            if isinstance(record['Percentile'], str) or record['Percentile'] == '': #Percentile
                record['Percentile'] = None    
            if isinstance(record['Grouping'], str) or record['Grouping'] == '': #Grouping
                record['Grouping'] = None    
            if isinstance(record['PhonologicalAwarenessScaleScore'], str) or record['PhonologicalAwarenessScaleScore'] == '': #PhonologicalAwarenessScaleScore
                record['PhonologicalAwarenessScaleScore'] = None    
            if isinstance(record['PhonicsScaleScore'], str) or record['PhonicsScaleScore'] == '': #PhonicsScaleScore
                record['PhonicsScaleScore'] = None    
            if isinstance(record['High-FrequencyWordsScaleScore'], str) or record['High-FrequencyWordsScaleScore'] == '': #High-FrequencyWordsScaleScore
                record['High-FrequencyWordsScaleScore'] = None    
            if isinstance(record['VocabularyScaleScore'], str) or record['VocabularyScaleScore'] == '': #VocabularyScaleScore
                record['VocabularyScaleScore'] = None    
            if isinstance(record['Comprehension:OverallScaleScore'], str) or record['Comprehension:OverallScaleScore'] == '': #Comprehension:OverallScaleScore
                record['Comprehension:OverallScaleScore'] = None    
            if isinstance(record['Comprehension:LiteratureScaleScore'], str) or record['Comprehension:LiteratureScaleScore'] == '': #Comprehension:LiteratureScaleScore
                record['Comprehension:LiteratureScaleScore'] = None    
            if isinstance(record['Comprehension:InformationalTextScaleScore'], str) or record['Comprehension:InformationalTextScaleScore'] == '': #Comprehension:InformationalTextScaleScore
                record['Comprehension:InformationalTextScaleScore'] = None    
            if isinstance(record['DiagnosticGain'], str) or record['DiagnosticGain'] == '': #DiagnosticGain
                record['DiagnosticGain'] = None
            if isinstance(record['AnnualTypicalGrowthMeasure'], str) or record['AnnualTypicalGrowthMeasure'] == '': #AnnualTypicalGrowthMeasure
                record['AnnualTypicalGrowthMeasure'] = None
            if isinstance(record['AnnualStretchGrowthMeasure'], str) or record['AnnualStretchGrowthMeasure'] == '': #AnnualStretchGrowthMeasure
                record['AnnualStretchGrowthMeasure'] = None
            if isinstance(record['PercentProgresstoAnnualTypicalGrowth(%)'], str) or record['PercentProgresstoAnnualTypicalGrowth(%)'] == '': #[PercentProgresstoAnnualTypicalGrowth(%)]
                record['PercentProgresstoAnnualTypicalGrowth(%)'] = None
            if isinstance(record['PercentProgresstoAnnualStretchGrowth(%)'], str) or record['PercentProgresstoAnnualStretchGrowth(%)'] == '': #[PercentProgresstoAnnualStretchGrowth(%)]
                record['PercentProgresstoAnnualStretchGrowth(%)'] = None
            if isinstance(record['MidOnGradeLevelScaleScore'], str) or record['MidOnGradeLevelScaleScore'] == '': #MidOnGradeLevelScaleScore
                record['MidOnGradeLevelScaleScore'] = None

        db = sm.db

        try:
            trans = db.connection.begin()
            db.connection.execute("TRUNCATE TABLE AI.ReadingIreadyStaging")
            trans.commit()
        except Exception as e:    
            logger.error(f"Error truncating {file_source} table: {type(e).__name__}: {str(e)}")
            raise


        if records:
            try:
                trans = db.connection.begin()
                db.connection.execute(sm.ReadingiReady.__table__.insert(), records)
                trans.commit()
                messages.success(request, f'{file_source} Data Successfully Imported!')
            except Exception as e:
                logger.error(f"Error inserting {file_source} records: {type(e).__name__}: {str(e)}")
                messages.error(request, f'Something went wrong importing {file_source} data: {str(e)}. Contact {help_desk_email}.')
                raise

            lp = sm.LoadProduction()
            lp.load_production_tables(1)

    except Exception as e:
        logger.error(f"Error loading production tables: {type(e).__name__}: {str(e)}")
        messages.error(request, f'Something went wrong loading production tables: {str(e)}. Contact {help_desk_email}.')



    return render(request, 'graph_connector_app/file_data.html', context)

def get_all_eligibility(request):
    context = initialize_context(request)
    file_source = 'Eligibility'

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
            for idx, row in enumerate(file_data_tab):
                if idx == 0:
                    row.insert(0, 'AcademicYear')
                    row.insert(0, 'Subject')
                    row.insert(0, 'District')
                else:
                    row.insert(0, file['AcademicYear'])
                    row.insert(0, file_source)
                    row.insert(0, file['ParentDirectory'])
            context['file_data'] = file_data_tab
        else:
            file_data_tab = get_file_data(token,drive,file['id'],file['WorksheetName'])['values']
            for idx, row in enumerate(file_data_tab):
                if idx == 0:
                    row.insert(0, 'AcademicYear')
                    row.insert(0, 'Subject')
                    row.insert(0, 'District')                    
                else:
                    row.insert(0, file['AcademicYear'])
                    row.insert(0, file_source)
                    row.insert(0, file['ParentDirectory'])
            context['file_data'] += file_data_tab
    
    # First pass: Remove header rows
    excel_headers = None
    for idx, record in reversed (list (enumerate (context['file_data']))):
        if record[3] == 'School Name':
            excel_headers = [item.replace(' ', '') if isinstance(item, str) else item for item in record]
            for idx_header, header in enumerate (excel_headers):
                if 'ReferralSubject' in header:
                    excel_headers[idx_header] = 'ReferralType'
                if 'DistrictStudentUserId' in header:
                    excel_headers[idx_header] = 'StudentId'
                if 'Gender' in header:
                    excel_headers[idx_header] = 'Gender'    
                if 'DOB' in header:
                    excel_headers[idx_header] = 'DateOfBirth'
                if 'Ethnicity' in header:
                    excel_headers[idx_header] = 'Ethnicity'
                if 'ESOL' in header:
                    excel_headers[idx_header] = 'ESL'
                if 'SchoolNPSIS' in header:
                    excel_headers[idx_header] = 'SchoolNPSIS'
                if 'NPSCode' in header:
                    excel_headers[idx_header] = 'NPSCode'
            context['file_data'].pop(idx)
        if record[6] == '':
            context['file_data'].pop(idx)
    
    # Check if headers were found
    if excel_headers is None:
        error_msg = f"{file_source} Error: Header row not found in Excel file. Expected to find 'School Name' in column 4 of the header row. Contact {help_desk_email}."
        logger.error(error_msg)
        messages.error(request, error_msg)
        return render(request, 'graph_connector_app/file_data.html', context)
    
    # Validate Excel headers match database columns
    db_columns = [c.name for c in sm.Eligibility.__table__.columns]
    extra_columns = [col for col in excel_headers if col not in db_columns]
    
    if extra_columns:
        error_msg = f"{file_source} column mismatch: "
        if extra_columns:
            error_msg += f"Incorrect column name in Excel: {', '.join(extra_columns)}. "
        error_msg += f"Contact {help_desk_email}."
        logger.error(error_msg)
        messages.error(request, error_msg)
        return render(request, 'graph_connector_app/file_data.html', context)
    
    # Second pass: Convert data types
    for idx, record in enumerate (context['file_data']):
        if not isinstance(record[7], str): #StudentGrade
            record[7] = str(record[7])
        if isinstance(record[14], str): #DateOfBirth
            record[14] = None  
        else:
            record[14] = excel_date_to_python(record[14]) 


    db = sm.db

    try:
        trans = db.connection.begin()
        db.connection.execute("TRUNCATE TABLE AI.EligibilityStaging")
        trans.commit() 
    except Exception as e:
        logger.error(f"Error truncating {file_source} staging table: {type(e).__name__}: {str(e)}")
        raise e

    if context['file_data']:
        #columns = [c.name for c in sm.Eligibility.__table__.columns]
        columns = excel_headers
        records = []
        for row in context['file_data']:
            if len(row) != len(columns):
                raise ValueError(f"{file_source}: Row length {len(row)} does not match columns length {len(columns)}")
            records.append(dict(zip(columns, row)))


        if records:
            try:                
                trans = db.connection.begin()
                db.connection.execute(sm.Eligibility.__table__.insert(), records)
            except Exception as e:
                logger.error(f"Error inserting {file_source} records: {type(e).__name__}: {str(e)}")
                raise e

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

def get_ims_data(request):
    context = initialize_context(request)

    #Get User Information List from SharePoint via Graph API
    try:
        # Get app token for Microsoft Graph/SharePoint access
        token = get_token_for_app(request)
        logger.info(f"Token obtained: {token[:20] if token else 'None'}...")
        list_name = 'User Information List'

        # Use Graph API to fetch user data from SharePoint site
        sharepoint_data = get_sharepoint_users_via_graph(
            token, 
            'learnbehavioral.sharepoint.com', 
            '/sites/testingroup3',
            list_name
        )
        
        logger.info(f"SharePoint: {list_name} response keys: {sharepoint_data.keys() if isinstance(sharepoint_data, dict) else 'Not a dict'}")

        # Process the data
        records = []
        if 'value' in sharepoint_data:
            for item in sharepoint_data['value']:
                record = {}
                # Map available fields from SharePoint lists to database structure
                fields = item.get('fields', {})
                record['ContentTypeID'] = item.get('contentType', {}).get('id') or fields.get('ContentTypeId', None)
                record['Name'] = fields.get('Title', None)
                record['ComplianceAssetId'] = fields.get('_ComplianceTag', None)
                record['Account'] = fields.get('Name', None)
                record['EMail'] = fields.get('EMail') or fields.get('Email', None)
                record['OtherMail'] = fields.get('OtherMail', None)
                record['UserExpiration'] = parser.parse(fields['UserExpiration']) if fields.get('UserExpiration') else None
                record['UserLastDeletionTime'] = parser.parse(fields['UserLastDeletionTime']) if fields.get('UserLastDeletionTime') else None
                record['MobileNumber'] = fields.get('MobilePhone', None)
                record['AboutMe'] = fields.get('Notes', None)
                record['SIPAddress'] = fields.get('SipAddress') or fields.get('SIPAddress', None)
                record['IsSiteAdmin'] = fields.get('IsSiteAdmin', None)
                record['Deleted'] = fields.get('Deleted', None)
                record['Hidden'] = fields.get('UserInfoHidden', None)
                picture_val = fields.get('Picture', None)
                if isinstance(picture_val, dict):
                    desc = picture_val.get('Description')
                    url = picture_val.get('Url')
                    joined = ",".join([p for p in (desc, url) if p])
                    record['Picture'] = joined or None
                else:
                    record['Picture'] = picture_val
                record['Department'] = fields.get('Department', None)
                record['JobTitle'] = fields.get('JobTitle', None)
                record['FirstName'] = fields.get('FirstName', None)
                record['LastName'] = fields.get('LastName', None)
                record['WorkPhone'] = fields.get('WorkPhone', None)
                record['UserName'] = fields.get('UserName', None)
                record['WebSite'] = item.get('webUrl', None)
                record['AskMeAbout'] = fields.get('SPSResponsibility', None)
                record['Office'] = fields.get('Office', None)
                # prefer top-level timestamps from Graph item, fall back to fields
                modified_str = item.get('lastModifiedDateTime') or fields.get('Modified')
                record['Modified'] = parser.parse(modified_str) if modified_str else None
                id_val = item.get('id') or fields.get('id') or fields.get('ID')
                try:
                    record['Id'] = int(id_val) if id_val is not None else None
                except Exception:
                    record['Id'] = id_val
                record['ContentType'] = item.get('contentType', {}).get('name') or fields.get('ContentType', None)
                created_str = item.get('createdDateTime') or fields.get('Created') or item.get('Created')
                record['Created'] = parser.parse(created_str) if created_str else None
                createdById = fields.get('AuthorLookupId')
                try:
                    record['CreatedById'] = int(createdById) if createdById is not None else None
                except Exception:
                    record['CreatedById'] = createdById
                modifiedById = fields.get('EditorLookupId')
                try:
                    record['ModifiedById'] = int(modifiedById) if modifiedById is not None else None
                except Exception:
                    record['ModifiedById'] = modifiedById
                # Extract version from eTag (second element after comma)
                etag_val = item.get('eTag')
                if etag_val:
                    etag_clean = etag_val.strip('"')
                    etag_parts = etag_clean.split(',')
                    record['Owshiddenversion'] = etag_parts[1].strip() if len(etag_parts) > 1 else None
                else:
                    record['Owshiddenversion'] = None
                record['Version'] = fields.get('_UIVersionString', None)
                # Extract path from webUrl: remove domain and last segment
                weburl = item.get('webUrl')
                if weburl:
                    path_part = weburl.replace('https://learnbehavioral.sharepoint.com', '')
                    # Remove last URL segment
                    if '/' in path_part:
                        path_part = path_part.rsplit('/', 1)[0]
                    record['Path'] = path_part or None
                else:
                    record['Path'] = None
                
                records.append(record)
        
        logger.info(f"{list_name} - Total records processed: {len(records)}")

        # Convert records to list format for display
        display_records = []
        for record in records:
            display_records.append([
                record.get('Id', ''),
                record.get('Name', ''),
                record.get('EMail', ''),
                record.get('UserName', '')
            ])
        
        context['file_data'] = display_records

        # Only insert to database if we have records
        if records:
            # Truncate and load data
            db = sm.DatabaseConnection("Integration")

            trans = db.connection.begin()
            db.connection.execute(sm.sa.text("TRUNCATE TABLE [IMS].[UserInformationListStaging]"))
            trans.commit()
            logger.info(f"{list_name.replace(' ', '')}Staging truncated successfully")

            # Bulk insert data (much faster than one-by-one)
            if records:
                trans = db.connection.begin()
                db.connection.execute(sm.UserInformation.__table__.insert(), records)
                trans.commit()
            logger.info(f"{list_name.replace(' ', '')} - Inserted {len(records)} records")

            db.connection.close()
        else:
            logger.warning(f"{list_name} - No records to insert")

    except Exception as e:
        logger.error(f"Error in get_ims_data - {list_name}: {type(e).__name__}: {str(e)}")
        logger.error(traceback.format_exc())
        context['errors'] = [{'message': f'Error fetching {list_name} data: {str(e)}'}]

    #Get Asset Management List from SharePoint via Graph API
    try:
        # Get app token for Microsoft Graph/SharePoint access
        token = get_token_for_app(request)
        logger.info(f"Token obtained: {token[:20] if token else 'None'}...")
        list_name = 'Asset Management List'

        # Use Graph API to fetch user data from SharePoint site
        sharepoint_data = get_sharepoint_users_via_graph(
            token, 
            'learnbehavioral.sharepoint.com', 
            '/sites/testingroup3',
            list_name
        )
        
        logger.info(f"SharePoint: {list_name} response keys: {sharepoint_data.keys() if isinstance(sharepoint_data, dict) else 'Not a dict'}")

        # Process the data
        records = []
        if 'value' in sharepoint_data:
            for item in sharepoint_data['value']:
                record = {}
                fields = item.get('fields', {})
                # Map available fields from SharePoint lists to database structure
                id_val = item.get('id') or fields.get('id') or fields.get('ID')
                try:
                    record['Id'] = int(id_val) if id_val is not None else None
                except Exception:
                    record['Id'] = id_val
                record['ContentTypeID'] = item.get('contentType', {}).get('id') or fields.get('ContentTypeId', None)
                record['ContentType'] = fields.get('ContentType', None)
                record['Title'] = fields.get('Title', None)
                # prefer top-level timestamps from Graph item, fall back to fields
                modified_str = item.get('lastModifiedDateTime') or fields.get('Modified')                
                record['Modified'] = parser.parse(modified_str) if modified_str else None
                created_str = item.get('createdDateTime') or fields.get('Created')
                record['Created'] = parser.parse(created_str) if created_str else None      
                record['CreatedById'] = fields.get('AuthorLookupId', None)
                record['ModifiedById'] = fields.get('EditorLookupId', None)
                record['Owshiddenversion'] = fields.get('_UIVersionString', None)
                record['Version'] = fields.get('_UIVersionString', None)
                # Extract path from webUrl: remove domain and last segment
                weburl = item.get('webUrl')
                if weburl:
                    path_part = weburl.replace('https://learnbehavioral.sharepoint.com', '')
                    # Remove last URL segment
                    if '/' in path_part:
                        path_part = path_part.rsplit('/', 1)[0]
                    record['Path'] = path_part or None
                else:
                    record['Path'] = None
                record['ComplianceAssetId'] = fields.get('_ComplianceTag', None)
                record['StatusValue'] = fields.get('Status', None)
                record['Manufacturer'] = fields.get('Manufacturer', None)
                record['Model'] = fields.get('Model', None)
                record['ColorValue'] = fields.get('Color', None)
                record['TabletNumber'] = fields.get('SerialNumber', None)
                record['CurrentOwnerId'] = fields.get('CurrentOwnerLookupId', None)
                record['PreviousOwnerId'] = fields.get('PreviousOwnerLookupId', None)
                record['DueDate'] = parser.parse(fields['DueDate']) if fields.get('DueDate') else None
                record['CurrentOwnerPreviousOwnerId'] = fields.get('Current_x0020_Owner_x002c__x0020LookupId', None)
                record['DateAssigned'] = parser.parse(fields['Dateassigned']) if fields.get('Dateassigned') else None
                record['DateReachedOutForCollection'] = parser.parse(fields['Datereachedoutforcollection']) if fields.get('Datereachedoutforcollection') else None
                record['StaffLastAssignedToEmailId'] = fields.get('Assign_x0020_to', None)
                record['AssignedById'] = fields.get('AssignedbyLookupId', None)
                record['LocationValue'] = fields.get('Location', None)
                record['TrackingNumber'] = fields.get('Tracking_x0020_Number', None)
                record['HasAWorkingCharger'] = fields.get('HasaWorkingCharger', None)
                record['StaffLastAssignedToFullName'] = fields.get('StaffLastAssignedto_x0028_fullna', None)
                record['ColorTag'] = fields.get('Colortag', None)
                record['ActivityCount'] = fields.get('ActivityCount', None)
                
                records.append(record)
        
        logger.info(f"{list_name} - Total records processed: {len(records)}")

        # Convert records to list format for display
        display_records = []
        for record in records:
            display_records.append([
                record.get('Id', ''),
                record.get('TabletNumber', ''),
                record.get('LocationValue', ''),
                record.get('StatusValue', '')
            ])
        
        context['file_data'] = display_records

        # Only insert to database if we have records
        if records:
            # Truncate and load data
            db = sm.DatabaseConnection("Integration")

            trans = db.connection.begin()
            db.connection.execute(sm.sa.text("TRUNCATE TABLE [IMS].[AssetManagementListStaging]"))
            trans.commit()
            logger.info("IMS.AssetManagementListStaging truncated successfully")

            # Bulk insert data (much faster than one-by-one)
            if records:
                trans = db.connection.begin()
                db.connection.execute(sm.AssetManagement.__table__.insert(), records)
                trans.commit()
            logger.info(f"{list_name} - Inserted {len(records)} records")

            db.connection.close()
        else:
            logger.warning(f"{list_name} - No records to insert")

    except Exception as e:
        logger.error(f"Error in get_ims_data - {list_name}: {type(e).__name__}: {str(e)}")
        logger.error(traceback.format_exc())
        context['errors'] = [{'message': f'Error fetching {list_name} data: {str(e)}'}]

    return render(request, 'graph_connector_app/file_data.html', context)

def debug(request):
    context = initialize_context(request)
        
    print(settings.CSRF_TRUSTED_ORIGINS)
    return render(request, 'graph_connector_app/file_data.html', context)