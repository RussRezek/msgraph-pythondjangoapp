# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

from django.urls import path

from . import views

urlpatterns = [
  # /
  path('', views.home, name='home'),
  path('signin', views.sign_in, name='signin'),
  path('signout', views.sign_out, name='signout'),
  path('calendar', views.calendar, name='calendar'),
  path('callback', views.callback, name='callback'),
  path('calendar/new', views.new_event, name='newevent'),
  path('ai_files', views.ai_files, name='ai_files'),
  path('file_data/<str:file_id>/<str:worksheet_name>', views.file_data, name = 'file_data'),
  path('get_all_reading_iready', views.get_all_reading_iready, name='get_all_reading_iready'),
  path('get_all_math_iready', views.get_all_math_iready, name='get_all_math_iready'),
  path('get_all_eligibility', views.get_all_eligibility, name='get_all_eligibility'),
  path('get_picker', views.get_picker, name='get_years'),
  path('load_tables', views.load_tables, name='load_tables'),
  path('get_districts', views.get_districts, name='get_districts')
]
