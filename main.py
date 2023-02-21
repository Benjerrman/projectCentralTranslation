import datetime
import logging
import os

from googletrans import Translator
from shareplum import Site
from shareplum import Office365
from shareplum.site import Version

# set up logging
log_filename = f'Logs\\{datetime.datetime.now().strftime("%a %b %d %I-%M-%S %p")}.log'
logging.basicConfig(filename=log_filename, level=logging.INFO)

# SharePoint credentials
username = "flowbot@choicehumanitarian.org"
password = "IShouldBeHashed19!"
site_url = "https://choicems.sharepoint.com"
site_name = "chprojects"

# set up SharePoint connection
authcookie = Office365(site_url, username=username, password=password).GetCookies()
site = Site(site_url, version=Version.v365, authcookie=authcookie)

# set up Google Translate client
translator = Translator()

# list and column names
project_list_name = "Projects"
project_columns_to_translate = ["Title", "Description", "Objective", "Outcome", "Status", "Location", "Contact"]
request_list_name = "Requests"
request_columns_to_translate = ["Title", "Description", "Objective", "Outcome", "Status", "Location", "Contact", "Requester", "Approver", "Priority", "Category", "Subcategory", "Target Date", "Due Date"]
status_update_list_name = "Status Updates"
status_update_columns_to_translate = ["Title", "Status", "Update"]

# function to translate a single record
def translate_record(record, columns_to_translate):
    for column_name in columns_to_translate:
        if column_name in record and record[column_name] and record["language_submitted"] != "en":
            translated = translator.translate(record[column_name], src="es", dest="en").text
            record[column_name] = translated
    return record

# function to translate all records in a list
def translate_records(list_name, columns_to_translate):
    list_obj = site.List(list_name)
    records = list_obj.GetListItems()
    translated_records = [translate_record(record, columns_to_translate) for record in records]
    return translated_records

# translate all records
project_records = translate_records(project_list_name, project_columns_to_translate)
request_records = translate_records(request_list_name, request_columns_to_translate)
status_update_records = translate_records(status_update_list_name, status_update_columns_to_translate)

# append translated records to lists of dictionaries
project_dicts = [dict(record) for record in project_records]
request_dicts = [dict(record) for record in request_records]
status_update_dicts = [dict(record) for record in status_update_records]

# update SharePoint lists with translated records
project_list_obj = site.List(project_list_name)
project_list_obj.UpdateListItems(data=project_dicts)
logging.info(f"Updated {len(project_dicts)} records in {project_list_name} list.")

request_list_obj = site.List(request_list_name)
request_list_obj.UpdateListItems(data=request_dicts)
logging.info(f"Updated {len(request_dicts)} records in {request_list_name} list.")

status_update_list_obj = site.List(status_update_list_name)
status_update_list_obj.UpdateListItems(data=status_update_dicts)
logging.info(f"Updated {len(status_update_dicts)} records in {status_update_list_name} list.")