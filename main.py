import json
import os
from icon import icon_data_uri
import zipfile
import io
import base64
import pandas as pd
import os, urllib3, requests
from os import path
from os import listdir
from os.path import isfile, join
from google.cloud import storage
from datetime import datetime

# single function to handle action api routing
def route_handler(request):
    routes = {
        "/": [action_list],
        "/tabbed_excel_download/form": [action_form],
        "/tabbed_excel_download/execute": [action_execute],
        "/status": [action_status]
    }
    try:
        route_handlers = routes[request.path] 
        # or [route_not_found]
        for handler in route_handlers:
            handler_response = handler(request)
            if not handler_response:
                continue
            return handler_response
    except:
        return {
            'status': 401,
            'body':{
                'error': 'There was an error with handler'
            }
        }

# action list endpoint

def action_list(request):
    payload = request.get_json()
    print("list_actions received request:", payload)
    
    actions_list = {
      "integrations": [
        {
          "name": "tabbed_excel_download",
          "label": "Tabbed Excel (xlsx) Download",
          "description": "Download a dashboard as a tabbed Excel file",
           "form_url": os.environ.get('CALLBACK_URL_PREFIX','') + '/tabbed_excel_download/form',
           "url": os.environ.get('CALLBACK_URL_PREFIX','') + '/tabbed_excel_download/execute',
          "supported_action_types": ["dashboard"],
          "supported_formats": ["csv_zip"],
          "icon_data_uri": icon_data_uri,
          # "params": [
          #   {
          #     "name": "include_vis", 
          #     "label": "Include Visualizations", 
          #     "required": True, 
          #     "sensitive": False
          #   }
          # ],
          "supported_download_settings": ["push"]
        }
      ]
    }

    return actions_list

# action form endpoint
def action_form(request):
    payload = request.get_json()
    print("action_form received request:", payload)
    print("request data for form: ", request)
  
    form = [
        {
        "name": "Include Cover Page", 
        "label": "Include page with title, time of download, and filters", 
        "type":"select",
        "required": True, 
        "sensitive": False, 
        "options": [
            {
            "name": "include_yes",
            "label": "Yes"
            },
            {
            "name": "include_no", 
            "label": "No"
            },
        ]
        },
        {
        "name": "Email addresses",
        "type": "text",
        "required": True
        }
    ]

    return json.dumps(form)


def action_execute(request):
    request = request.get_json()
    print('action_execute payload:', request)
    data = request['attachment']['data']
    title = request['scheduled_plan']['title']
    title = title.replace(" ", "_").lower()

    decoded_file = base64.b64decode(data)
    zfp = zipfile.ZipFile(io.BytesIO(decoded_file), "r")
    
    root = path.dirname(path.abspath(__file__))
    children = os.listdir(root)
    files = [c for c in children if path.isfile(path.join(root, c))]
    print('Files: {}'.format(files))
    writer = pd.ExcelWriter("/tmp/sample.xlsx", engine='xlsxwriter')
       
    for f in zfp.namelist():
        df = str(zfp.open(f).read(), 'utf-8')
        df = pd.read_csv(io.StringIO(df))
        df.to_excel(writer, sheet_name=os.path.split(f)[1].split('.')[0][0:30], index=False)
        print(os.path.split(f)[1].split('.')[0][0:30], ' file written')

    writer.save()

    onlyfiles = [f for f in listdir('/tmp') if isfile(join('/tmp', f))]

    print('Files in tmp: ', onlyfiles)
    now = datetime.now()
 
    dt_string = now.strftime("%d-%m-%Y-%H:%M:%S")


    storage_client = storage.Client()
    bucket = storage_client.bucket("tabbed-excel-files")
    blob = bucket.blob("excel-files/"+str(title) + str(dt_string))

    blob.upload_from_filename("/tmp/sample.xlsx")

    return {"success": 200}

# check action hub status
def action_status(request):
    return