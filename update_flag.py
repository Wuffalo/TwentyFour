# -*- coding: utf-8 -*-
"""
2021
@author: wuffalo
"""

import os
from datetime import datetime as dt
import json
import uuid

def flag_for_update():
   if update == 1:
      return True
   else:
      return False

output_directory = "/mnt/shared-drive/05 - Office/OTS/Wolf/"
output_file_name = "24Hour.xlsx"
path_to_output = output_directory+output_file_name

store_file = "/mnt/shared-drive/05 - Office/OTS/Wolf/stored_time.json"

path_to_sharedSOS = '/mnt/shared-drive/Operations/Data/Shipment Order Summary (PICK ZONE).csv'
file_time_shared = os.path.getctime(path_to_sharedSOS)

with open(store_file, 'r') as f:
    data = json.load(f)
    stored_time = data['time']

if file_time_shared > stored_time:
    file_time_best = file_time_shared
else:
    update = 1
    raise SystemExit

with open(store_file, 'r') as f:
    data = json.load(f)
    data['time'] = file_time_best
    data['readable'] = dt.fromtimestamp(file_time_best).strftime('%m/%d/%Y %H:%M')

# avoid interference with other thread request
tempfile = os.path.join(os.path.dirname(store_file), str(uuid.uuid1()))
with open(tempfile, 'w') as f:
    json.dump(data, f, indent=4)

# replace temporary file replacing old file
os.replace(tempfile, store_file)

# troubleshooting. Ouput not available in cron deployment. Extension: return output possibly as error code.
#print("Waveplanner updated SOS at " + dt.fromtimestamp(file_time_best).strftime('%m/%d/%Y %H:%M'))