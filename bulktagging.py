#python function to read the excel file containing hostnames
import os
from dotenv import load_dotenv
from falconpy import Hosts,DeviceControlPolicies
import openpyxl
import sys

load_dotenv()

tag_action = sys.argv[1]
tag_name = sys.argv[2]
file_name = sys.argv[3]

falcon = Hosts(client_id=os.getenv("CLIENT_ID"),
               client_secret=os.getenv("CLIENT_SECRET")
        )

def hostnames_in_workbook(file_location):
    hostnames = []
    workbook = openpyxl.load_workbook(file_location)
    sheet = workbook.active
    for row in sheet.iter_rows():
        row_values = []
        for cell in row:
            row_values.append(cell.value)
        hostnames.append(row_values[0])
    return hostnames


def hostname_to_id(hostname):
    result = falcon.query_devices_by_filter(filter = f"hostname:'{hostname}'")
    return result["body"]["resources"][0]

def get_ids_for_hosts(hostnames):
    hostids = []
    for hostname in hostnames:
        hostid = hostname_to_id(hostname)
        hostids.append(hostid)
    return hostids

def get_tag_list(size):
    tag_list = [tag_name]*size
    return tag_list

def remove_falcon_tag_from_object(hostname,falconTag):
    host_id = hostname_to_id(hostname)
    response = falcon.update_device_tags(action_name="remove", id=host_id, tags=tag_list)

if __name__ == "__main__":
    hostnames = hostnames_in_workbook(file_location=file_name)
    id_list = get_ids_for_hosts(hostnames)
    tag_list = get_tag_list(len(hostnames))
    response = falcon.update_device_tags(action_name=tag_action, ids=id_list, tags=tag_list)
    print(response)

 