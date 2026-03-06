import json
import requests

import global_state
from company_api import CompanyAPI
from utils.path import get_resource_path


def read_config(config_path):
    with open(config_path, 'r') as file:
        return json.load(file)


class EcargoAPI(CompanyAPI):

    def __init__(self):
        config_path = get_resource_path("config/config.json")
        config = read_config(config_path)
        self.base_url = config['api']['base_url'] + '/Ecargo'

    # def steps(self):
    #     return ["AddAwb", "AddArn", "AddEns", "GetDsk", "AddItemsClearances", "AddItemsReleases"]

    def add_awb(self, flight, carrier, manifest_path):
        data = {
            "flight": flight,
            "defaultCarrier": carrier,
        }
        files = {
            "file": open(manifest_path, 'rb')
        }
        try:
            response = requests.post(f"{self.base_url}/addAWB", data=data, files=files)
            return response

        except requests.exceptions.RequestException as e:
            print(f"add_awb request failed: {e}")
            return None

    def add_arn(self):
        data = {
            "flight": global_state.flight,
            "mawb": global_state.airway_bill
        }
        files = {
            "file": open(global_state.process_file_path, 'rb')
        }
        try:
            response = requests.post(f"{self.base_url}/addArn", data=data, files=files)
            return response

        except requests.exceptions.RequestException as e:
            print(f"add_arn request failed: {e}")
            return None

    def add_ens(self):
        data = {
            "flight": global_state.flight,
            "mawb": global_state.airway_bill
        }
        files = {
            "file": open(global_state.process_file_path, 'rb')
        }
        try:
            response = requests.post(f"{self.base_url}/addEns", data=data, files=files)
            return response

        except requests.exceptions.RequestException as e:
            print(f"add_ens request failed: {e}")
            return None

    def get_dsk(self):
        data = {
            "flight": global_state.flight,
            "mawb": global_state.airway_bill
        }
        files = {
            "manifestFile": open(global_state.manifest_file_path, 'rb'),
        }
        try:
            response = requests.post(f"{self.base_url}/getDSK", data=data, files=files)
            return response

        except requests.exceptions.RequestException as e:
            print(f"get_dsk request failed: {e}")
            return None

    def init_items_clearances(self):
        files = {
            "manifestFile": open(global_state.manifest_file_path, 'rb')
        }
        try:
            response = requests.post(f"{self.base_url}/init", files=files)
            return response

        except requests.exceptions.RequestException as e:
            print(f"init_items_clearances request failed: {e}")
            return None

    def add_items_clearances(self):
        data = {
            "MAWB": global_state.airway_bill,
            "flight": global_state.flight
        }
        files = {
            "file": open(global_state.process_file_path, 'rb')
        }

        try:
            response = requests.post(f"{self.base_url}/addItemsClereances", data=data, files=files)
            return response

        except requests.exceptions.RequestException as e:
            print(f"add_items_clearances request failed: {e}")
            return None

    def add_items_releases(self):
        data = {
            "MAWB": global_state.airway_bill,
            "flight": global_state.flight
        }

        files = {
            "file": open(global_state.process_file_path, 'rb')
        }
        try:
            response = requests.post(f"{self.base_url}/addItemsRelease", data=data, files=files)
            return response

        except requests.exceptions.RequestException as e:
            print(f"add_items_releases request failed: {e}")
            return None

    def download_zc415(self, airway_bill):
        headers = {'Content-Type': 'application/json'}
        data = {
            "airWayBill": airway_bill
        }

        try:
            response = requests.post(f"{self.base_url}/downloadZC415", headers=headers, data=data)
            return response

        except requests.exceptions.RequestException as e:
            print(f"download_zc415 request failed: {e}")
            return None