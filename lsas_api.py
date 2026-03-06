import json
import requests

import global_state
from company_api import CompanyAPI
from utils.path import get_resource_path


def read_config(config_path):
    with open(config_path, 'r') as file:
        return json.load(file)


class LSASAPI(CompanyAPI):

    def __init__(self):
        config_path = get_resource_path("config/config.json")
        config = read_config(config_path)
        self.base_url = config['api']['base_url'] + '/LSAS'

        self.token = None

    # def steps(self):
    #     return ["Login", "AddAwb", "AddEns", "MarkArrived", "AddArn", "AddItemsReleasesEasy", "GetDsk", "AddItemsClearancesEasy"]

    def login(self, username, password):
        data = {"username": username, "password": password}
        try:
            resp = requests.post(f"{self.base_url}/login", data=data)
            return resp

        except requests.exceptions.RequestException as e:
            print(f"login failed: {e}")
            return None

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

    def mark_arrived(self):
        data = {
            "flight": global_state.flight,
            "mawb": global_state.airway_bill
        }
        files = {
            "file": open(global_state.process_file_path, 'rb')
        }
        try:
            response = requests.post(f"{self.base_url}/markArrived", data=data, files=files)
            return response

        except requests.exceptions.RequestException as e:
            print(f"mark_arrived request failed: {e}")
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