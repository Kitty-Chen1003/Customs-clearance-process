import json

import requests

import global_state
from utils.path import get_resource_path


def read_config(config_path):
    with open(config_path, 'r') as file:
        config = json.load(file)
    return config


log_path = get_resource_path("config/config.json")
url = read_config(log_path)['api']['base_url']


def add_awb_request(flight, carrier, manifest_path):
    # headers = {'Content-Type': 'application/json'}
    data = {
        "flight": flight,
        "defaultCarrier": carrier,
    }
    files = {
        "file": open(manifest_path, 'rb')
    }

    try:
        response = requests.post(url + '/Ecargo/addAWB', data=data, files=files)
        return response
        # pass

    except requests.exceptions.RequestException as e:
        # 捕获所有请求相关的异常，如网络连接问题、超时等
        print(f"Request failed: {e}")
        return None


def add_arn_request():
    # headers = {'Content-Type': 'application/json'}
    data = {
        "flight": global_state.flight,
        "mawb": global_state.airway_bill
    }
    files = {
        "file": open(global_state.process_file_path, 'rb')
    }

    try:
        response = requests.post(url + '/Ecargo/addArn', data=data, files=files)
        return response
        # pass

    except requests.exceptions.RequestException as e:
        # 捕获所有请求相关的异常，如网络连接问题、超时等
        print(f"Request failed: {e}")
        return None


def add_ens_request():
    # headers = {'Content-Type': 'application/json'}
    data = {
        "flight": global_state.flight,
        "mawb": global_state.airway_bill
    }
    files = {
        "file": open(global_state.process_file_path, 'rb')
    }

    try:
        response = requests.post(url + '/Ecargo/addEns', data=data, files=files)
        return response

    except requests.exceptions.RequestException as e:
        # 捕获所有请求相关的异常，如网络连接问题、超时等
        print(f"Request failed: {e}")
        return None


def get_dsk_request():
    # headers = {'Content-Type': 'application/json'}
    data = {
        "flight": global_state.flight,
        "mawb": global_state.airway_bill
    }
    files = {
        "manifestFile": open(global_state.manifest_file_path, 'rb'),
    }

    try:
        response = requests.post(url + '/Ecargo/getDSK', data=data, files=files)
        return response
        # pass

    except requests.exceptions.RequestException as e:
        # 捕获所有请求相关的异常，如网络连接问题、超时等
        print(f"Request failed: {e}")
        return None


def init_items_clearances_request():
    files = {
        "manifestFile": open(global_state.manifest_file_path, 'rb')
    }

    try:
        response = requests.post(url + '/Ecargo/init', files=files)
        return response

    except requests.exceptions.RequestException as e:
        # 捕获所有请求相关的异常，如网络连接问题、超时等
        print(f"Request failed: {e}")
        return None


def add_items_clearances_request():
    data = {
        "MAWB": global_state.airway_bill,
        "flight": global_state.flight
    }
    files = {
        "file": open(global_state.process_file_path, 'rb')
    }

    try:
        response = requests.post(url + '/Ecargo/addItemsClereances', data=data, files=files)
        return response

    except requests.exceptions.RequestException as e:
        # 捕获所有请求相关的异常，如网络连接问题、超时等
        print(f"Request failed: {e}")
        return None


def add_items_releases_request():
    data = {
        "MAWB": global_state.airway_bill,
        "flight": global_state.flight
    }

    files = {
        "file": open(global_state.process_file_path, 'rb')
    }

    try:
        response = requests.post(url + '/Ecargo/addItemsRelease', data=data, files=files)
        return response

    except requests.exceptions.RequestException as e:
        # 捕获所有请求相关的异常，如网络连接问题、超时等
        print(f"Request failed: {e}")
        return None


def download_zc415_file(airway_bill):
    headers = {'Content-Type': 'application/json'}
    data = {
        "airWayBill": airway_bill
    }

    try:
        response = requests.post(url + '/getZC415', headers=headers, data=data)
        return response

    except requests.exceptions.RequestException as e:
        # 捕获所有请求相关的异常，如网络连接问题、超时等
        print(f"Request failed: {e}")
        return None
