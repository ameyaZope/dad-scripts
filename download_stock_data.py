import requests
import json
from tqdm import tqdm
import csv
from datetime import datetime


def get_companies_list(gl_type, indx_grp_val):
    headers = {
        'authority': 'api.bseindia.com',
        'sec-ch-ua': '" Not;A Brand";v="99", "Google Chrome";v="91", "Chromium";v="91"',
        'accept': 'application/json, text/plain, */*',
        'sec-ch-ua-mobile': '?1',
        'user-agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Mobile Safari/537.36',
        'origin': 'https://www.bseindia.com',
        'sec-fetch-site': 'same-site',
        'sec-fetch-mode': 'cors',
        'sec-fetch-dest': 'empty',
        'referer': 'https://www.bseindia.com/',
        'accept-language': 'en-GB,en-US;q=0.9,en;q=0.8',
    }
    url = f"https://api.bseindia.com/BseIndiaAPI/api/MktRGainerLoserData/w?GLtype={gl_type}&IndxGrp=group&IndxGrpval={indx_grp_val}&orderby=all"
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    if response.status_code == 200:
        return json.loads(response.text)['Table']
    else:
        return []


def get_company_top_data():
    headers = {
        'authority': 'api.bseindia.com',
        'pragma': 'no-cache',
        'cache-control': 'no-cache',
        'sec-ch-ua': '" Not;A Brand";v="99", "Google Chrome";v="91", "Chromium";v="91"',
        'accept': 'application/json, text/plain, */*',
        'sec-ch-ua-mobile': '?1',
        'user-agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Mobile Safari/537.36',
        'origin': 'https://www.bseindia.com',
        'sec-fetch-site': 'same-site',
        'sec-fetch-mode': 'cors',
        'sec-fetch-dest': 'empty',
        'referer': 'https://www.bseindia.com/',
        'accept-language': 'en-GB,en-US;q=0.9,en;q=0.8',
    }
    url = 'https://api.bseindia.com/BseIndiaAPI/api/getScripHeaderData/w?Debtflag=&scripcode=532890&seriesid='
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    if response.status_code == 200:
        return json.loads(response.text)
    else:
        return {}


def get_company_main_data(scrip_code):
    headers = {
        'authority': 'api.bseindia.com',
        'sec-ch-ua': '" Not;A Brand";v="99", "Google Chrome";v="91", "Chromium";v="91"',
        'accept': 'application/json, text/plain, */*',
        'sec-ch-ua-mobile': '?1',
        'user-agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Mobile Safari/537.36',
        'origin': 'https://www.bseindia.com',
        'sec-fetch-site': 'same-site',
        'sec-fetch-mode': 'cors',
        'sec-fetch-dest': 'empty',
        'referer': 'https://www.bseindia.com/',
        'accept-language': 'en-GB,en-US;q=0.9,en;q=0.8',
        'if-modified-since': 'Fri, 25 Jun 2021 17:12:37 GMT',
    }
    url = f"https://api.bseindia.com/BseIndiaAPI/api/HighLow/w?Type=EQ&scripcode={scrip_code}"
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    if response.status_code == 200:
        return json.loads(response.text)
    else:
        return {}

def convert_json_to_csv(json_array, csv_file_path):
    # Extract field names from the first object in the JSON array
    field_names = list(json_array[0].keys())

    # Open the CSV file in write mode
    with open(csv_file_path, 'w', newline='') as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=field_names)

        # Write the header row
        writer.writeheader()

        # Write each object in the JSON array as a row in the CSV file
        writer.writerows(json_array)


def main():
    gl_types_list = ['loser', 'gainer']
    indx_grp_val_list = ['A', 'B']

    all_selected_companies = []

    for gl_type in gl_types_list:
        for indx_grp_val in indx_grp_val_list:
            print('\nFetching gainerLooserType:', gl_type, 'and indexGroup:', indx_grp_val)
            company_list = get_companies_list(gl_type, indx_grp_val)
            print('\nFetched', len(company_list), 'companies of gainerLooserType:', gl_type, 'and indexGroup:',
                  indx_grp_val)
            print('Processing above mentioned categories companies')

            for company in tqdm(company_list):
                company_main_data = get_company_main_data(company['scrip_cd'])
                required_company_data = {}
                required_company_data['scrip_cd'] = company['scrip_cd']
                required_company_data['scrip_grp'] = company['scrip_grp']
                required_company_data['longName'] = company['LONG_NAME']
                required_company_data['last_trade_rate'] = company['ltradert']
                required_company_data['mean_of_fifty2Week'] = (float(company_main_data['Fifty2WkHigh_adj']) + float(
                    company_main_data['Fifty2WkLow_adj'])) / 2
                required_company_data['percentage_above_52_week_low'] = (100 * (
                            company['ltradert'] - float(company_main_data['Fifty2WkLow_adj']))) / float(company_main_data[
                                                                            'Fifty2WkLow_adj'])
                required_company_data['change_val'] = company['change_val']
                required_company_data['change_percentage'] = company['change_percent']
                required_company_data['fifty2WkHigh'] = company_main_data['Fifty2WkHigh_adj']
                required_company_data['fifty2WkLow'] = company_main_data['Fifty2WkLow_adj']
                required_company_data['openrate'] = company['openrate']
                required_company_data['highrate'] = company['highrate']
                required_company_data['lowrate'] = company['lowrate']
                required_company_data['date_time'] = company['dt_tm']
                required_company_data['url'] = company['URL']
                all_selected_companies.append(required_company_data)

    print('Attempting to create Excel Sheet')
    convert_json_to_csv(all_selected_companies, str(datetime.now()) + ".csv")
    print('Excel sheet creation completed')

if __name__ == '__main__':
    main()