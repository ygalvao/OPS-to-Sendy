#!/usr/bin/env python3

#*******************************************************************************#
# Copyright 2023 - Ramsons Enterprises Inc. <webadmin@printprodigital.com>      #
#                                                                               #
# This project cannot be copied and/or distributed without the express          #
# permission of Ramsons Enterprises Inc.                                        #
#                                                                               #
# Unauthorized copying of this file, via any medium is strictly prohibited.     #
#                                                                               #
# Written by Yuri H. Galvao <yuri@galvao.ca>, March 2023                        #
#*******************************************************************************#

from ops_manipulator import *
from pandarallel import pandarallel
import requests as req

def load_api_credentials()->str:
    """"""

    SENDY_API_FILE = 'config/sendy_api.conf'

    if check_file(SENDY_API_FILE) and confirm('Do you want to use the last used API data for Sendy? '):
        api_key = json.load(open(SENDY_API_FILE, 'r'))['api key']
        sendy_url = json.load(open(SENDY_API_FILE, 'r'))['url']
    else:
        api_data = ask_for_data(('api key', 'url'), 'sendy_api')
        api_key, sendy_url = api_data['api key'], api_data['url']

    return api_key, sendy_url

def scrap_and_get_dfs(ops_credentials:dict, driver:object)->list:
    """"""

    logging.info('Webscrapping OPS website and extracting data regarding subscribers and subscribes...\n')

    # Declaring necessary variables
    username = ops_credentials['username']
    password = ops_credentials['password']
    login_uri = ops_credentials['login_uri']

    # Login page
    login(login_uri, username, password, driver)

    # Dashboard page
    customer_link = WebDriverWait(driver, timeout=7).until(lambda x: x.find_element(by=By.LINK_TEXT, value='Customer'))
    customer_link.click()
    subscribers_link = WebDriverWait(driver, timeout=7).until(lambda x: x.find_element(by=By.LINK_TEXT, value='Newsletter Subscribers'))
    subscribers_link.click()
    time.sleep(2.5)

    # Newsletter Subscribers page
    ## Importing tables indo Pandas DataFrames
    subscribers_df = pd.read_html(driver.page_source)[0]
    while not WebDriverWait(driver, timeout=1).until(lambda x: x.find_element(by=By.XPATH, value='//li[@class="paginate_button page-item next disabled"]')).is_displayed():
        next_button = WebDriverWait(driver, timeout=7).until(lambda x: x.find_element(by=By.XPATH, value='//*[@id="ops-table_next"]/a'))
        next_button.click()
        time.sleep(1.5)
        subscribers_df_temp = pd.read_html(driver.page_source)[0]
        subscribers_df = pd.concat([subscribers_df, subscribers_df_temp], ignore_index=True)
    
    unsubscribes_df = pd.read_html(driver.page_source)[1]

    driver.quit()

    assert subscribers_df.shape[0] > 0, "The subscribers DataFrame is empty!"

    subscribers_df['Registered Date'] = pd.to_datetime(subscribers_df['Registered Date'])
    subscribers_df.drop(columns='Sr#', inplace=True)

    logging.info('The lists of subscribers (and unsubscribes, if there is any) were successfully imported from OPS into Pandas DataFrames!\n')

    return [subscribers_df, unsubscribes_df]

def fetch_statuses(df:pd.DataFrame, main_list_id:str, sendy_url:str, api_key:str)->pd.DataFrame:
    """"""

    pandarallel.initialize(nb_workers = 6, progress_bar=True)

    logging.info('Checking status in Sendy for each email...\n')

    df['sendy_status'] = df.Email.parallel_apply(lambda email: req.post(
        sendy_url + '/api/subscribers/subscription-status.php',
        {
            'api_key' : api_key,
            'email' : email,
            'list_id' : main_list_id,
        }
        ).text)
    
    print()
    
    return df

def export_to_sendy(dfs:tuple, list_id:str, sendy_url:str, api_key:str)->tuple:
    """"""

    subscribers_df = dfs[0]
    new_subs_df = subscribers_df[subscribers_df.sendy_status.str.contains('does not exist')].copy()
    unsubscribes_df = dfs[1]
    success = 0
    error = 0

    for index, row in new_subs_df.iterrows():
        response = req.post(
            sendy_url + '/subscribe',
            {
                'api_key' : api_key,
                'name' : f'''{row['First Name']} {row['Last Name']}'''.strip().title(),
                'email' : row.Email,
                'list' : list_id,
                'Optin_time' : row['Registered Date'],
                'boolean' : 'true',
                # 'Country' : ,
            }
            ).text
        
        if response in (True, 'true', 1, '1'):
            logging.info(f'{row.Email} was successfully inserted into Sendy!')
            success += 1
        else:
            logging.error(f'Error when inserting {row.Email}: {response}')
            error += 1

    print()
    if success > 0 or error > 0:
        logging.info(f'''Report: {success} new subscribers were successfully entered into Sendy and {error} new subscribers couldn't be entered.\n''')
    else:
        logging.info(f'''Report: no new subscribers.\n''')

    return success, error

def main(api_credentials:tuple, dfs:list,)->None:
    """"""

    api_key = api_credentials[0]
    sendy_url = api_credentials[1]

    # Using Sendy's API to extract and load data
    ## Pulling the id of the main list of the main (first) brand
    main_list_id = req.post(
        sendy_url + '/api/lists/get-lists.php',
        {
            'api_key' : api_key,
            'brand_id' : '1',
        }
        ).text
    
    main_list_id = json.loads(main_list_id)
    main_list_id = main_list_id['list1']['id']

    ## Extracting statuses from Sendy
    dfs[0] = fetch_statuses(dfs[0], main_list_id, sendy_url, api_key)

    ## Loading relevant data from subscribers and/or unsubscribes through Sendy's API

    success, error = export_to_sendy(dfs, main_list_id, sendy_url, api_key)
    if error == 0 and success > 0:
        logging.info('The new subscribers and the unsubscribes (if there was any) were successfully exported from OPS to Sendy!\n')
    if error > 0:
        logging.error('Error! Please, check the details above and try again later.\n')

if __name__ == '__main__':
    print('###########################################\n')
    logging.info('Starting the program.\n')
    main(load_api_credentials(), scrap_and_get_dfs(ops_credentials, setup_driver()))

    logging.info('End of the program.\n')

