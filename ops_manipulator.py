#!/usr/bin/env python3

#*******************************************************************************#
# Copyright 2023 - Ramsons Enterprises Inc. <webadmin@printprodigital.com>      #
#                                                                               #
# This file is part of OPS Invoice Checker and OPS-QBO integration projects.    #
#                                                                               #
# OPS Invoice Checker and OPS-QBO integration projects can not be copied and/or #
# distributed without the express permission of Ramsons Enterprises Inc.        #
#                                                                               #
# Unauthorized copying of this file, via any medium is strictly prohibited      #
#                                                                               #
# Written by Yuri H. Galvao <yuri@galvao.ca>, November 2022                     #
#*******************************************************************************#

# Importing the necessary libraries
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
from basic_functions import *
import numpy as np
import pandas as pd

# Declaring constants
LINUX_USERAGENT = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"
OPTIONS = Options()
headless = False if '--show-browser' in args else True if '--headless' in args else confirm('Do you want the browser to be headless (i.e., you won\'t see it)? ')
if headless:
    OPTIONS.add_argument('--headless')

OPTIONS.add_argument(f'user-agent={LINUX_USERAGENT}')
OPS_CREDENTIALS_FILE = 'config/ops_credentials.conf'

# Declaring some variables
if check_file(OPS_CREDENTIALS_FILE) and confirm('Do you want to use the last used data for OPS? '):
    ops_credentials = json.load(open(OPS_CREDENTIALS_FILE, 'r'))
else:
    ops_credentials = ask_for_data(('username', 'password', 'login_uri'), 'ops_credentials')

login_uri = ops_credentials['login_uri']

# Defining the functions
def get_payment_response(driver:object)->dict:
    """"""
    try:
        time.sleep(.25)
        payment_response = WebDriverWait(driver, timeout=3).until(lambda x: x.find_element(by=By.XPATH, value='''(//a[@title='Payment Response'])'''))
        time.sleep(.25)
        payment_response = payment_response.get_attribute('data-content')
        payment_response = payment_response.split('"')
    except:
        return

    clean_payment_response_list = []
    for i, element in enumerate(payment_response):
        if i%2 != 0:
            clean_payment_response_list.append(element)

    try: #If the payment is from Moneris rather than QBP, change some elements to avoid bugs/error later on
        clean_payment_response_list[clean_payment_response_list.index('Message')] = 'status'
        clean_payment_response_list[clean_payment_response_list.index('TransAmount')] = 'amount'
        clean_payment_response_list[clean_payment_response_list.index('APPROVED           *                    =')] = 'CAPTURED'
    except:
        try:
            clean_payment_response_list[clean_payment_response_list.index('confirmed')] = 'CAPTURED'
        except:
            pass
    
    payment_response = {}
    for i, element in enumerate(clean_payment_response_list):
        if i%2 == 0:
            try:
                payment_response[element] = clean_payment_response_list[i+1]
            except:
                pass
    
    return payment_response

def comment(driver:object, inv_n:int)->None:
    """"""

    try:
        if inv_n:
            comments_textarea = WebDriverWait(driver, timeout=3).until(lambda x: x.find_element(by=By.ID, value='comments'))
            actions = ActionChains(driver)
            actions.move_to_element(comments_textarea).perform()
            comments_textarea.send_keys(f'Invoice # {inv_n}')
    except Exception as e:
        logging.error(f'Error! Exception: {repr(e)}')

def close_wo(driver:object, wo_number:int, inv_n:int=None)->dict:
    """"""

    try: # Tries to close the work order
        payment_data = None
        time.sleep(.75)
        view_orders_link = WebDriverWait(driver, timeout=1).until(lambda x: x.find_element(by=By.LINK_TEXT, value='View Orders'))
        view_orders_link.click()
        time.sleep(.75)
        first_production_status = 'Prepress'
        try:
            previous_statuses = WebDriverWait(driver, timeout=3).until(lambda x: x.find_elements(by=By.XPATH, value='''//div[@class='card blk-ord-prd-history history_0']'''))[1:]
        except:
            previous_statuses = [first_production_status]
        else:
            previous_statuses = [x.find_element(by=By.XPATH, value='''.//span[@class='text-secondary']''').text for x in previous_statuses]

        update_order_link = WebDriverWait(driver, timeout=1).until(lambda x: x.find_element(by=By.LINK_TEXT, value='Update Order'))
        update_order_link.click()
        status_select = WebDriverWait(driver, timeout=3).until(lambda x: x.find_element(by=By.ID, value='''status'''))
        status_select = Select(status_select)
        #cod = WebDriverWait(driver, timeout=1).until(lambda x: x.find_element(by=By.XPATH, value='''//select[option='Ready for Invoicing (C.O.D.)']''')).is_selected()
        selected = status_select.first_selected_option.text

        if selected in ('Ready for Invoicing (C.O.D.)', 'Ready for Invoicing (Pre-Payment)', 'Ready for Invoicing (before delivery)'):
            unfinished = True
            while unfinished:
                try:
                    status_select.select_by_visible_text(previous_statuses[0])
                except:
                    previous_statuses.pop(0)
                    continue
                else:
                    unfinished = False
        else:
            status_select.select_by_visible_text('Delivered, Invoiced, Closed')

        comment(driver, inv_n)
        save_button = WebDriverWait(driver, timeout=1).until(lambda x: x.find_element(by=By.ID, value='''btn-action-save'''))
        save_button.click()
        time.sleep(.75)
        view_orders_link = WebDriverWait(driver, timeout=3).until(lambda x: x.find_element(by=By.LINK_TEXT, value='View Orders'))
        view_orders_link.click()
        payment_data = get_payment_response(driver)
        if payment_data:
            payment_data['wo_n'] = wo_number
            payment_data['inv_n'] = inv_n
    except Exception as e:
        logging.critical(repr(e))
    else: # W.O. status changed to "Delivered, Invoiced, Closed" successfully
        if selected in ('Ready for Invoicing (C.O.D.)', 'Ready for Invoicing (Pre-Payment)', 'Ready for Invoicing (before delivery)'):
            logging.info(f'''W.O. # {wo_number} status was changed to "{previous_statuses[0]}" successfully!\n''')
        else:
            logging.info(f'''W.O. # {wo_number} status was changed to "Delivered, Invoiced, Closed" successfully!\n''')

    return payment_data
    
def setup_driver()->object:
    """"""

    driver = webdriver.Chrome(options = OPTIONS)
    driver.set_window_size(1600,900)

    return driver
    
def login(login_uri:str, username:str, password:str, driver:object)->None:
    """Logs in OPS. 
    
    Arguments:
        login_uri: string
        username: string
        password: string
        driver: Selenium webdriver instance object

    Returns:
        Nothing."""

    driver.get(login_uri)
    username_input = WebDriverWait(driver, timeout=5).until(lambda x: x.find_element(by=By.ID, value='username'))
    password_input = driver.find_element(by=By.ID, value='password')
    login_button = driver.find_element(by=By.TAG_NAME, value='button')
    username_input.send_keys(username)
    password_input.send_keys(password)
    login_button.click()

def manipulate_ops_payments(
    payments_dict:dict,
    login_uri:str=login_uri,
    credentials:dict=ops_credentials
    )->dict:
    """"""

    ## Declaring local variables and objects
    driver = setup_driver()
    username = credentials['username']
    password = credentials['password']
    found = True
    customers_not_found = {}

    ## Login page
    login(login_uri, username, password, driver)

    ## Dashboard page
    customer_nav = WebDriverWait(driver, timeout=5).until(lambda x: x.find_element(by=By.XPATH, value='''//i[@class='nav-icon far fa-user']'''))
    customer_nav.click()

    ## Iterating payments_dict
    for i, (key, value) in enumerate(payments_dict.items()):
        customer_name = key
        amount_paid = str(value[0])
        inv_n = str(value[1])

        ### Still in Dashboard page (or "Website Customers" page, if this is not the first iteration)
        if found:
            time.sleep(1)
            customers_link = WebDriverWait(driver, timeout=7).until(lambda x: x.find_element(by=By.LINK_TEXT, value='Website Customers'))
            customers_link.click()

        ### Website Customers" page
        time.sleep(.5)
        search_input = WebDriverWait(driver, timeout=3).until(lambda x: x.find_element(by=By.ID, value='keyword'))

        if found == False:
            search_input.clear()

        search_input.send_keys(customer_name)
        search_input.submit()

        try:
            time.sleep(.75)
            action_dropdown = WebDriverWait(driver, timeout=3).until(lambda x: x.find_element(by=By.XPATH, value='''(//button[@class='btn btn-white btn-inverse dropdown-toggle btn-sm'])'''))
            #actions = ActionChains(driver)
            #actions.move_to_element(action_dropdown)
            #actions.perform()
        except:
            found = False
            customers_not_found[customer_name] = value
            continue
        else:
            action_dropdown.click()
            time.sleep(.25)
            found = True            
            pay_on_account_link = WebDriverWait(driver, timeout=7).until(lambda x: x.find_element(by=By.LINK_TEXT, value='Manage Pay On Account'))
            pay_on_account_link.click()
            time.sleep(1.5)
            paid_input = WebDriverWait(driver, timeout=3).until(lambda x: x.find_element(by=By.ID, value='amount'))
            paid_input.send_keys(amount_paid)
            paid_comments_input = WebDriverWait(driver, timeout=3).until(lambda x: x.find_element(by=By.ID, value='comments'))
            paid_comments_input.send_keys(f'Invoice #: {inv_n}')
            paid_submit = WebDriverWait(driver, timeout=3).until(lambda x: x.find_element(by=By.ID, value='btnSubmit'))
            paid_submit.click()
            time.sleep(.5)
            breakpoint()
            logging.info(f'The paid amount of ${value} was successfully entered for {customer_name} in OPS!')

def manipulate_ops_wo(
    wo_numbers:list,
    credentials:dict,
    mode:str='check', # It can be: 'check' or 'checker'; 'insert' or 'i'; or 'insert bills'
    invoices_numbers:list=None,
    bills_data:list=None,
    login_uri:str=login_uri,
    order_type:str='List Orders' # It can be: 'List Orders', 'Archive Orders'
)->list:
    """
    """
    
    ## Declaring local variables and objects
    driver = setup_driver()
    not_found = False
    wo_and_inv_numbers = {}
    username = credentials['username']
    password = credentials['password']
    mode = mode.lower()
    wo_not_found = []
    
    def shorten_collection(collection:list or dict, index_or_key:int or str)->None:
        """"""
        
        collection.pop(index_or_key)

    ## Login page
    login(login_uri, username, password, driver)

    ## Dashboard page
    orders_nav = WebDriverWait(driver, timeout=5).until(lambda x: x.find_element(by=By.XPATH, value='''//i[@class='nav-icon far fa-shopping-cart']'''))
    orders_nav.click()
    
    ## Checking data to avoid errors and mistakes
    if invoices_numbers != None and mode != 'insert bills':
        if len(wo_numbers) != len(invoices_numbers):
                    logging.error('''The number of work orders doesn't match the number of invoices!''')
                    time.sleep(1)
                    driver.quit()
                    return

    ## Iterating the work orders list
    for i, wo_number in enumerate(wo_numbers):
        wo_n_str = str(wo_number) # Converts to string
        
        ### Still in Dashboard page (or "List Orders" / "Archive Orders" page, if this is not the first iteration)
        if not_found == False:
            time.sleep(1)
            orders_link = WebDriverWait(driver, timeout=7).until(lambda x: x.find_element(by=By.LINK_TEXT, value=order_type))
            orders_link.click()
        
        ### "List Orders" or "Archive Orders" page
        time.sleep(.5)
        search_input = WebDriverWait(driver, timeout=3).until(lambda x: x.find_element(by=By.ID, value='search_string'))
        
        if not_found:
            search_input.clear()
            
        search_input.send_keys(wo_n_str)
        search_input.submit()
        #action_dropdown = driver.find_element(by=By.XPATH, value='''//button[@class='btn btn-white btn-inverse dropdown-toggle btn-sm']''')
        #action_dropdown.click()     # Uncomment this section if you want/need to use the dropdown   
        try:
            time.sleep(1)
            if mode != 'insert bills':
                view_order_link = WebDriverWait(driver, timeout=3).until(lambda x: x.find_element(by=By.LINK_TEXT, value=wo_n_str))
            else:
                dropdown_button_chevron_down = WebDriverWait(driver, timeout=3).until(lambda x: x.find_element(by=By.XPATH, value='''(//button[@class='btn btn-sm btn-primary collapsed btn-minier'])[1]'''))
                dropdown_button_chevron_down.click()
                time.sleep(.5)
                view_printer_link = WebDriverWait(driver, timeout=3).until(lambda x: x.find_element(by=By.XPATH, value='''(//a[@class='lnk_ord_printer'])[1]'''))
        except:
            if mode != 'insert bills':
                logging.warning(f'''W.O. # {wo_number} wasn't found in {order_type}!''')

                if mode in ('insert', 'i'):
                    shorten_collection(invoices_numbers, 0)
            else:
                logging.warning(f'''W.O. # {wo_number} wasn't found in {order_type} or it doesn't appear in OPS as outsourced!''')
                shorten_collection(bills_data, 0)
            not_found = True            
            wo_not_found.append(wo_number)
            continue
        else:
            not_found = False
            if mode != 'insert bills':
                view_order_link.click()
            else:
                view_printer_link.click()
        if mode != 'insert bills':
            order_notes_link = WebDriverWait(driver, timeout=3).until(lambda x: x.find_element(by=By.LINK_TEXT, value='Order Notes'))
            order_notes_link.click()
        
        if mode in ('checker', 'check'): # The function will just get invoice information
            try: # Tries to get the invoice number, date, and user who inserted it
                time.sleep(.5)
                order_notes_invoice_row = WebDriverWait(driver, timeout=3).until(lambda x: x.find_element(by=By.XPATH, value='''//tr[td='Invoice']'''))
            except:
                logging.info(f'No invoice was found for W.O. # {wo_number}.')
                wo_and_inv_numbers[wo_number] = ''
            else: # Cleans the data and save it for output
                invoice_info_list = []
                invoice_info_raw_list = order_notes_invoice_row.text.split('\n')[-2:]
                invoice_info_raw_list[0] = invoice_info_raw_list[0].split(' ')
                invoice_number = invoice_info_raw_list[0][0]
                logging.info(f'Invoice # {invoice_number} found for W.O. # {wo_number}!')
                invoice_note_date = ''.join(date_info+' ' for date_info in invoice_info_raw_list[0][1:])[:-1] # Date and time 
                invoice_info_list.extend([invoice_number, invoice_note_date, invoice_info_raw_list[-1]])
                wo_and_inv_numbers[wo_number] = invoice_info_list
                if order_type =='List Orders':
                    payment_data = close_wo(driver, wo_number)

        elif mode in ('insert', 'i'): # The function will try to insert the invoice number and close the work order
            invoice_n = invoices_numbers[0]
            try: # Tries to insert the invoice number
                try: # Tries to get the invoice number, date, and user who inserted it
                    time.sleep(.5)
                    order_notes_invoice_row = WebDriverWait(driver, timeout=3).until(lambda x: x.find_element(by=By.XPATH, value='''//tr[td='Invoice']'''))
                except:
                    logging.info(f'No invoice # was found for W.O. # {wo_number} in OPS.')
                else:
                    invoice_info_list = []
                    invoice_info_raw_list = order_notes_invoice_row.text.split('\n')[-2:]
                    invoice_info_raw_list[0] = invoice_info_raw_list[0].split(' ')
                    invoice_number = invoice_info_raw_list[0][0]
                    logging.info(f'Invoice # {invoice_number} found for W.O. # {wo_number} in OPS!')
                    if order_type =='List Orders':
                        payment_data = close_wo(driver, wo_number)
                    continue
                    
                # The code below will only be executed if there is no invoice inserted in OPS for that W.O.
                time.sleep(.2)
                add_order_notes_select = WebDriverWait(driver, timeout=3).until(lambda x: x.find_element(by=By.ID, value='''order_note_category'''))
                #add_order_notes_select.click()
                add_order_notes_select = Select(add_order_notes_select)
                add_order_notes_select.select_by_visible_text('Invoice')
                textarea_input = WebDriverWait(driver, timeout=1).until(lambda x: x.find_element(by=By.ID, value='''ordernotecomments'''))
                textarea_input.send_keys(str(invoice_n))
                save_and_back_button = WebDriverWait(driver, timeout=1).until(lambda x: x.find_element(by=By.ID, value='''btn-action-saveback'''))
                save_and_back_button.click()
            except Exception as e:
                logging.critical(e)
            else: # Invoice number inserted successfully
                logging.info(f'Invoice # {invoice_n} was inserted for W.O. # {wo_number} in OPS.')
                payment_data = close_wo(driver, wo_number, invoice_n)
            finally:
                shorten_collection(invoices_numbers, 0)

        elif mode == 'insert bills':
            bill_n = bills_data[0][0]
            supplier_name = str(bills_data[0][1])
            bill_date = str(bills_data[0][2])
            due_date = str(bills_data[0][3])
            shipping_cost = str(bills_data[0][4])
            docket_n = str(bills_data[0][5])
            quote_n = str(bills_data[0][6])
            bill_subtotal = str(bills_data[0][7])
            bill_amt = str(bills_data[0][8])
            try: # Tries to insert the invoice number
                try: # Tries to get the invoice number, date, and user who inserted it
                    time.sleep(.5)
                    order_history_bill_table = WebDriverWait(driver, timeout=3).until(lambda x: x.find_element(by=By.XPATH, value='''//table[@class='dataTable table table-striped table-bordered table-hover']'''))
                    ## Extracting and transforming the data found in the page
                    bill_info_list = []
                    columns_texts = []
                    values_rows = []                   
                    for row in order_history_bill_table.find_elements(by=By.XPATH, value='.//tr'):
                        columns_list = []
                        values_list = []
                        
                        columns_list.append(row.find_elements(by=By.XPATH, value='./th'))
                        for columns in columns_list:
                            for column in columns:
                                columns_texts.append(column.text)
                            
                        values_list.append(row.find_elements(by=By.XPATH, value='./td'))
                        for values in values_list:
                            values_row_texts = []
                            for value in values:
                                values_row_texts.append(value.text)
                            values_rows.append(values_row_texts)
                    
                    values_rows.remove([])
                    ops_bills_df = pd.DataFrame(values_rows, columns=columns_texts)
                    ops_bills_df.drop(columns='Notified', inplace=True)
                    ops_bills_df.Products = ops_bills_df.Products.apply(lambda x: x.partition('\n')[0])
                    ops_bills_df.Comments = ops_bills_df.Comments.apply(lambda x: [s for s in x.splitlines()] if type(x) is str else np.nan)
                    ops_bills_df['Bill #'] = ops_bills_df.Comments.apply(lambda x: [''.join(filter(str.isdigit, s.lower().partition('invoice')[2])) for s in x])
                    ops_bills_df['Bill #'] = ops_bills_df['Bill #'].apply(lambda x: [n for n in x if n != ''] if (type(x) is list and len(x) > 0 and x != '') else np.nan)
                    ops_bills_df['Bill #'] = ops_bills_df['Bill #'].apply(lambda x: x[0] if (type(x) is list and len(x) > 0 and x != '') else np.nan)

                    ## Checking if there is a match
                    if (ops_bills_df['Bill #'] != str(bill_n)).any():
                        raise Exception

                except Exception as e:
                    logging.info(f'No Order History was found for W.O. # {wo_number} and Bill # {bill_n} in OPS.')
                else:
                    try:                
                        logging.info(f'''Bill # {ops_bills_df[ops_bills_df['Bill #'] == str(bill_n)].iloc[0]['Bill #']} found for W.O. # {wo_number} in OPS!''')
                    except Exception as e:
                        logging.critical(f'Error! Exception: {repr(e)}')
                    continue

                time.sleep(1)
                #add_order_history_select = WebDriverWait(driver, timeout=3).until(lambda x: x.find_element(by=By.XPATH, value='''(//select[@id='order_comments_message'])[1]'''))
                #add_order_history_select = Select(add_order_history_select)
                #add_order_history_select.select_by_visible_text('Invoiced')
                textarea_input = WebDriverWait(driver, timeout=1).until(lambda x: x.find_element(by=By.XPATH, value='''(//textarea)[1]'''))
                textarea_input.send_keys(f'Invoice #: {str(bill_n)}\nSupplier: {supplier_name}\nInvoice subtotal: ${bill_subtotal}\nInvoice total amount: ${bill_amt}\nBill date: {bill_date}\nShipping cost: {shipping_cost}\nDocket #: {docket_n}\nQuote #: {quote_n}')
                save_and_back_button = WebDriverWait(driver, timeout=1).until(lambda x: x.find_element(by=By.ID, value='''btn-action-saveback'''))
                save_and_back_button.click()
            except Exception as e:
                logging.critical(f'Error when inserting bill data for W.O. # {wo_number} in OPS! Exception: {repr(e)}')
            else: # bill number inserted successfully
                logging.info(f'Bill # {bill_n} was inserted for W.O. # {wo_number} in OPS.')
            finally:
                shorten_collection(bills_data, 0)
        #else:

    driver.quit()
    
    if mode in ('checker', 'check'):
        return wo_and_inv_numbers

    if mode == 'insert bills':
        return wo_not_found

    if mode == 'insert':
        return payment_data

if __name__ == '__main__':
    print('###########################################\n')
    logging.info('Starting the program.\n')

    wo_type = 'List Orders'
    mode = 'check'
    chosen_mode = input('Choose the mode ([c]heck / [i]nsert / insert [b]ills): ').strip().lower()
    work_orders = list_from_input('Please, enter the w.o. numbers separated by comma (e.g. 7001,7009): ')

    if chosen_mode in ('i', 'insert'):
        mode = 'insert'
        invoices = list_from_input('Please, enter the invoices numbers separated by comma (e.g. 14452, 14485): ')
        wo_inv = manipulate_ops_wo(work_orders, invoices_numbers=invoices, credentials=ops_credentials, order_type=wo_type, mode=mode)
    elif chosen_mode in ('b', 'insert bills'):
        mode = 'insert bills'
        list_of_bills_data = []
        for i, wo in enumerate(work_orders):
            print('This is the template for inserting new data from bills from outsourced services: bill #, bill date, shipping cost, docket #, quote #, bill amt.')
            bills_data = list_from_input(f'Please, enter the bill data for the W.O. # {wo}, with each data separated by commas (e.g. 64452, 2022-11-02, N/A, N/A, N/A, 1075.00): ')
            list_of_bills_data.append(bills_data)
        chosen_wo_type = input('What kind of w.o. do you want to check ([l]ist orders / [a]rchived orders): ').strip().lower()
        if chosen_wo_type not in ('l', 'list', 'list orders'):
            wo_type = 'Archive Orders'            
        wo_inv = manipulate_ops_wo(work_orders, bills_data=list_of_bills_data, credentials=ops_credentials, order_type=wo_type, mode=mode)
    else:
        chosen_wo_type = input('What kind of w.o. do you want to check ([l]ist orders / [a]rchived orders): ').strip().lower()
        if chosen_wo_type not in ('l', 'list', 'list orders'):
            wo_type = 'Archive Orders'
        
        wo_inv = manipulate_ops_wo(work_orders, credentials=ops_credentials, order_type=wo_type, mode=mode)

    logging.info('End of the program.\n')
