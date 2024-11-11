import pdfplumber
import pandas as pd
from boxsdk import Client, OAuth2
import calendar
import argparse
import os
import sys


data = {
    'Price': [],
    'Address': [],
    'Meter': [],
    'Consumption': [],
    'Month': []
}

def parse_pdf(list_file_name):
    for file_name in list_file_name:
        path = "bills/" + file_name
        with pdfplumber.open(path) as pdf:
            size_doc = len(pdf.pages)
            for page in range(size_doc):
                p0 = pdf.pages[page]
                text = p0.extract_text(keep_blank_chars=True)
                output = open("output.txt",'w')
                output.write(text)
                output = open('output.txt', 'r')
                save = -1
                for num, line in enumerate(output.readlines()):
                    list_line = str(line).strip().split(' ')
                    # print(list_line)
                    if 'Total' in list_line and 'Current' in list_line and 'Activity' in list_line:
                        price = list_line[-1]
                        data['Price'].append(price)
                        # print(f'Price: {price} \n')

                    if 'Status:' in list_line:
                        slice = list_line.index('Status:')
                        address = " ".join(list_line[:slice])
                        data['Address'].append(address)

                        # print(f'Address: {address}')
                    if list_line == ['Meter', '#', 'Rate', 'Cycle', 'Days'] or list_line == ['Meter', '#', 'Rate', 'Cycle', 'Days', 'Demand']:
                        save = num + 2
                    if num == save:
                        if len(list_line[0]) == 5:
                            meter = list_line[0]
                            data['Meter'].append(int(meter))
                            # print(f'Meter: {meter}')
                        else: 
                            data['Meter'].append('NA')
                            data['Consumption'].append('NA')
                            # print('No Meter associated')

                    if 'Current' in list_line and 'Month' in list_line and 'Previous' not in list_line:
                        consumption = list_line[3]
                        data['Consumption'].append(int(consumption))
                        # print(f'Consumption: {consumption} kW')

                    if 'Billing' in list_line and 'Date:' in list_line:
                        month = calendar.month_name[int(list_line[-1][:2])]
            
            tot = len(data['Price'])
            for i in range(tot):
                data['Month'].append(month)

        return data

            

if __name__ == "__main__":
    #parse arguments
    parser = argparse.ArgumentParser("Parse info in pdf files")
    parser.add_argument("-f", help="folder name") 
    args = parser.parse_args()

    files = os.listdir(args.f)
    dic_data = parse_pdf(files)
    df = pd.DataFrame(data)
    df.to_excel('ElectricityInfo.xlsx', index = False)

    




# client_id = 'bswd6a0faank6k2jvg5jxfdpzovaikse'
# client_secret = '00WAQ25TWag9GWwZgJu0QhSjhXQanUi8'
# access_token = 'Ta7NVGDDYS8ScxRaorMrrDvw38KcHm5v'  



# # Define a callback function to save the new tokens
# def store_tokens_callback(new_access_token, new_refresh_token):
#     # Save the tokens securely (e.g., in environment variables, a secure file, or a database)
#     global access_token, refresh_token
#     access_token = new_access_token
#     refresh_token = new_refresh_token

# Initialize OAuth2 with the client credentials and token callback
# auth = OAuth2(client_id, client_secret, access_token)
# client = Client(auth)

# file_path = 'ElectricityInfo.xlsx'
# folder_id = '0'

# with open('ElectricityInfo.csv', 'rb') as file_content: 
    # client.folder(folder_id).upload_stream(file_content,file_path)
