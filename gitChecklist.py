import time

import openpyxl

# Enter the name of the file of the sheet that you wish to modify.
file_name = 'iAM-MIX_AoIP Traveler XXXXXX.xlsx'

# These are pretend values that will be given by Jim & Juan's program
# and passed into the fill_sheet function as a dictionary.
d = {}

d['travelerName'] = 'iAM-MIX (AoIP)'
d['serialNumber'] = '159639'
d['MACaddress'] = 'fc:69:47:94:9a:25'

d['channel'] = 8

d['systemStatus'] = ['2.7-9', '2.8-39', 'A3.4B']

d['loaded'] = [True, True, True]

d['licenseKeys'] = ['EEProm', 'AES', 'Analog', 'ST-2110-2']


# This function fills the values on the given 'Traveler' sheet which can
# be filled automatically.
def fill_sheet(dic, file):
    # Open the excel sheet for editing.
    book = openpyxl.load_workbook(file)
    sheets = book.sheetnames
    sheet = book[sheets[1]]

    # Fill in the unit type, S/N and MAC Address.
    sheet['A1'] = dic['travelerName'] + ' Traveler'
    sheet['E3'] = dic['serialNumber']
    sheet['E5'] = dic['MACaddress']

    # Mark down whether this is an 8 or 16 channel unit.
    if dic['channel'] == 8:
        sheet['A9'] = 'X'
    else:
        sheet['A8'] = 'X'

    # System Status is filled in here with strings from the dictionary.
    sheet['L8'] = dic['systemStatus'][0]
    sheet['L9'] = dic['systemStatus'][1]
    sheet['L10'] = dic['systemStatus'][2]

    # Automated parts of the Pre Test include 'Load All Licenses,'
    # 'Load FPGA,' and 'Load Software.'
    if dic['loaded'][0]:
        sheet['M14'] = 'X'
        sheet['V14'] = 'X'
    if dic['loaded'][1]:
        sheet['M15'] = 'X'
        sheet['V15'] = 'X'
    if dic['loaded'][2]:
        sheet['M16'] = 'X'
        sheet['V16'] = 'X'

    # Mark an 'X' next to each License Key that the customer has paid
    # to enable.
    if 'EEProm' in dic['licenseKeys']:
        sheet['A39'] = 'X'
    if 'AES' in dic['licenseKeys']:
        sheet['A40'] = 'X'
    if 'Analog' in dic['licenseKeys']:
        sheet['A41'] = 'X'
    if 'AoIP' in dic['licenseKeys']:
        sheet['A42'] = 'X'
    if 'MADI BNC' in dic['licenseKeys']:
        sheet['A43'] = 'X'
    if 'MADI Optical' in dic['licenseKeys']:
        sheet['A44'] = 'X'
    if 'SDI-1' in dic['licenseKeys']:
        sheet['A45'] = 'X'
    if 'SDI-2' in dic['licenseKeys']:
        sheet['A46'] = 'X'
    if 'ST-2110-1' in dic['licenseKeys']:
        sheet['A47'] = 'X'
    if 'ST-2110-2' in dic['licenseKeys']:
        sheet['A48'] = 'X'
    if 'ST-2022-6/7-1' in dic['licenseKeys']:
        sheet['A49'] = 'X'
    if 'ST-2022-6/7-2' in dic['licenseKeys']:
        sheet['A50'] = 'X'
    if 'Output Routing' in dic['licenseKeys']:
        sheet['A51'] = 'X'
    # Indicate that this process was automated.
    sheet['D39'] = 'X'

    # Note the date of completion for this checklist or 'Traveler.'
    sheet['Q51'] = time.strftime("%x")

    book.save(file)
    return


fill_sheet(d, file_name)
