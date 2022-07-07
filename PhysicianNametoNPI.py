# import module
import requests
import json
import openpyxl
class PhysicianName:
  def __init__(self, first_name, last_name, state, row):
    self.first_name = first_name
    self.last_name = last_name
    self.state = state
    self.row = row
# load excel with its path

#load excel to get list of fname, lname, state
def get_physician_name_from_excel(file):
    physician_names = []

    # load excel with its path
    wrkbk = openpyxl.load_workbook(file)

    sh = wrkbk.get_sheet_by_name('Triad HealthCare Network')

    # iterate through excel and display data
    for i in range(2, sh.max_row+1):
        physician_name = PhysicianName(sh.cell(row=i, column = 1).value, sh.cell(row=i, column=2).value, "NC", row = i)
        physician_names.append(physician_name)
    get_npi_lists(physician_names, file, wrkbk, sh)
#process each entry in the list
def get_npi_lists(physician_names, file, wrkbk, sh):
    for physician_name in physician_names:
        get_npi_from_name(physician_name, file, wrkbk, sh)
#write the number back into excel

def get_npi_from_name(physician_name, file, wrkbk, sh):
    api_url = "https://npiregistry.cms.hhs.gov/api/?version=2.0&&pretty=true&state={st}&first_name={fname}&last_name={lname}".format(st = physician_name.state,fname=physician_name.first_name, lname = physician_name.last_name)
    print(api_url)
    response = requests.get(api_url)
    if(response.json()["result_count"] > 0):
        #edge case of multiple doctors in the same state with the same physician_name
        #edge case if the doctor doesnt exist
        #edge case multiple last names
        #cellref=sh.cell(row=physician_name.row, column=5)
        #cellref=response.json()["results"][0]["number"]
        cell = 'E' + str(physician_name.row)
        sh[cell] = response.json()["results"][0]["number"]
        print(response.json()["results"][0]["number"])
    wrkbk.save("/Users/ptran/Downloads/ACOs.xlsx")

get_physician_name_from_excel("/Users/ptran/Downloads/ACOs.xlsx")



