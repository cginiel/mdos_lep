from openpyxl import Workbook
from openpyxl import load_workbook
import requests
import json
import secrets # file that contains your API key

###### global vars #########
MAPQUEST_KEY = secrets.CONSUMER_KEY
RESOURCE_URL = "http://www.mapquestapi.com/search/v2/radius?"

CACHE_FILENAME = "michigan_LEP_cache.json"
CACHE_DICT = {}

mdos_addresses_xlsx = "mdos-building-addresses.xlsx"
lep_by_county_xlsx = "lep-by-county-michigan.xlsx"

############################

###################### caching ##########################
def open_cache():
    ''' Opens the cache file if it exists and loads the JSON into
    the CACHE_DICT dictionary.
    if the cache file doesn't exist, creates a new cache dictionary
    
    Parameters
    ----------
    None
    
    Returns
    -------
    dict
        the opened cache
    '''
    try:
        cache_file = open(CACHE_FILENAME, 'r')
        cache_contents = cache_file.read()
        cache_dict = json.loads(cache_contents)
        cache_file.close()
    except:
        cache_dict = {}
    return cache_dict


def save_cache(cache_dict):
    ''' Saves the current state of the cache to disk
    
    Parameters
    ----------
    cache_dict: dict
        The dictionary to save
    
    Returns
    -------
    None
    '''
    dumped_json_cache = json.dumps(cache_dict)
    fw = open(CACHE_FILENAME,"w")
    fw.write(dumped_json_cache)
    fw.close() 


def make_request_with_cache(key, value):
    '''Issues a request to the cache saved to the device.

    If the item exists in the cache, the program will pull from that data.
    If the item does not exist in the cache, the program will create a key/value
    pair and save it to the cache dictionary

    Parameters
    ----------
    key
        list
    value
        list

    '''

    if key in CACHE_DICT.keys():
        print("Using cache")
        return CACHE_DICT[key]
    else:
        print("Fetching")
        CACHE_DICT[key] = value
        save_cache(CACHE_DICT)
        return CACHE_DICT[key]

################################################

def import_workbook(filename):
    '''imports an excel workbook as a file python can work with.

    params
    ------
    filename : str
        string that points to .xlsx file

    returns
    -------
    wb : object
        excel workbook object
    '''
    wb = load_workbook(filename)

    return wb


def get_mdos_building_zipcodes():
    '''reads the mdos-building-addresses.xlsx file and extracts zipcodes that can
    eventually be sent through the mapquest API in order to find out which county
    each zipcode is in.

    params
    ------
    none

    returns
    -------
    zipcode_list : list
        list of zipcodes as they correspond to the row on the excel sheet. 
        e.g., C2 = 49221; C145 = 48202
    '''
    zipcode_list = []
    ## importing our workbook
    branches = import_workbook(mdos_addresses_xlsx)

    ## accessing a specific worksheet
    address_sheet = branches['Address']

    ## access zipcode from each address cell
    for address in address_sheet['C2':'C145']:
        full_address = address[0].value

        ## some zipcodes have the extension so I cleaned that off
        if "-" in full_address[-10:]:
            zipcode = full_address[-10:-5]
            zipcode_list.append(zipcode)
        else:
            zipcode = full_address[-5:]
            zipcode_list.append(zipcode)

    return zipcode_list
    

def search_for_county_with_zipcode(zipcode):
    '''searches a zipcode on the mapquest API and returns a corresponding county.

    params
    ------
    zipcode : int

    returns
    -------
    county : str
        the county associated with the zipcode provided
    '''
    params = {
    "key" : MAPQUEST_KEY,
    "origin" : zipcode,
    "radius" : 10,
    "maxMatches" : 10,
    "ambiguities" : "ignore",
    "outFormat" : "json"
    }

    param_strings = []
    connector = '&'
    for k in params.keys():
        param_strings.append(f'{k}={params[k]}')
    param_strings.sort()
    unique_key = RESOURCE_URL + connector.join(param_strings)
    response = requests.get(unique_key).json()
    county = response['origin']['adminArea4']
    make_request_with_cache(zipcode, county)

    return county


def make_county_list_from_zipcode():
    '''Uses the county/zipcode match from search_for_county_with_zipcode() to create a county list
    that matches each MDOS building from the mdos-building-addresses.xlsx file.

    params
    ------
    none

    returns
    -------
    county_list : list
        list of counties as they correspond to the mdos-building-addresses.xlsx rows.
    '''
    county_list = []
    for zipcode in get_mdos_building_zipcodes():
        county = search_for_county_with_zipcode(zipcode)
        if county == "":
            county = "Saginaw County"
        else:
            county_list.append(county)
    print(county_list)


if __name__ == "__main__":
    CACHE_DICT = open_cache()
    make_county_list_from_zipcode()