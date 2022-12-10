"""
This file takes data from a xlsx file and uploads it to a mongo database.
"""

from pymongo import MongoClient
import openpyxl


"""
User Values. 
These are the values that the user will need to change to get the script to work.
"""
columns = [0, 10, 12]  # The columns in the excel file that you want to upload to mongo.
keys = [  # These are the keys that will be used in the mongo database.
    "granite_id",
    "device_ip",
    "hostname",
    ]
EXCEL_FILE = "E:\\Chrome Downloads\\20221207_SNMP_Unreachable.xlsx"
MONGO_DB = "granite"  # The name of the mongo database you want to write to.
MONGO_COLLECTION = "devices"  # The name of the mongo collection you want to write to.

HEADER = "EQUIP_INST_ID"  # This is the header of the first column in the excel file.

KEY = keys[1]  # The "primary key" of the collection. The unique identifier for each document (in keys).
"""
End User Values
"""


client = MongoClient(
        'mongodb://g_admin:graniteinventory2022@172.20.26.130:27017,172.20.26.131:27017/admin?replicaSet=LNE&authMechanism=DEFAULT&tls=false',
        27017)
db = client[MONGO_DB]
collection = db[MONGO_COLLECTION]


wb = openpyxl.load_workbook(EXCEL_FILE)
sheet = wb.active

for row in sheet.iter_rows():
    if row[0].value != HEADER:
        data = {keys[index]: row[columns[index]].value for index in range(len(columns))}

        collection.replace_one({KEY: data[KEY]}, data, upsert=True)
        print(data)