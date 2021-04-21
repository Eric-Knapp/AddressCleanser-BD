# -*- coding: utf-8 -*-
"""
Created on Mon Feb  8 16:44:05 2021

@author: Eric Knapp
"""
import pandas as pd
import pandas_usaddress
import os

#load dataframe
df = pd.read_excel(r'C: YOUR INPUT FILE LOCATION') # input your location for reading your input file


#initiate usaddress
df = pandas_usaddress.tag(df, ['address'], granularity='full', standardize=True)
df['StreetNamePreDirectional'] = df['StreetNamePreDirectional'].str.upper() 
df['StateName'] = df['StateName'].str.upper() 
df["PlaceName"]= df["PlaceName"].str.upper().str.title() 
df["StreetName"]= df["StreetName"].str.upper().str.title() 
df["StreetNamePostDirectional"]= df["StreetNamePostDirectional"].str.upper().str.title() 
df["OccupancyType"]= df["OccupancyType"].str.upper().str.title() 
df["StreetNamePreType"]= df["StreetNamePreType"].str.upper().str.title() 
df["StreetNamePostType"]= df["StreetNamePostType"].str.upper().str.title() 
df["StreetNamePostDirectional"]= df["StreetNamePostDirectional"].str.upper()
df["OccupancyType"]= df["OccupancyType"].str.upper().str.title() 
df["Recipient"]= df["Recipient"].str.upper().str.title() 
df["OccupancyIdentifier"]= df["OccupancyIdentifier"].str.upper().str.title() 
df["SubaddressType"]= df["SubaddressType"].str.upper().str.title() 
df['SubaddressIdentifier'] = df['SubaddressIdentifier'].str.upper()
df['USPSBoxType'] = df['USPSBoxType'].str.upper()


df['USPS Box'] = df['USPSBoxType']+ (' ' + df['USPSBoxID']).fillna('')

df['Address 1'] = df['AddressNumber'] + (' ' + df['StreetNamePreDirectional']).fillna('')+ (' ' + df['StreetNamePreType']).fillna('')+ (' ' + df['StreetName']).fillna('')+ (' ' + df['StreetNamePostType']).fillna('')+ (' ' + df['StreetNamePostDirectional']).fillna('')

df['Address 2a'] = df['OccupancyType'] + (' ' + df['OccupancyIdentifier']).fillna('')+ (' ' + df['SubaddressType']).fillna('')+ (' ' + df['SubaddressIdentifier']).fillna('')
df['Address 2b'] = df['SubaddressType']+ (' ' + df['SubaddressIdentifier']).fillna('')
df['Address 2'] = df['Address 2a'].fillna('')+ ('' + df['Address 2b']).fillna('')


df=df.rename(columns={"address": "Old Address", "PlaceName": "City", "StateName": "State","ZipCode": "Zip Code"})
df = df[['SAP Number', 'GAG', 'name', 'Old Address', 'Address 1','Address 2','City','State','Zip Code','USPS Box']]

with pd.ExcelWriter(r'C:YOUR OUTPUT FILE LOCATION.xlsx') as writer:  # input your location for reading your output file
     df.to_excel(writer, sheet_name='new data',index=False)
    

os.startfile(r'C: YOUR OUTPUT FILE LOCATION') # input your location for reading your output file
