#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Mar 24 11:16:49 2025

@author: ashutoshgoenka
"""

import pandas as pd
import numpy as np
import streamlit as st

# from exif import Image as Image2


#import os
import zipfile
from zipfile import ZipFile, ZIP_DEFLATED
#import pathlib
#import shutil
#import docx
#import docxtpl
import random
from random import randint
from streamlit import session_state
import openpyxl
from openpyxl import load_workbook

st.set_page_config(layout="wide")




state = session_state
if "key" not in state:
    state["key"] = str(randint(1000, 100000000))

if "photo_saved" not in state:
    state["photo_saved"] = False

if "sample_file" not in state:
    state["sample_file"] = False
    
    
if "location_file" not in state:
    state["location_file"] = False

if "page_first_loaded" not in state:
    state["page_first_loaded"] = True
    
if "row_no" not in state:
    state["row_no"] = 1

    


def alp_pos(df_pos_temp, val ) -> str:
    """
    Returns the alphabet letter(s) corresponding to the given position in an extended sequence (1-based index).
    :param position: An integer representing the position in the alphabet sequence
    :return: The corresponding letter(s), or an empty string if input is invalid
    """
    
    position = df_pos_temp.columns.get_loc(val) + 2
    result = ""
    while position > 0:
        position -= 1
        result = chr(ord('A') + (position % 26)) + result
        position //= 26
    
    
    
    return result


#Function to generate Output Files
def df_excel_calculation(df_temp_excel):
    df_temp_excel["Temp"] = df_temp_excel.index+2
    df_temp_excel["Temp"] = df_temp_excel["Temp"].astype(str)
    supplier_serv_col  = alp_pos(df_temp_excel,"Supplier Service Fees")
    ta_serv_col = alp_pos(df_temp_excel,"TA Service Fees")
    total_serv_col = alp_pos(df_temp_excel,"Total Service Fees")
    gst_cat_col = alp_pos(df_temp_excel,"CGST/IGST")
    cgst_col = alp_pos(df_temp_excel,"CGST")
    sgst_col = alp_pos(df_temp_excel,"SGST")
    igst_col = alp_pos(df_temp_excel,"IGST")
    inv_val_col = alp_pos(df_temp_excel,"Invoice Value")
    air_charge_col = alp_pos(df_temp_excel,"Airline/Insuranance Charges")
    refund_charge_col = alp_pos(df_temp_excel,"Refund/Credit")
    airline_total_charge_col  = alp_pos(df_temp_excel,"Airline Charges Total Amount")

    df_temp_excel["Booking Date"] = pd.to_datetime(df_temp_excel["Booking Date"])
    df_temp_excel["Booking Date"] = df_temp_excel['Booking Date'].dt.strftime('%d/%m/%Y')
    df_temp_excel["Travel Date"] = pd.to_datetime(df_temp_excel["Travel Date"])
    df_temp_excel["Travel Date"] = df_temp_excel['Travel Date'].dt.strftime('%d/%m/%Y')
 
    
    df_temp_excel["Total Service Fees"] = "="+supplier_serv_col+df_temp_excel["Temp"]+"+"+ta_serv_col+df_temp_excel["Temp"]
    df_temp_excel["CGST"] = "="+total_serv_col+df_temp_excel["Temp"]+"*(2-"+gst_cat_col+df_temp_excel["Temp"]+")*0.09"
    df_temp_excel["SGST"] = "="+total_serv_col+df_temp_excel["Temp"]+"*(2-"+gst_cat_col+df_temp_excel["Temp"]+")*0.09"
    df_temp_excel["IGST"] = "="+total_serv_col+df_temp_excel["Temp"]+"*("+gst_cat_col+df_temp_excel["Temp"]+"-1)*0.18"
    df_temp_excel["Invoice Value"] = "="+total_serv_col+df_temp_excel["Temp"] + "+"+cgst_col+df_temp_excel["Temp"] + "+"+sgst_col+df_temp_excel["Temp"] + "+"+igst_col+df_temp_excel["Temp"]
    df_temp_excel["Airline Charges Total Amount"] =  "="+air_charge_col+df_temp_excel["Temp"] + "+"+refund_charge_col +df_temp_excel["Temp"] 
    df_temp_excel["Payable Amount"] = "="+airline_total_charge_col+df_temp_excel["Temp"] + "+"+inv_val_col +df_temp_excel["Temp"] 
    df_temp_excel.drop("Temp",axis=1 ,inplace=True)
    



#Function to generate Output File for Hotel

def df_excel_calculation_hotel(df_temp_excel):
    df_temp_excel["Temp"] = df_temp_excel.index+2
    df_temp_excel["Temp"] = df_temp_excel["Temp"].astype(str)
    supplier_serv_col  = alp_pos(df_temp_excel,"Supplier Service Fees")
    ta_serv_col = alp_pos(df_temp_excel,"TA Service Fees")
    total_serv_col = alp_pos(df_temp_excel,"Total Service Fees")
    gst_cat_col = alp_pos(df_temp_excel,"CGST/IGST")
    cgst_col = alp_pos(df_temp_excel,"CGST")
    sgst_col = alp_pos(df_temp_excel,"SGST")
    igst_col = alp_pos(df_temp_excel,"IGST")
    inv_val_col = alp_pos(df_temp_excel,"Invoice Value")
    hotel_charge_col = alp_pos(df_temp_excel,"Hotel Sales")
    refund_charge_col = alp_pos(df_temp_excel,"Refund/Credit")
    hotel_total_charge_col  = alp_pos(df_temp_excel,"Hotel Charges Total Amount")
    
    
    df_temp_excel["Total Service Fees"] = "="+supplier_serv_col+df_temp_excel["Temp"]+"+"+ta_serv_col+df_temp_excel["Temp"]
    df_temp_excel["CGST"] = "="+total_serv_col+df_temp_excel["Temp"]+"*(2-"+gst_cat_col+df_temp_excel["Temp"]+")*0.09"
    df_temp_excel["SGST"] = "="+total_serv_col+df_temp_excel["Temp"]+"*(2-"+gst_cat_col+df_temp_excel["Temp"]+")*0.09"
    df_temp_excel["IGST"] = "="+total_serv_col+df_temp_excel["Temp"]+"*("+gst_cat_col+df_temp_excel["Temp"]+"-1)*0.18"
    df_temp_excel["Invoice Value"] = "="+total_serv_col+df_temp_excel["Temp"] + "+"+cgst_col+df_temp_excel["Temp"] + "+"+sgst_col+df_temp_excel["Temp"] + "+"+igst_col+df_temp_excel["Temp"]
    df_temp_excel["Hotel Charges Total Amount"] =  "="+hotel_charge_col+df_temp_excel["Temp"] + "+"+refund_charge_col +df_temp_excel["Temp"] 
    df_temp_excel["Payable Amount"] = "="+hotel_total_charge_col+df_temp_excel["Temp"] + "+"+inv_val_col +df_temp_excel["Temp"] 
    df_temp_excel.drop("Temp",axis=1 ,inplace=True)






cntr=1    
st.title("Flight Accounts Reconciliation")
data_file = st.file_uploader("Upload Riya Ledger/Data File", type=['csv','xlsx'],accept_multiple_files=False,key="fileUploader_data")
booked_history_file = st.file_uploader("Upload Riya Airline Passenger Booked History File", type=['csv','xlsx'],accept_multiple_files=False,key="fileUploader_booked_history")
hotel_booking_file =  st.file_uploader("Upload Riya Hotel Booking File", type=['csv','xlsx'],accept_multiple_files=False,key="fileUploader_hotel_booked_history")
pas_master_file = st.file_uploader("Upload Riya Passenger Data Master File", type=['csv','xlsx'],accept_multiple_files=False,key="fileUploader_pass")
master_record_file = st.file_uploader("Upload Riya Master Record File", type=['csv','xlsx'],accept_multiple_files=False,key="fileUploader_master_history")
if pas_master_file is not None and data_file is not None and booked_history_file is not None and master_record_file is not None:
    df1 = pd.read_excel(data_file)
    df_passenger_master = pd.read_excel(pas_master_file)
    df_b1 = pd.read_excel(booked_history_file)
    df_existing = pd.read_excel(master_record_file)
    df_passenger_master.rename(columns = {'Name':'lead passenger'}, inplace = True)
    df_hotel_booking =  pd.read_excel(hotel_booking_file)
    st.write("Ledger Sheet Display")
    st.write(df1)
    st.write("Passenger Master List")
    st.write(df_passenger_master)
    df1["Diff"] = df1["DateTime"].diff() 
    df1["SamePNR"] = df1["AirlinePNR"] == df1["AirlinePNR"].shift()
    df1[["SamePNR"]]= df1[["SamePNR"]].shift(-1) 
    df1[["Diff"]]= df1[["Diff"]].shift(-1) 
    df1["DateTimeNew"] = df1["DateTime"] + df1["Diff"]
    df1["DateTimeNew"] = np.where(df1["SamePNR"] == True, df1["DateTimeNew"], df1["DateTime"])
    df1["Diff"] = df1["Diff"].astype('int64').astype(int)/1000000000
    df1["DateTimeNew"] = np.where(df1["Diff"] < 2, df1["DateTimeNew"], df1["DateTime"] )
    df1["DateTime"] = df1["DateTimeNew"] 
    df1.drop("Diff",axis=1 ,inplace=True)
    df1.drop("SamePNR",axis=1 ,inplace=True)
    df1.drop(0,axis=0,inplace=True)
    opening_balance = df1["Remaining"].values[0] - df1["CreditAmount"].values[0] + df1["DebitAmount"].values[0]
    df1.drop("AgentId", axis=1, inplace=True)
    df1.drop("Ref", axis=1, inplace=True)
    df1.drop("Agency Name", axis=1, inplace=True)
    df2 = df1.assign(Value = lambda x: x.CreditAmount - x.DebitAmount) 
    df2["TransactionType"] = df2["TransactionType"].fillna("Others")
    list_transaction = list(df2.TransactionType.unique())
    df2['Airline Sales'] = np.where(df2["TransactionType"] == 'Airline Sales', df2.Value, 0)
    for i in list_transaction:
        df2[i] = np.where(df2["TransactionType"] == i, df2.Value, 0)
    df2["RiyaPNR"] = df2["RiyaPNR"].fillna("No Input")
    df2["AirlinePNR"] = df2["AirlinePNR"].fillna("No Input")
    df2["Others Present"] = 0
    key = (df2["TransactionType"]=="Others") | (df2["TransactionType"]=="Seat Selection")
    df2.loc[key, "Others Present"] = df2.index[key]
    
    df2.groupby(["DateTime"]).sum(numeric_only=True)
    df3 = df2.groupby(["DateTime", "Description", "RiyaPNR", "AirlinePNR", "Others Present"]).sum(numeric_only=True)
    df3.drop("Remaining", axis = 1 , inplace = True)
    df3.reset_index(inplace=True) 
    df3.drop(["Others Present"], axis=1, inplace = True)
    #df3.to_excel("output.xlsx")
    df_b1 = df_b1[df_b1["Ticket Status"] != "TO TICKET"]
    df_b1["Passenger Name Split"] = df_b1["Passenger Name"].str.split(",")
    df_b1["No of PAX"] = df_b1["Passenger Name Split"].str.len()
    try:
        df_b1[["Airline Code", "Flight Number"]] = df_b1["Flight No"].str.split(" ", expand = True)
    except:
        split_string = df_b1["Flight No"].str.split(" ", expand = True)
        df_b1[["Airline Code", "Flight Number"]]  = split_string.iloc[:,:2]
    df_b1.drop("Passenger Name Split", axis = 1, inplace=True)
    df_b2 = df_b1[["Riya PNR", "Passenger Name", "Sector", "Departure Date", "Airline Code", "Airport Id", "No of PAX"]].copy()
    
    # Realigning Hotel Booking File
    no_of_rows =  df_hotel_booking[df_hotel_booking.columns[0]].count()
    col_idx_2 = ['S PNR', "Guest Name", "City Name","Room Type", "Check IN", "Check OUT"]
    df_hotel_booking = df_hotel_booking.reindex(columns=col_idx_2)
    df_hotel_booking["No of Nights"] = (df_hotel_booking["Check OUT"] -  df_hotel_booking["Check IN"]).dt.days +1
    df_hotel_booking.rename(columns={'S PNR': 'RiyaPNR'}, inplace=True)
    # Realignment done
    
    
    passenger_list = list(df_b2["Passenger Name"].str.split(","))
    passenger_list = list(set([i for name in passenger_list for i in name]))
    condition = True
    while condition and pas_master_file is not None:
        pass_master_list = list(df_passenger_master["lead passenger"])
        passenger_list_not_in_master = []
        for i in passenger_list:
            if i not in pass_master_list:
                passenger_list_not_in_master.append(i)
        df_pass = pd.DataFrame(passenger_list_not_in_master)
        try:
            no_of_missing_pass =  df_pass[df_pass.columns[0]].count()
        except:
            no_of_missing_pass = 0
            
        if no_of_missing_pass > 0:
            st.write("Missing Passneger List")
            st.write(df_pass)
            df_pass.to_excel("missing_passenger.xlsx")
            try:
                with open("missing_passenger.xlsx", "rb") as template_file:
                    template_byte = template_file.read()
                    btn_1 = st.download_button(
                            label="Download Missing Passenger List",
                            data=template_byte,
                            file_name="missing_passenger.xlsx",
                            mime='application/octet-stream'
                            )
            except:
                pass
            cntr_str = cntr
            cntr = cntr + 1
            
            pas_master_file = st.file_uploader("Upload the Updated Riya Passenger Data Master File", type=['csv','xlsx'],accept_multiple_files=False,key="fileUploader_pass_"+str(cntr_str))
            if pas_master_file is not None:
                df_passenger_master = pd.read_excel(pas_master_file)
                df_passenger_master.rename(columns = {'Name':'lead passenger'}, inplace = True)
        else:
            condition = False
        
    if no_of_missing_pass == 0:
        # st.write("No of Missing passenger")
        # st.write(no_of_missing_pass)
        
        
        df_b2.rename(columns = {'Riya PNR':'RiyaPNR'}, inplace = True)
        df_final = pd.merge(df3, df_b2, on='RiyaPNR', how ="left")
        
        #Merge Hotel File data in the Final Data
        
        df_final = pd.merge(df_final, df_hotel_booking, on='RiyaPNR', how ="left")
        df_final.loc[df_final["Check IN"].notna(), "Departure Date"] = df_final["Check IN"]
        df_final.loc[df_final["Guest Name"].notna(), "Passenger Name"] = df_final["Guest Name"]
        
        
        
        try:
            a= df_final["Airline Reschedule(FARE DIFFERENCE)"] 
        except:
            df_final["Airline Reschedule(FARE DIFFERENCE)"]   = 0
            
        try:
            a= df_final["Airline Other Services"] 
        except:
            df_final["Airline Other Services"]   = 0
        
        try:
            a= df_final["Airline Reschedule(SUPPILER PENALTY)"] 
        except:
            df_final["Airline Reschedule(SUPPILER PENALTY)"]   = 0
            
        try:
            df_final["Airline Reschedule(SUPPILER PENALTY)"] = df_final["Airline Reschedule(SUPPILER PENALTY)"] + df_final["Airline Reschedule(FARE DIFFERENCE)"] + df_final["Airline Other Services"]
        except:
            df_final["Airline Reschedule(SUPPILER PENALTY)"] = df_final["Airline Reschedule(SUPPILER PENALTY)"] + df_final["Airline Reschedule(FARE DIFFERENCE)"]
            
            
            
            
        
        df_final["Booking Date"] = pd.to_datetime(df_final['DateTime']).dt.date
        df_final["Booking Time"] = pd.to_datetime(df_final['DateTime']).dt.time
        df_final["Travel Date"] = pd.to_datetime(df_final['Departure Date']).dt.date
        
        st.write("Check")
        st.write(df_final)
        try:
            a = df_final["Airline Cancellation(SOLD AMOUNT REVERSAL)"]
        except:
            df_final["Airline Cancellation(SOLD AMOUNT REVERSAL)"] = 0
            
        try: 
            a = df_final["PG Online Transfer"]
        except:
            df_final["PG Online Transfer"] = 0
        try: 
            a = df_final["PG Online Transfer Incentive"]
        except:
            df_final["PG Online Transfer Incentive"] = 0
            
        try:
            a = df_final["Hotel Sales"]
        except:
            df_final["Hotel Sales"] = 0
        
        
        try:
            a = df_final["Seat Selection"]
        except:
            df_final["Seat Selection"] = 0
        
        try:
            a = df_final["Others"]
        except:
            df_final["Others"] = 0
            
        try:
            a = df_final["Infant Charge"]
        except:
            df_final["Infant Charge"] = 0
            
            
        
        df_final.loc[df_final["Airline Cancellation(SOLD AMOUNT REVERSAL)"]>0, "Product Type"] = "Ticket Cancellation"
        df_final.loc[df_final["PG Online Transfer"]>0, "Product Type"] = "Deposit"
        df_final.loc[df_final["Airline Reschedule(SUPPILER PENALTY)"]!=0, "Product Type"] = "Ticket Rescheduled"
        # df_final.loc[df_final["Insurance Sales"]<0, "Product Type"] = "Insurance"
        df_final.loc[df_final["Seat Selection"]<0, "Product Type"] = "Seat Selection"
        # df_final.loc[df_final["Airline Cancellation(Seat Selection)"]>0, "Product Type"] = "Seat Selection Refund"
        df_final.loc[df_final["PG Online Transfer Incentive"]>0, "Product Type"] = "Deposit Incentive"
        df_final.loc[df_final["PG Online Transfer Incentive"]>0, "Passenger Name"] = "Deposit Incentive"
        # df_final.loc[df_final["Airline Baggage"]<0, "Product Type"] = "Airline Baggage"
        df_final.loc[df_final["Airline Sales"]<0, "Product Type"] = "Ticket Issued"
        #df_final.loc[(df_final["Airline Sales"]<0) & (df_final["Seat Selection"]<0), "Product Type"] = "Ticket & Seat Issued"
        df_final.loc[df_final["Others"]!=0, "Product Type"] = "REQUIRE MANUAL VERIFICATION/OTHERS"
        df_final.loc[df_final["Offline Adjustment"]!=0, "Product Type"] = "REQUIRE MANUAL VERIFICATION/OFFLINE ADJUSTMENT"
        df_final.loc[df_final["Hotel Sales"]<0, "Product Type"] = "Hotel Sales"
        df_final.loc[df_final["Infant Charge"]<0, "Product Type"] = "Ticket Issued (Infant Charges)"
        
        
        
        df_final.drop("DateTime", axis=1, inplace=True)
        df_final.drop("Departure Date", axis=1, inplace=True)
        
        col_list = list(df_final.columns)
        df_final.drop(['CreditAmount','DebitAmount','Value'], axis=1, inplace=True)
        
        if no_of_rows >0:
            col_list = ['Booking Date', "Booking Time",'Product Type', 'Airport Id', 'Description','RiyaPNR', 'AirlinePNR','Passenger Name',
                    "No of PAX",'Airline Code','Sector', 'Travel Date', "Check OUT","City Name", "Room Type", 'No of Nights',
                    'Airline Sales','Others','Airline Commission','Airline TDS On Earnings','Service Fee','GST on Service Fee', 'Infant Charge',
                    'Airline Cancellation(SOLD AMOUNT REVERSAL)','Airline Cancellation(PENALTY)','Agent Cancellation(SERVICE FEE)',
                    'Agent Cancellation(GST ON SERVICE FEE)','Airline Earnings Reversal','PG Online Transfer',
                    'Airline TDS Amount Reversal','Offline Adjustment','Airline Reschedule(SUPPILER PENALTY)','Airline Earnings','Hotel Sales',
                    'Hotel Commission','Hotel TDS on earnings','Insurance Sales','Insurance Commission','Insurance TDS On Earnings','Seat Selection','Airline Cancellation(Seat Selection)',
                    'PG Online Transfer Incentive','PG Online Transfer Incentive TDS','Airline Baggage']
            df_final = df_final.reindex(columns=col_list)
            data_colimn_position = df_final.columns.get_loc('Airline Sales')
            df_final["total amount charged"] = df_final.iloc[:,data_colimn_position:].sum(axis = 1)
        else:
            col_list = ['Booking Date', "Booking Time",'Product Type', 'Airport Id', 'Description','RiyaPNR', 'AirlinePNR','Passenger Name',
                    "No of PAX",'Airline Code','Sector','Travel Date',
                    'Airline Sales','Others','Airline Commission','Airline TDS On Earnings','Service Fee','GST on Service Fee', 'Infant Charge',
                    'Airline Cancellation(SOLD AMOUNT REVERSAL)','Airline Cancellation(PENALTY)','Agent Cancellation(SERVICE FEE)',
                    'Agent Cancellation(GST ON SERVICE FEE)','Airline Earnings Reversal','PG Online Transfer',
                    'Airline TDS Amount Reversal','Offline Adjustment','Airline Reschedule(SUPPILER PENALTY)','Airline Earnings','Hotel Sales',
                    'Hotel Commission','Hotel TDS on earnings','Insurance Sales','Insurance Commission','Insurance TDS On Earnings','Seat Selection','Airline Cancellation(Seat Selection)',
                    'PG Online Transfer Incentive','PG Online Transfer Incentive TDS','Airline Baggage']
            df_final = df_final.reindex(columns=col_list)
            data_colimn_position = df_final.columns.get_loc('Airline Sales')
            df_final["total amount charged"] = df_final.iloc[:,data_colimn_position:].sum(axis = 1)
                
        
        df_final[df_final["Product Type"] == "Ticket Cancellation"]
        df_final["Closing balance"]  = df_final["total amount charged"].cumsum() + opening_balance 
        df_master_data = df_final.copy()
        df_master_data  = df_master_data[(df_master_data["Product Type"] == "Ticket Issued") | (df_master_data["Product Type"] == "Ticket Issued (Infant Charges)")]
        
        col_index= ["Booking Date", "Airport Id", "Description", "RiyaPNR", "AirlinePNR", "Passenger Name", "No of PAX", "Airline Code", "Sector", "Travel Date"]
        df_master_data = df_master_data.reindex(columns=col_index)
        
        st.write("Existing DF")
        st.write(df_existing)
        
        
        st.write("Master Data")
        st.write(df_master_data)
        
        index_to_delete = []
        for i, row in  df_master_data.iterrows():
            if df_existing['RiyaPNR'].eq(row[3]).any():
                index_to_delete.append(i)
        #st.write(index_to_delete)
        df_master_data.drop(index_to_delete, inplace=True)
        
        existing_list = list(df_existing.columns)
        existing_list.remove('Travel Date')
        existing_list.append('Travel Date')
        df_existing = df_existing.reindex(columns=existing_list)
        try:
            df_existing["Booking Date"] = df_existing["Booking Date"].dt.date
        except:
            pass
        
        try:
            df_existing["Travel Date"] = df_existing["Travel Date"].dt.date
        except:
            pass
        
        
        st.write("Riya Master Data")
        st.write(df_existing)
        st.write(df_master_data)
        
        merged_df = pd.concat([df_existing.copy(), df_master_data.copy()],ignore_index=True)
        merged_df.to_excel("Riya_Master_Record.xlsx")
        
        #filename = "Riya_Master_Record.xlsx"
        # workbook = load_workbook(master_record_file)
        # worksheet = workbook.active
        # for i, row in df_master_data.iterrows():
        #     worksheet.append(list(row))
        # workbook.save(master_record_file)
        
        
        # df_master_data_2 = pd.read_excel(master_record_file)
        # df_master_data_2.to_excel("Riya_Master_Record.xlsx")
        # st.write(df_master_data_2)
        
        
        
        df_final["Passenger Name"] = df_final["Passenger Name"].fillna("Passenger Name Missing")
        for i, row in df_final.iterrows():
            if row[5] != "No Input" and row[7] == "Passenger Name Missing":
                val = row[5]
                #print(row[5])
                rows = df_existing.index[df_existing["RiyaPNR"]==val]
                if rows.size !=0:
                    
                    pos_t_date = df_existing.columns.get_loc("Travel Date")
                    row_data = list(df_existing.iloc[rows[0],:])
                    #st.write(row_data)
                    #st.write(pos_t_date)
                    row_data[pos_t_date] = pd.to_datetime(row_data[pos_t_date]).date()
                    #st.write(row_data)
                
                    df_final.iloc[i,3:12] = row_data[1:]
        #             index_to_delete.append(i)
        
        
        df_final["Product Type"] = df_final["Product Type"].fillna("REQUIRE MANUAL VERIFICATION/RECONCILIATION")
        df_final.loc[df_final["Product Type"] == "nan", "Product Type"] = "REQUIRE MANUAL VERIFICATION/RECONCILIATION"
        df_final.to_excel("Supplier_Master.xlsx")
        
        key_a = ((df_final["Product Type"]=="REQUIRE MANUAL VERIFICATION/RECONCILIATION") | 
                 (df_final["Product Type"] == "REQUIRE MANUAL VERIFICATION/OTHERS") | 
                 (df_final["Product Type"] == "REQUIRE MANUAL VERIFICATION/OFFLINE ADJUSTMENT"))
        df_manual_entry =  df_final[key_a]
        
        df_manual_entry.to_excel("Require_Manual_Verification_Data.xlsx")
        
                    
     
        df_temp_2 = df_final.copy()
        df_temp_2["Base Amount"] = df_temp_2.fillna(0)['Airline Sales'] + df_temp_2.fillna(0)['Service Fee']+df_temp_2.fillna(0)['GST on Service Fee'] + \
        df_temp_2.fillna(0)['Airline Cancellation(SOLD AMOUNT REVERSAL)']  + df_temp_2.fillna(0)['Insurance Sales']
        
        df_temp_2['Airline Commission'] = df_temp_2.fillna(0)['Airline Commission'] + df_temp_2.fillna(0)['Airline Earnings Reversal']
        df_temp_2['Airline TDS On Earnings'] = df_temp_2.fillna(0)['Airline TDS On Earnings'] + df_temp_2.fillna(0)['Airline TDS Amount Reversal']
    
        df_temp_2["Debit Amount"] = df_temp_2.fillna(0)["PG Online Transfer"] 
    
        df_temp_2["Total Amount"] = df_temp_2.fillna(0)['Base Amount'] + df_temp_2.fillna(0)['Airline Cancellation(PENALTY)']  + \
        df_temp_2.fillna(0)['Airline Commission'] + df_temp_2.fillna(0)['Airline TDS On Earnings'] + df_temp_2.fillna(0)["Debit Amount"]
        df_temp_2["Credit Amount"] = np.where(df_temp_2["Total Amount"]<0,df_temp_2["Total Amount"],0)
        df_temp_2["Debit Amount"] = np.where(df_temp_2["Total Amount"]>0,df_temp_2["Total Amount"],0)
    
    
        col_list = ['Supplier Code','Booking Date', 'Airline Code', 'Sector','Travel Date', 'AirlinePNR','Passenger Name', 'Base Amount', 'Airline Cancellation(PENALTY)','Airline Commission','Airline TDS On Earnings', 'Debit Amount','Credit Amount', "Closing balance", 'RiyaPNR',   'Others','Airline Cancellation(SOLD AMOUNT REVERSAL)','Airline Cancellatcion(PENALTY)','Airline Earnings Reversal','PG Online Transfer','Airline TDS Amount Reversal','Offline Adjustment','Airline Reschedule(SUPPILER PENALTY)','Airline Earnings','Insurance Sales','Insurance Commission','Insurance TDS On Earnings','Seat Selection','Airline Cancellation(Seat Selection)','PG Online Transfer Incentive','PG Online Transfer Incentive TDS','Airline Baggage',"Booking Time",]
        df_temp_2 = df_temp_2.reindex(columns=col_list)
        df_temp_2.to_excel("output_temp_2.xlsx")
        df_temp = df_final.copy()
        df_temp["Supplier Code"] = "RC"
        df_temp["DT Service Fees"] = 200*df_temp["No of PAX"]
        df_temp["Total Service Fees"] = df_temp["DT Service Fees"] - df_temp["Service Fee"]
        
        df_temp["GST Amt"] = df_temp["Total Service Fees"] * 0.18
        df_temp["Net Amt"] = (df_temp["Airline Sales"] + df_temp["Service Fee"] + df_temp["GST on Service Fee"] +  df_temp["Airline Commission"] + df_temp["Airline TDS On Earnings"])*(-1) 
        df_temp["Round Off"] = 0
        df_temp["Invoice Amt"] = - df_temp["Airline Sales"] + df_temp["Total Service Fees"] + df_temp["GST Amt"]
        
        col_list = ['Supplier Code','Booking Date', 'Airline Code', 'Sector','Travel Date', 'AirlinePNR',"No of PAX",'Product Type','Passenger Name', 'Airline Sales','Service Fee','GST on Service Fee','Airline Commission','Airline TDS On Earnings','Total Service Fees','GST Amt', 'Net Amt', 'Round Off', 'Invoice Amt', 'Airport Id', 'Description','RiyaPNR',   'Others','Airline Cancellation(SOLD AMOUNT REVERSAL)','Airline Cancellatcion(PENALTY)','Airline Earnings Reversal','PG Online Transfer','Airline TDS Amount Reversal','Offline Adjustment','Airline Reschedule(SUPPILER PENALTY)','Airline Earnings','Insurance Sales','Insurance Commission','Insurance TDS On Earnings','Seat Selection','Airline Cancellation(Seat Selection)','PG Online Transfer Incentive','PG Online Transfer Incentive TDS','Airline Baggage',"Booking Time",]
        df_temp = df_temp.reindex(columns=col_list)
        df_temp['Airline Sales'] = df_temp['Airline Sales']*(-1)
        df_temp['Service Fee'] = df_temp['Service Fee'] *(-1)
        df_temp['GST on Service Fee'] = df_temp['GST on Service Fee'] * (-1)
        df_temp['Airline TDS On Earnings'] = df_temp['Airline TDS On Earnings']  * (-1)
        df_customer =df_final.copy()
        df_customer = df_customer[df_customer["RiyaPNR"] != "No Input"]
        df_customer["Airline/Insuranance Charges"] = df_customer.fillna(0)['Airline Sales'] + df_customer.fillna(0)['Airline Cancellation(PENALTY)'] + df_customer.fillna(0)['Airline Reschedule(SUPPILER PENALTY)'] + df_customer.fillna(0)['Insurance Sales'] + df_customer.fillna(0)['Seat Selection'] + df_customer.fillna(0)['Airline Baggage']
        # df_customer["Airline/Insuranance Charges"] = -1*df_customer["Airline/Insuranance Charges"]
        df_customer["Refund/Credit"] = df_customer.fillna(0)['Airline Cancellation(SOLD AMOUNT REVERSAL)'] + df_customer.fillna(0)["Airline Cancellation(Seat Selection)"] 
        # df_customer["Refund/Credit"] = -1*df_customer["Refund/Credit"]
           
        
        df_customer_1 = df_customer[["Booking Date", "Booking Time", "Product Type", "Airport Id", "Description", "RiyaPNR", "AirlinePNR", "Passenger Name", "No of PAX", "Travel Date",'Airline Code', 'Sector', "Airline/Insuranance Charges", "Hotel Sales", "Refund/Credit" ,'Service Fee', 'GST on Service Fee']].copy()
        
        df_customer_1['Airline/Insuranance Charges'] = df_customer_1['Airline/Insuranance Charges'] *(-1)
        df_customer_1['Refund/Credit'] = df_customer_1['Refund/Credit'] *(-1)
        df_customer_1['Service Fee'] = df_customer_1['Service Fee'] *(-1)
        df_customer_1['GST on Service Fee'] = df_customer_1['GST on Service Fee'] *(-1)
        
        
        df_customer_1.rename(columns = {'Service Fee':'Supplier Service Fees'}, inplace = True)
        df_customer_1.rename(columns = {'GST on Service Fee':'GST on Supplier Service Fees'}, inplace = True)
        df_customer_1["DT Service Fees"] = 200*df_customer_1["No of PAX"]
        df_customer_1["Total Service Fees"] = df_customer_1["DT Service Fees"] + df_customer_1["Supplier Service Fees"]
        df_customer_1["CGST/IGST"] = 0
        df_customer_1["CGST"] = 0
        df_customer_1["SGST"] = 0
        df_customer_1["IGST"] = 0
        df_customer_1["Invoice Value"] = df_customer_1["CGST"] + df_customer_1["SGST"] + df_customer_1["IGST"] + df_customer_1["Total Service Fees"]
        df_customer_1["Payable Amount"] =  df_customer_1['Airline/Insuranance Charges'] + df_customer_1["Invoice Value"] + df_customer_1["Refund/Credit"]
        # df_customer_1["B2B/B2C"] = "B2C"
        df_customer_1.to_excel("output_2.xlsx")
        df_customer_dom = df_customer_1[(df_customer_1["Airport Id"] == "Domestic") & ((df_customer_1["Product Type"] =="Ticket Issued") | (df_customer_1["Product Type"] =="Ticket Issued (Infant Charges)"))]

        df_remaining  = pd.concat([df_customer_1,df_customer_dom]).drop_duplicates(keep=False)
        df_customer_dom["lead passenger"] = df_customer_dom["Passenger Name"].str.split(",")
        df_customer_dom["lead passenger"] = df_customer_dom["lead passenger"].apply(lambda x: x[0])
        df_passenger_master["B2B/B2C"] = "B2C"
        df_passenger_master["B2B/B2C"] = df_passenger_master["B2B/B2C"].where(df_passenger_master["GST Number"].isna(),"B2B")
        df_passenger_master.rename(columns = {'Name':'lead passenger'}, inplace = True)
        df_dom_final = pd.merge(df_customer_dom, df_passenger_master, on='lead passenger', how ="left")
        df_dom_final["CGST/IGST"] = df_dom_final["CGST/IGST"].where(df_dom_final["State"] != "Maharashtra", 1)
        df_dom_final["CGST/IGST"] = df_dom_final["CGST/IGST"].where(df_dom_final["State"] == "Maharashtra", 2)
        df_dom_final["CGST"] = df_dom_final["CGST"].where(df_dom_final["State"] != "Maharashtra", df_dom_final["Total Service Fees"]* 0.09)
        df_dom_final["SGST"] = df_dom_final["SGST"].where(df_dom_final["State"] != "Maharashtra", df_dom_final["Total Service Fees"]* 0.09)
        df_dom_final["IGST"] = df_dom_final["IGST"].where(df_dom_final["State"] == "Maharashtra", df_dom_final["Total Service Fees"]* 0.18)
        df_dom_final["Invoice Value"] = df_dom_final["CGST"] + df_dom_final["SGST"] + df_dom_final["IGST"] + df_dom_final["Total Service Fees"]
        df_dom_final["Payable Amount"] = df_dom_final["Invoice Value"] + df_dom_final["Airline/Insuranance Charges"] + df_dom_final["Refund/Credit"]
        
        
        col_list_2 = ['Booking Date', 'Invoice to', 'City', 'State', 'GST Number', 'Airline Code', 'Sector', 'Travel Date', 'AirlinePNR', 'Passenger Name', 'No of PAX', "Airline/Insuranance Charges", "Refund/Credit" , 'Supplier Service Fees', 'GST on Service Fee', 'TA Service Fees','Total Service Fees',"CGST/IGST",'CGST', 'SGST', 'IGST', 'Invoice Value',"Airline Charges Total Amount",'Payable Amount', "B2B/B2C"]

        

        
        
        df_dom_final = df_dom_final.reindex(columns=col_list_2)
        
        
        
        
        df_dom_final["Temp"] = df_dom_final.index+2
        df_dom_final["Temp"] = df_dom_final["Temp"].astype(str)
        df_excel_calculation(df_dom_final)
        df_dom_final.to_excel("domestic_final.xlsx")
        
        
        
        df_customer_intl = df_customer_1[(df_customer_1["Airport Id"] == "International") & ((df_customer_1["Product Type"] =="Ticket Issued")|(df_customer_1["Product Type"] =="Ticket Issued (Infant Charges)")) ]
        df_remaining = pd.concat([df_remaining,df_customer_intl]).drop_duplicates(keep=False)
        df_customer_intl["lead passenger"] = df_customer_intl["Passenger Name"].str.split(",")
        df_customer_intl["lead passenger"] = df_customer_intl["lead passenger"].apply(lambda x: x[0])
        df_intl_final = pd.merge(df_customer_intl, df_passenger_master, on='lead passenger', how ="left")
        df_intl_final = df_intl_final.reindex(columns=col_list_2)
        
        df_intl_final["CGST/IGST"] = df_intl_final["CGST/IGST"].where(df_intl_final["State"] != "Maharashtra", 1)
        df_intl_final["CGST/IGST"] = df_intl_final["CGST/IGST"].where(df_intl_final["State"] == "Maharashtra", 2)
        
        df_excel_calculation(df_intl_final)
        df_intl_final.to_excel("international_final.xlsx")
        
        
        
        
        df_all_flight = df_intl_final.copy(deep=True)
        df_all_flight["Domestic/International"] = "International"
        df_temp_dom = df_dom_final.copy(deep=True)
        df_temp_dom["Domestic/International"] = "Domestic"
        # df_all_flight = df_all_flight.append(df_temp_dom)
        df_all_flight = pd.concat([df_all_flight, df_temp_dom], ignore_index=True)
        df_all_flight = df_all_flight.sort_values(by=['Booking Date'], ascending=True)
        df_all_flight = df_all_flight.reset_index(drop=True)
        
        
        df_all_flight['Booking Date'] = pd.to_datetime(df_all_flight['Booking Date'], format='%d/%m/%Y')
        df_all_flight['Booking Date'] = df_all_flight['Booking Date'].dt.strftime('%Y/%m/%d')
        df_all_flight['Travel Date'] = pd.to_datetime(df_all_flight['Travel Date'], format='%d/%m/%Y')
        df_all_flight['Travel Date'] = df_all_flight['Travel Date'].dt.strftime('%Y/%m/%d')
        
        
        df_excel_calculation(df_all_flight)
        df_all_flight.to_excel("All_Tickets_final.xlsx")
        
        passenger_list = list(df_customer_intl["Passenger Name"].str.split(","))
        passenger_list = list(set([i for name in passenger_list for i in name]))
        df_pass = pd.DataFrame(passenger_list)
        
        df_customer_ticket_cancellation = df_customer_1[(df_customer_1["Product Type"] =="Ticket Cancellation")]
        df_remaining = pd.concat([df_remaining,df_customer_ticket_cancellation]).drop_duplicates(keep=False)
        df_customer_ticket_cancellation["lead passenger"] = df_customer_ticket_cancellation["Passenger Name"].str.split(",")
        df_customer_ticket_cancellation["lead passenger"] = df_customer_ticket_cancellation["lead passenger"].apply(lambda x: x[0])
        df_cancel_final = pd.merge(df_customer_ticket_cancellation, df_passenger_master, on='lead passenger', how ="left")
        
        #Adding Missing Names in Passenger List
        missing_name_list = list(df_cancel_final[df_cancel_final["Invoice to"].isnull()]["lead passenger"])
        passenger_list_not_in_master = passenger_list_not_in_master + missing_name_list
        df_pass = pd.DataFrame(passenger_list_not_in_master)
        df_pass.to_excel("missing_passenger.xlsx")
        #Addition Completed
        df_cancel_final = df_cancel_final.reindex(columns=col_list_2)
        
        df_cancel_final["CGST/IGST"] = df_cancel_final["CGST/IGST"].where(df_cancel_final["State"] != "Maharashtra", 1)
        df_cancel_final["CGST/IGST"] = df_cancel_final["CGST/IGST"].where(df_cancel_final["State"] == "Maharashtra", 2)
        
        df_excel_calculation(df_cancel_final)
        df_cancel_final.to_excel("cancellation_final.xlsx")
        
        df_customer_ticket_reschedule = df_customer_1[(df_customer_1["Product Type"] =="Ticket Rescheduled")]
        df_remaining = pd.concat([df_remaining,df_customer_ticket_reschedule]).drop_duplicates(keep=False)
        df_customer_ticket_reschedule["lead passenger"] = df_customer_ticket_reschedule["Passenger Name"].str.split(",")
        df_customer_ticket_reschedule["lead passenger"] = df_customer_ticket_reschedule["lead passenger"].apply(lambda x: x[0])
        df_reschedule_final = pd.merge(df_customer_ticket_reschedule, df_passenger_master, on='lead passenger', how ="left")
        #Adding Missing Names in Passenger List
        missing_name_list = list(df_reschedule_final[df_reschedule_final["Invoice to"].isnull()]["lead passenger"])
        passenger_list_not_in_master = passenger_list_not_in_master + missing_name_list
        df_pass = pd.DataFrame(passenger_list_not_in_master)
        df_pass.to_excel("missing_passenger.xlsx")
        #Addition Completed
        
        
        df_reschedule_final = df_reschedule_final.reindex(columns=col_list_2)
        
        df_reschedule_final["CGST/IGST"] = df_reschedule_final["CGST/IGST"].where(df_reschedule_final["State"] != "Maharashtra", 1)
        df_reschedule_final["CGST/IGST"] = df_reschedule_final["CGST/IGST"].where(df_reschedule_final["State"] == "Maharashtra", 2)
        
        df_excel_calculation(df_reschedule_final)
        df_reschedule_final.to_excel("rescheduling_final.xlsx")
            
        
        
        
        df_customer_insurance= df_customer_1[(df_customer_1["Product Type"] =="Insurance")]
        df_remaining = pd.concat([df_remaining,df_customer_insurance]).drop_duplicates(keep=False)
        df_customer_insurance["lead passenger"] = df_customer_insurance["Passenger Name"].str.split(",")
        df_customer_insurance["lead passenger"] = df_customer_insurance["lead passenger"].apply(lambda x: x[0])
        df_insurance_final = pd.merge(df_customer_insurance, df_passenger_master, on='lead passenger', how ="left")
        df_insurance_final = df_insurance_final.reindex(columns=col_list_2)
        
        df_insurance_final["CGST/IGST"] = df_insurance_final["CGST/IGST"].where(df_insurance_final["State"] != "Maharashtra", 1)
        df_insurance_final["CGST/IGST"] = df_insurance_final["CGST/IGST"].where(df_insurance_final["State"] == "Maharashtra", 2)
        
        df_excel_calculation(df_insurance_final)
        df_insurance_final.to_excel("insurance_final.xlsx")
        
        
        
        df_customer_seat_selection= df_customer_1[(df_customer_1["Product Type"] =="Seat Selection") | (df_customer_1["Product Type"] =="Seat Selection Refund")]
        df_remaining = pd.concat([df_remaining,df_customer_seat_selection]).drop_duplicates(keep=False)
        df_customer_seat_selection["lead passenger"] = df_customer_seat_selection["Passenger Name"].str.split(",")
        df_customer_seat_selection["lead passenger"] = df_customer_seat_selection["lead passenger"].apply(lambda x: x[0])
        df_seat_selection_final = pd.merge(df_customer_seat_selection, df_passenger_master, on='lead passenger', how ="left")
        df_seat_selection_final = df_seat_selection_final.reindex(columns=col_list_2)
        
        df_seat_selection_final["CGST/IGST"] = df_seat_selection_final["CGST/IGST"].where(df_seat_selection_final["State"] != "Maharashtra", 1)
        df_seat_selection_final["CGST/IGST"] = df_seat_selection_final["CGST/IGST"].where(df_seat_selection_final["State"] == "Maharashtra", 2)
        df_seat_selection_final["TA Service Fees"] = 0
        df_excel_calculation(df_seat_selection_final)
        df_seat_selection_final.to_excel("seat_final.xlsx")
        
        
        df_customer_hotel_data = df_customer[["Booking Date", "Booking Time", "Product Type", 'Description', 'City Name',"RiyaPNR", 'Room Type', 'No of Nights',"AirlinePNR", "Passenger Name", "Travel Date",'Check OUT', "Hotel Sales", "Refund/Credit" ,'Service Fee', 'GST on Service Fee']].copy()
        
        df_customer_hotel_sales = df_customer_hotel_data[(df_customer_hotel_data["Product Type"] =="Hotel Sales")]
        df_remaining = pd.concat([df_remaining,df_customer_hotel_sales]).drop_duplicates(keep=False)
        df_customer_hotel_sales["lead passenger"] = df_customer_hotel_sales["Passenger Name"].str.split(",")
        df_customer_hotel_sales["lead passenger"] = df_customer_hotel_sales["lead passenger"].apply(lambda x: x[0])
        df_hotel_final = pd.merge(df_customer_hotel_sales, df_passenger_master, on='lead passenger', how ="left")
        col_hotel_list = ["Booking Date", "Invoice to", "City", "State", "GST Number", "AirlinePNR","Passenger Name", "Description", 
                          "Room Type", "City Name", "Travel Date","Check OUT", "No of Nights", "Hotel Sales", "Refund/Credit", "Supplier Service Fees",
                         "GST on Service Fee", "TA Service Fees", "Total Service Fees", "CGST/IGST", "CGST", "SGST", "IGST", "Invoice Value",
                         "Airline Charges Total Amount", "Payable Amount", "B2B/B2C"]
        df_hotel_final = df_hotel_final.reindex(columns=col_hotel_list)
        
        df_hotel_final.rename(columns={"Check OUT": "Check Out", "Airline Charges Total Amount": "Hotel Charges Total Amount",  }, inplace=True)
        df_hotel_final["Hotel Sales"] = - df_hotel_final["Hotel Sales"]
        
        df_hotel_final["CGST/IGST"] = df_hotel_final["CGST/IGST"].where(df_hotel_final["State"] != "Maharashtra", 1)
        df_hotel_final["CGST/IGST"] = df_hotel_final["CGST/IGST"].where(df_hotel_final["State"] == "Maharashtra", 2)
        
        df_excel_calculation_hotel(df_hotel_final)
        df_hotel_final.to_excel("hotel_final.xlsx")
        
        
        
        
        final_output_file_list = ["Supplier_Master.xlsx","output_2.xlsx", "All_Tickets_final.xlsx", "domestic_final.xlsx", "international_final.xlsx", "All_Tickets_final.xlsx", "cancellation_final.xlsx", 
                                  "rescheduling_final.xlsx", "insurance_final.xlsx",  "seat_final.xlsx", "hotel_final.xlsx", "Riya_Master_Record.xlsx"]
        
        
        zip_path = "final_output.zip"
        
        with ZipFile(zip_path, 'w', ZIP_DEFLATED) as zip:
            for file in final_output_file_list:
                zip.write(file, arcname=file)
                
                
        with open("final_output.zip", "rb") as fp:
            btn = st.download_button(
                label="Download ZIP",
                data=fp,
                file_name="final_output.zip",
                mime="application/zip"
            )
        
        # try:
        #     with open("output_2.xlsx", "rb") as template_file:
        #         template_byte = template_file.read()
        #         btn_1 = st.download_button(
        #                 label="Download Output File",
        #                 data=template_byte,
        #                 file_name="output_2.xlsx",
        #                 mime='application/octet-stream'
        #                 )
        # except:
        #     pass
    