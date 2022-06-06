import os
import re
import glob
import numpy as np
import math
import matplotlib.colors as cl
import matplotlib.pyplot as plt
import peakutils
from natsort import natsorted
import pandas as pd
import lmfit.models
from openpyxl import load_workbook
import shutil
from shutil import copyfile
import xlrd
import xrdtools
from collections import Counter
from matplotlib.colors import rgb2hex

def Read_in_DataFrame(file_path_c): #Read in DataFrame, DF1, containing extracted peaks
    loc = file_path_c + "/Unique_Fitted_Picked_Peaks.xlsx" #Locate desired excel file
    DF1 = pd.read_excel(loc) #Read in file as a dataframe, DF1
    return DF1

def Similarities(DF1, EP_len_label, Position_tol, Width_tol, writer): #Identify similarities between PXRD patterns
    filenames = DF1['file_name'].unique() #Determine unique filenames in DF1
    mdf = pd.DataFrame(columns=[i[:EP_len_label] for i in filenames]) #Set all filenames as columns
    mdf['files']=[i[:EP_len_label] for i in filenames] #Add a column of filenames
    mdf = mdf.set_index('files') #Set column of filenames as index (rows)
    mwdf = pd.DataFrame(columns=[i[:EP_len_label] for i in filenames]) #Set all filenames as columns
    mwdf['files']=[i[:EP_len_label] for i in filenames] #Add a column of filenames
    mwdf = mwdf.set_index('files')  #Set column of filenames as index (rows)

    d = dict(tuple(DF1.groupby('file_name'))) #Group data by filenames in DF1
    p_duplicates_l = [] #New list to collect two theta values in file i
    p_duplicates_r = [] #New list to collect two theta values in file e
    w_duplicates_l = [] #New list to collect fwhm values in file i, two theta j
    w_duplicates_r = [] #New list to collect fwhm values in file e, two theta y
    pp1 = [] #New list to collect position matched filenames (i)
    pp2 = [] #New list to collect position matched filenames (e)
    pw1 = [] #New list to collect fwhm matched filenames (i)
    pw2 = [] #New list to collect fwhm matched filenames (e)
    for line, i in enumerate(filenames[:-1]): #For each filename, except the last
        for e in filenames[(line+1):]: #For each succeeding filename
            a = 0 #Number of similar widths
            b = 0 #Number of similar positions
            for j in d[i]['two_theta']: #For each two theta value in file i
                for y in d[e]['two_theta']: #For each two theta value in file e
                    x = math.isclose(j, y, abs_tol=(Position_tol)) #If two theta values are within a set tolerence
                    if x == True:
                        p_duplicates_l.append(j) #Collect two theta position in file i
                        p_duplicates_r.append(y) #Collect two theta position in file e
                        b = b+1 #Add 1 to the number of similar positions
                        k= DF1.loc[(DF1['two_theta'] == j) & (DF1['file_name'] == i), 'fwhm'].iloc[0] #Identify fwhm value corresponding to file i, two theta, j
                        l= DF1.loc[(DF1['two_theta'] == y) & (DF1['file_name'] == e), 'fwhm'].iloc[0] # Identify fwhm value corresponding to file e, two theta, y
                        x = math.isclose(k, l, abs_tol=(Width_tol)) #If fwhm values are within a set tolerence
                        if x == True:
                            w_duplicates_l.append(k) #Collect fwhm value in file i, two theta j
                            w_duplicates_r.append(l) #Collect fwhm value in file e, two theta y
                            a = a+1
                        else:
                            pass
                        
            pdf = pd.DataFrame({'A': p_duplicates_l, 'B':p_duplicates_r}) #Add similar two theta positions to dataframe pdf
            pdf['C'] = pdf['A'] - pdf['B'] #Determine the difference between two theta values
            pdf = pdf.abs() #Determine the absolute difference (remove negative signs)
            pdf = pdf.sort_values('C', ascending=True).drop_duplicates(['A']) #Drop duplicates in column 'A'
            pdf = pdf.sort_values('C', ascending=True).drop_duplicates(['B']) #Drop duplicates in column 'B'
            pdf = pdf.sort_values('A', ascending=True).reset_index() #Sort dataframe with respect to values in A
            pdf = pdf.drop(['index'], axis=1) #Drop index of pdf
            del p_duplicates_l[:] #Empty list (two theta values in file i)
            del p_duplicates_r[:] #Empty list (two theta values in file e)
            
            dwd = pd.DataFrame({'A': w_duplicates_l, 'B':w_duplicates_r}) #Add similar fwhm values to dataframe pdf
            dwd['C'] = dwd['A'] - dwd['B'] #Determine the difference between fwhm values
            dwd = dwd.abs() #Determine the absolute difference (remove negative signs)
            wdf = dwd.sort_values('C', ascending=True).drop_duplicates(['A'])  #Drop duplicates in column 'A'
            wdf = wdf.sort_values('C', ascending=True).drop_duplicates(['B']) #Drop duplicates in column 'B'
            wdf = wdf.sort_values('A', ascending=True).reset_index() #Sort dataframe with respect to values in A
            wdf = wdf.drop(['index'], axis=1) #Drop index of pdf
            del w_duplicates_l[:] #Empty list (fwhm values in file i)
            del w_duplicates_r[:] #Empty list (fwhm values in file e)
     
            i_p = ((len(pdf))/len(d[i]['two_theta'])*100) #Determine % number of matched peaks in file i, positions
            e_p = ((len(pdf))/len(d[e]['two_theta'])*100) #Determine % number of matched peaks in file e, positions   
            i_w = ((len(wdf))/len(d[i]['two_theta'])*100) #Determine % number of matched peaks in file i, widths
            e_w = ((len(wdf))/len(d[e]['two_theta'])*100) #Determine % number of matched peaks in file e, widths

            mdf.loc[(i[:EP_len_label]), i[:EP_len_label]] = 100 #Set same column/row name as 100% match
            mdf.loc[(e[:EP_len_label]), i[:EP_len_label]] = i_p #Determine % number of matched peaks in file i
            mdf.loc[(i[:EP_len_label]), e[:EP_len_label]] = e_p #Determine % number of matched peaks in file e
            mdf.loc[(e[:EP_len_label]), e[:EP_len_label]] = 100 #Set same column/row name as 100% match
            mwdf.loc[(i[:EP_len_label]), i[:EP_len_label]] = 100 #Set same column/row name as number of widths/peaks
            mwdf.loc[(e[:EP_len_label]), i[:EP_len_label]] = i_w  #Determine % number of matched peaks in file i
            mwdf.loc[(i[:EP_len_label]), e[:EP_len_label]] = e_w #Determine % number of matched peaks in file e (same as i)
            mwdf.loc[(e[:EP_len_label]), e[:EP_len_label]] = 100 #Set same column/row name as number of widths/peaks
    
    mdf.style.applymap(color_pos).to_excel(writer, 'Position_Match', index=True) #Write dataframe, mdf, to excel
    mwdf.style.applymap(color_wid).to_excel(writer, 'Width_Match', index=True) #Write dataframe, mwdf, to excel
    writer.save() #Save excel file
    return mdf, wdf

def color_pos(val): #Generate colour map for matched peak positions
    cmap = plt.get_cmap('RdYlGn', 10) #Generate 10 colours from colourmap
    colours = cmap(np.arange(0,cmap.N)) #Fix colours 
    cmap = [rgb2hex(rgb) for rgb in colours] #Convert colours to hex values
    for i, j in zip(range(0, 100, 10), cmap): #For % invervals of 10, up to 100
        if val >= i: #If value >= % interval
            color = j #Select representative colour
    return 'background-color: {}'.format(color)

def color_wid(val): #Generate colour map for matched peak widths
    cmap = plt.get_cmap('RdYlGn', 10) #Generate 10 colours from colourmap
    colours = cmap(np.arange(0,cmap.N)) #Fix colours  
    cmap = [rgb2hex(rgb) for rgb in colours] #Convert colours to hex values
    for i, j in zip(range(0, 100, 10), cmap): #For invervals of 10, up to 100
        if val >= i: #If value >= interval
            color = j  #Select representative colour
    return 'background-color: {}'.format(color)

def Clustering_Loop_1(file_path, EP_len_label, Position_tol, Width_tol): #Control Loop
    file_path_c = file_path + "/Baseline_Removed/Best/Peaks"
    os.chdir(file_path_c) #Change directory to desired directory
    retval = os.getcwd() #Get current directory
    print ("Clustering_Loop_1: Current working directory %s" % retval) #Show current directory
    writer = pd.ExcelWriter(file_path + '/Clustered_Data.xlsx', engine='openpyxl') #Create an excel spreadsheet
    DF1 = Read_in_DataFrame(file_path_c) #Read in DataFrame, DF1, containing extracted peaks
    mdf, wdf = Similarities(DF1, EP_len_label, Position_tol, Width_tol, writer) #Send to function Similarities (Identifying similarities between PXRD patterns)
    print('Finished')
    return 