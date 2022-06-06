import os
import re
import glob
import numpy as np
import math
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
from matplotlib.colors import rgb2hex

def Read_in_DataFrame(file_path, file_path_pr, Position_Match, Width_Match, writer): #Read in DataFrames containing extracted peaks (experimental and predicted)
    loc = file_path + "/Clustered_Data.xlsx" #Locate desired excel file
    loc1 = file_path_pr + "/EvPr_Clustered_Data.xlsx" #Locate desired excel file
    fcpw = pd.read_excel(loc, 'Final_Clustering_P' + str(Position_Match) + '%_W' + str(Width_Match) + '%') #Read in first sheet as % of matched peak positions
    fcpw = fcpw.loc[:,~fcpw.columns.str.startswith('W')]
    evpr = pd.read_excel(loc1, 'EvPr_Clustering') #Read in second sheet as matched peak widths
    evpr = evpr.iloc[:, 2:]
    final = pd.concat([fcpw, evpr], axis=1, sort=False)
    final=final.replace(np.nan,'None') #Make nan values blank
    final.style.applymap(color, subset=(final.columns[2:])).to_excel(writer, 'Clustered_Data', index=False)
    writer.save()
    return

def color(val): #Generate colour map
    cmap = plt.get_cmap('hsv', 30) #Generate colours from colourmap
    colours = cmap(np.arange(0,cmap.N)) #Fix colours  
    cmap = [rgb2hex(rgb) for rgb in colours] #Convert colours to hex values
    for i, j in zip(range(0, 30, 1), cmap): #For invervals of 1, up to 30
        if str(val) == 'None' or str(val) == 'Unknown': #If cell is blank
            color = 'white' #Make background colour white
        elif type(val) == str:
            if int(val[-1]) >= i: #If group assigned is >= i 
                color = j #Select representative colour
        elif type(val) == int or type(val) == float:
            if val >= i: #If group assigned is >= i 
                color = j #Select representative colour
    return 'background-color: {}'.format(color)

def Combine_Clustered_Data(file_path, file_path_pr, Position_Match, Width_Match): #Control Loop
    retval = os.getcwd() #Get current directory
    print ("Combine Clustered Data: Current working directory %s" % retval) #Show current directory
    writer = pd.ExcelWriter(file_path_pr + '/Combined_Clustered_Data.xlsx', engine='openpyxl') #Create an excel spreadsheet
    Read_in_DataFrame(file_path, file_path_pr, Position_Match, Width_Match, writer) #Read in DataFrame, DF1, containing extracted peaks
    print('Finished')