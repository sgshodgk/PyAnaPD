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

def Read_in_DataFrame(file_path_pr): #Read in DataFrame, DF1, containing extracted peaks
    loc = file_path_pr + "/Clustered_Data.xlsx" #Locate desired excel file
    mdf = pd.read_excel(loc, 'Position_Match') #Read in first sheet as % of matched peak positions
    mwdf = pd.read_excel(loc, 'Width_Match') #Read in second sheet as matched peak widths
    mdf = mdf.set_index('files') #Set first column (file names) as index
    mwdf = mwdf.set_index('files') #Set first column (file names) as index
    return mdf, mwdf

def Similarities(mdf, mwdf, Position_Match, Width_Match, writer): #Identifying similar PXRD patterns
    sfmdf = mdf[mdf >= Position_Match].fillna(0) #Highlight values >= set Percentage match value, else 0
    sfmwdf = mwdf[mwdf >= Width_Match].fillna(0) #Highlight values >= set similar widths value, else 0
    sfmdf.style.applymap(color_pos).to_excel(writer, 'Position_Match > ' + str(Position_Match) +'%', index=True) #Write dataframe, sfmdf, to excel
    sfmwdf.style.applymap(color_wid).to_excel(writer, 'Width_Match > ' + str(Width_Match) +'%', index=True) #Write dataframe, sfmwdf, to excel
    writer.save() #Save excel file
    return 

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
    for i, j in zip(range(0, 100, 10), cmap): #For invervals of of 10, up to 100
        if val >= i: #If value >= interval
            color = j  #Select representative colour
    return 'background-color: {}'.format(color)

def Pr_Clustering_Loop_2(file_path_pr, Position_Match, Width_Match): #Control Loop
    retval = os.getcwd() #Get current directory
    print ("Clustering_Loop_2: Current working directory %s" % retval) #Show current directory
    writer = pd.ExcelWriter(file_path_pr + '/Clustered_Data.xlsx', engine='openpyxl') #Create an excel spreadsheet
    writer.book = load_workbook(file_path_pr + '/Clustered_Data.xlsx') #Set writer as a book
    mdf, mwdf = Read_in_DataFrame(file_path_pr) #Read in DataFrame, DF1, containing extracted peaks
    Similarities(mdf, mwdf, Position_Match, Width_Match, writer) #Send to function Similarities (Identifying similarities between PXRD patterns)
    print('Finished')
    return