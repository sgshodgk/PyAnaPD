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

def Read_File_Extension_Data(file_name): #Read in data
    file_extension = os.path.splitext(file_name)[1] #Identify file extension type
    if file_extension == '.xrdml': #If files are in a xrdml formal
        data = xrdtools.read_xrdml(file_name) #Read in xrdml files
        two_theta= data['x'] #Collect two theta data
        intensity=data['data'] #Collect intensity data
    else:
        l = [] #Create new list
        data = open(file_name) #Open data file
        for line_number, line in enumerate(data): #For each line, line number in the data
            if re.match('^[a-zA-Z]+', line) is not None or line.startswith('&'): #If the line starts with a character or &
                pass #Ignore
            else: #Otherwise
                l.append(line_number) #Add the line number to a new list
                break #Break loop after first line number, with no character or &, is found
        data = np.loadtxt(file_name, skiprows=l[0], usecols = (0,1))  #Load in data from file, skipping the rows containing a character or &
        two_theta = data[:, 0] #Set first column as two theta
        intensity = data[:, 1] #Set second column as intensity
    intensity = intensity/intensity.max() #Normalise intensity
    np.seterr(divide='ignore', invalid='ignore') #Ignore floating point errors
    return two_theta, intensity

def Read_in_DataFrame(file_path_c): #Read in DataFrame, DF1, containing extracted peaks
    loc = file_path_c + "/Unique_Fitted_Picked_Peaks.xlsx" #Locate the excel file
    DF1 = pd.read_excel(loc) #Read in file as a dataframe, DF1
    return DF1

def Peak_in_range(peak_minimum, peak_maximum, file_path_c, file_path, DF1, filenames): #Determine files have a peak between specified two theta values given
    PRDF = DF1[DF1['two_theta'].between(peak_minimum, peak_maximum, inclusive=False)] #Idenify files with a peak between specified values
    PRDF_files = PRDF.file_name.values #Idendify the file names, record as PDF_files
    for f in filenames: #For each file 
        file_name = str(f) #Record the file name
        if f in PRDF_files: #If the file is in PDF_files
            if not os.path.exists('Peak between ' + str(peak_minimum) + ' & ' + str(peak_maximum)+ ' 2θ (°)'):  #If a peak between folder does not exist
                os.makedirs('Peak between ' + str(peak_minimum) + ' & ' + str(peak_maximum)+ ' 2θ (°)') #Make peak between directory
            new_folder = file_path_c + '/Peak between ' + str(peak_minimum) + ' & ' + str(peak_maximum)+ ' 2θ (°)' #Combine current directory with new folder extension
            shutil.copy(file_path + '/' + file_name , new_folder + '/' + file_name) #Copy original data from original file path to new folder
            os.chdir(file_path_c) #Change directory 
        else:
            if not os.path.exists('No Peak between ' + str(peak_minimum) + ' & ' + str(peak_maximum)+ ' 2θ (°)'):  #If a no peak between folder does not exist
                os.makedirs('No Peak between ' + str(peak_minimum) + ' & ' + str(peak_maximum)+ ' 2θ (°)') #Make peak between directory
            new_folder = file_path_c + '/No Peak between ' + str(peak_minimum) + ' & ' + str(peak_maximum)+ ' 2θ (°)' #Combine current directory with new folder extension
            shutil.copy(file_path + '/' + file_name , new_folder + '/' + file_name) #Copy original data from original file path to new folder
            os.chdir(file_path_c) #Change directory
    return

def main(peak_minimum_two_theta, peak_maximum_two_theta, file_path_c, file_path, filenames): #Main loop to control functions
    DF1 = Read_in_DataFrame(file_path_c) #Read in DataFrame, DF1, containing extracted peaks
    Peak_in_range(peak_minimum_two_theta, peak_maximum_two_theta, file_path_c, file_path, DF1, filenames) #Determine files have a peak between specified two theta values given
    return 

def twotheta_range(file_path_pb, file_extension, EP_len_label): #Determine two theta ranges of collected patterns
    os.chdir(file_path_pb) #Change directory
    max_tt = [] #Create new list of maximum collected two theta values
    min_tt = [] #Create new list of minimum collected two theta values
    filenames = [] #Create list for desired files
    for ex in file_extension: #For each given file extension
        filenames.extend(natsorted(glob.glob(ex))) #Identify files with set file extension
    for f in filenames: #For each filename
        file_name= str(f) #Make the filename a string
        two_theta, intensity = Read_File_Extension_Data(file_name) #Send to function Read_Data (Read in data)  
        max_tt.append(max(two_theta)) #Collect the maximum two theta value
        min_tt.append(min(two_theta)) #Collect the minumum two theta value
    max_tt = min(max_tt) # #Identify the minumum, maximum two theta value
    min_tt = max(min_tt) #Identify the maximum, minimum two theta value
    return max_tt, min_tt

def Visualise_Folder(file_path_pb, file_path_c, file_path_br, file_extension, EP_len_label, max_tt, min_tt, file): #Visualise the data in each folder
    df = pd.read_excel(io=file_path_br + "/Peak Picking.xlsx", sheet_name='Peak_Picking') #Obtain the Max to Mean Ratio file to locate the index of each file name (for colour assignment of pdfs)
    os.chdir(file_path_pb) #Change directory
    filenames = [] #Create list for desired files
    for ex in file_extension: #For each given file extension
        filenames.extend(natsorted(glob.glob(ex))) #Identify files with set file extension
    cs = plt.cm.brg(np.linspace(0,1,len(df)+1)) #Generate a colourmap
    i = 1 #Set as 1 for first file
    for f in filenames: #For each filename
        file_name= str(f) #Make the filename a string
        flabel = file_name[:EP_len_label] #Identify the reduced filename
        mask = (df['file_name'] == flabel) #Identify the file name in df
        x=np.sum(mask) #Check sum of True values in mask (should be 1)
        if x == 0: #If sum is zero
            flabel = int(flabel) #Save it as an integer
            mask = (df['file_name'] == flabel) #Identify the file name in df
        mask2 = df.index[mask] #Identify the index of the file name in df
        two_theta, intensity = Read_File_Extension_Data(file_name) #Read in data
        length_of_directory = (len(filenames)) #Determine number of files with the extensions
        mask = (min_tt <= two_theta) &  (two_theta <= max_tt)
        two_theta = two_theta[mask] #Obtain two_theta values up to the limit
        intensity = intensity[mask] #Obtain intensity values up to the limit         
        plt.subplot(length_of_directory, 1, i) #Create subplots, rows=no.of.files, columns=1, current subplot = i
        plt.plot(two_theta, intensity, label=file_name[:EP_len_label], color=cs[mask2[0]], linewidth=0.5) #Plot data with a label, random colour and a linewidth
        plt.legend(prop={'size': 5}, loc='right',) #Fix size of legend
        plt.yticks([]) #Remove y axis as all data is normalised
        if i != len(filenames): #If file is not the last in the directory
            plt.tick_params(axis='x', which='both', bottom=False, top=False, labelbottom= False) #Remove x axis
        i = i+1 #Loop through data
    plt.xlabel('2θ (°)', fontsize = 12) #Set x axis label
    plt.ylabel('Normalised Intensity (a.u.)', horizontalalignment='center', y=(length_of_directory/2.0), fontsize = 12) #Set y axis, central
    plt.savefig('2θ_Stacked_' + str(file)  + ".pdf",dpi=100) #Save subplots
    plt.clf() #Clear current figure
    os.chdir(file_path_c) #Change directory back to orginal file path
    return

def Visualisation_Loop(file_path, file_path_c, file_path_br, file_extension, EP_len_label, peak_minimum, peak_maximum): #Control loop
    if os.path.exists('Peak between ' + str(peak_minimum) + ' & ' + str(peak_maximum)+ ' 2θ (°)'): #If 'peak between' folder exists
        file = 'Peak between ' + str(peak_minimum) + ' & ' + str(peak_maximum)+ ' 2θ (°)' #Save name of folder
        file_path_pb = file_path_c + '/Peak between ' + str(peak_minimum) + ' & ' + str(peak_maximum)+ ' 2θ (°)' #Change path to specified directory
        max_tt, min_tt, = twotheta_range(file_path_pb, file_extension, EP_len_label) #Determine two theta ranges of collected patterns
        Visualise_Folder(file_path_pb, file_path_c, file_path_br, file_extension, EP_len_label, max_tt, min_tt, file)  #Visulaise the data in each folder
    if os.path.exists('No Peak between ' + str(peak_minimum) + ' & ' + str(peak_maximum)+ ' 2θ (°)'): #If 'no peak between' folder exists
        file = 'No Peak between ' + str(peak_minimum) + ' & ' + str(peak_maximum)+ ' 2θ (°)' #Save name of folder
        file_path_pb = file_path_c + '/No Peak between ' + str(peak_minimum) + ' & ' + str(peak_maximum)+ ' 2θ (°)' #Change path to specified directory
        max_tt, min_tt, = twotheta_range(file_path, file_extension, EP_len_label) #Determine two theta ranges of collected patterns
        Visualise_Folder(file_path_pb, file_path_c, file_path_br, file_extension, EP_len_label, max_tt, min_tt, file)  #Visulaise the data in each folder
    return 

def Peak_in_Range_Loop(file_path, file_extension, EP_len_label, peak_minimum, peak_maximum): #Control loop
    file_path_c = file_path + "/Baseline_Removed/Best/Peaks"
    file_path_b = file_path + "/Baseline_Removed/Best"
    file_path_br = file_path + "/Baseline_Removed"
    os.chdir(file_path_c) #Change directory to desired directory
    retval = os.getcwd() #Get current directory
    print ("Peak_in_Range_Loop: Current working directory %s" % retval) #Show current directory
    filenames = [] #Create list for desired files
    for ex in file_extension: #For each given file extension
        filenames.extend(natsorted(glob.glob(ex))) #Identify files with set file extension
    main(peak_minimum, peak_maximum, file_path_c, file_path, filenames) #Send data to main loop
    Visualisation_Loop(file_path, file_path_c, file_path_br, file_extension, EP_len_label, peak_minimum, peak_maximum) #Control loop
    print('Finished')
    return