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

def Generate_Folders(file_path, file_path_c, file_path_pr, file_path_pp, file_path_ppp, file_extension, filenames, filenames_pr, EP_len_label, PP_len_label, writer, Visualisations_Clust): #Generate folders for matched patterns
    spm = pd.read_excel(file_path_ppp + "/Selected_Matching.xlsx") #Locate the desired file
    final = pd.DataFrame(columns=['Filename', 'File_name', 'PPL', 'N']) #New Dataframe to assign groups
    lenf = 0 #Length of Final Dataframe
    a = 0
    for i in spm['PP'].unique(): #For each unique predicted pattern
        for j in spm['EP'][spm['PP']==i]: #Indentify matched experimental pattern
            for k in filenames: #For each experimental pattern
                if str(j) in str(k[:EP_len_label]): #If matched experimental pattern 
                    file_path_o = file_path #Identify experimental file path
                    if Visualisations_Clust == 'Yes':
                        make_folder(file_path_o, k, i) #Copy data into predicted pattern folder
                    final.loc[lenf,  'Filename'] = k #Add filename
                    final.loc[lenf,  'File_name'] = k[:EP_len_label] #Add reduced filename
                    final.loc[lenf, 'PPL'] = 'P!'+ str(i) + '_' + str(a) #Assign group number
                    lenf = lenf+1 #Increase length of Final Dataframe by 1
        a = a+1
        for k in filenames_pr: #For each predicted pattern
            if str(i) in str(k[:PP_len_label]): #If matched predicted pattern 
                file_path_o = file_path_pr #Identify predicted file path
                if Visualisations_Clust == 'Yes':
                    make_folder(file_path_o, k, i) #Copy data into predicted pattern folder
                
    for k in filenames: #For experimental patterns 
        if str(k[:EP_len_label]) not in str(spm['EP']): #If experimental pattern is not matched to a predicted pattern
            file_path_o = file_path #Identify experimental file path
            i = 'Unknown' #Set predicted pattern folder as 'Unknown'
            if Visualisations_Clust == 'Yes':
                make_folder(file_path_o, k, i) #Copy data into predicted pattern folder
            final.loc[lenf,  'Filename'] = k #Add filename
            final.loc[lenf,  'File_name'] = k[:EP_len_label] #Add reduced filename
            final.loc[lenf, 'PPL'] = 'P!Unknown' #Assign group number
            lenf = lenf+1 #Increase length of Final Dataframe by 1
           
    finall = final.fillna('').groupby(['Filename','File_name'], sort=False, as_index=False).agg({'PPL': ''.join}) #Group rows by Filename, join PL and WL rows
    e = finall['PPL'].str.split('P!', expand=True).rename(columns = lambda x: "PPL"+str(x)) #Split PL column into separate columns at letter 'P'
    result = pd.concat([finall, e], axis=1, sort=False) #Combine all columns 
    result=result.replace('',np.nan) #Make nan values blank
    result = result.dropna(axis='columns', how='all').set_index(['Filename']) #Drop first expanded columns (first P/W counts to give NAN columns)
    del result['PPL'] #Delete 'PL' column
    result = result.reindex(natsorted(result.index)).reset_index()
    
    result.style.applymap(color, subset=(result.columns[2:])).to_excel(writer, 'EvPr_Clustering', index=False) #Write result dataframe to execel
    writer.save() #Save excel file

    if Visualisations_Clust == 'Yes':
        Visualise_Folders(file_path, file_path_c, file_path_ppp, file_extension, EP_len_label, PP_len_label, filenames, filenames_pr) #Create a visual image of data
    return

def color(val): #Generate colour map
    cmap = plt.get_cmap('hsv', 30) #Generate colours from colourmap
    colours = cmap(np.arange(0,cmap.N)) #Fix colours  
    cmap = [rgb2hex(rgb) for rgb in colours] #Convert colours to hex values
    for i, j in zip(range(0, 30, 1), cmap): #For invervals of 1, up to 30
        if str(val) == 'None' or str(val) == 'Unknown': #If cell is blank
            color = 'white' #Make background colour white
        else:
            if int(val[-1]) >= i: #If group assigned is >= i 
                color = j #Select representative colour
    return 'background-color: {}'.format(color)

def make_folder(file_path_o, k, i): #Copy data into predicted pattern folder
    file_path = os.getcwd() #Get set file path
    if not os.path.exists('EvPr/' + str(i)): #If predicted pattern folder does not exist
        os.makedirs('EvPr/' +  str(i)) #Make predicted pattern directory
    new_folder = file_path + '/EvPr/' + str(i) #Highlifgt new folder named the predicted pattern
    shutil.copy(file_path_o + '/' + k , new_folder + '/' + k) #Copy original data from original file path to new folder
    os.chdir(file_path) #Change directory back to intial file path
    return

def twotheta_range(file_path_ppp, file_extension, file): #Determine two theta ranges of collected patterns
    max_tt = [] #Create new list of maximum collected two theta values
    min_tt = [] #Create new list of minimum collected two theta values
    os.chdir(file_path_ppp + '/EvPr' + '/' + file) #Change path to specified directory
    filenames = [] #Create list for desired files
    for ex in file_extension: #For each given file extension
        filenames.extend(natsorted(glob.glob(ex))) #Identify files with set file extension
    i = 1 #Set as 1 for first file
    for f in filenames: #For each filename
        file_name= str(f) #Make the filename a string
        two_theta, intensity = Read_File_Extension_Data(file_name) #Read in data
        max_tt.append(max(two_theta)) #Collect the maximum two theta value
        min_tt.append(min(two_theta)) #Collect the minumum two theta value
    max_tt = min(max_tt) # #Identify the minumum, maximum two theta value
    min_tt = max(min_tt) #Identify the maximum, minimum two theta value
    return max_tt, min_tt

def Visualise_Folders(file_path, file_path_c, file_path_ppp, file_extension, EP_len_label, PP_len_label, filenames, filenames_pr): #Create a visual image of data
    df = pd.read_excel(io=file_path + "/Baseline_Removed/Peak Picking.xlsx", sheet_name='Peak_Picking') #Obtain the Max to Mean Ratio file to locate the index of each file name (for colour assignment of pdfs)
    directories = [name for name in os.listdir(file_path_ppp + '/EvPr') if os.path.isdir(os.path.join(file_path_ppp + '/EvPr', name))] #Identify directories in the filepath
    if 'Visualisations' in directories:
        directories.remove('Visualisations')
    for file in directories: #For each directory
        max_tt, min_tt = twotheta_range(file_path_ppp, file_extension, file)
        os.chdir(file_path_ppp + '/EvPr/' + file) #Change path to specified directory
        filenames_c = [] #Create list for desired files
        for ex in file_extension: #For each given file extension
            filenames_c.extend(natsorted(glob.glob(ex))) #Identify files with set file extension
        i = 1 #Set as 1 for first file
        cs = plt.cm.brg(np.linspace(0,1,len(df)+1)) #Generate a colourmap
        for f in filenames_c: #For each filename
            if f in filenames_pr:
                file_name= str(f) #Make the filename a string
                c = 'k'
            else:
                file_name= str(f) #Make the filename a string
                flabel = file_name[:EP_len_label]
                mask = df['file_name'] == flabel #Identify the file name in df
                x=np.sum(mask) #Check sum of True values in mask (should be 1)
                if x == 0: #If sum is zero
                    flabel = int(flabel) #Save it as an integer
                    mask = df['file_name'] == flabel #Identify the file name in df
                mask2 = df.index[mask] #Identify the index of the file name in df
                c = cs[mask2[0]]
            two_theta, intensity = Read_File_Extension_Data(file_name)
            length_of_directory = (len(filenames_c)) #Determine number of files with the extension    
            mask = (min_tt <= two_theta) & (two_theta <= max_tt) #Determine a mask so all data is within the same two theta range
            two_theta = two_theta[mask] #Obtain two_theta values up to the limit
            intensity = intensity[mask] #Obtain intensity values up to the limit   
            plt.subplot(length_of_directory, 1, i) #Create subplots, rows=no.of.files, columns=1, current subplot = i 
            plt.plot(two_theta, intensity, label=file_name[:PP_len_label], color=c, linewidth=0.5) #Plot data with a label, random colour and a linewidth
            plt.legend(prop={'size': 5}, loc='right') #Fix size of legend
            plt.yticks([]) #Remove y axis as all data is normalised
            if i != len(filenames_c): #If file is not the last in the directory
                plt.tick_params(axis='x', which='both', bottom=False, top=False, labelbottom= False) #Remove x axis
            i = i+1 #Loop through data
        plt.xlabel('2θ (°)', fontsize = 12) #Set x axis label
        plt.ylabel('Normalised Intensity (a.u.)', horizontalalignment='center', y=(length_of_directory/2.0), fontsize = 12) #Set y axis, central
        plt.savefig('2θ_Stacked_' + str(file)  + ".pdf",dpi=100) #Save subplots
        plt.clf() #Clear current figure
        plt.close()
    return

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
        #skiprows issue
        data = np.loadtxt(file_name, skiprows=1, usecols = (0,1))  #Load in data from file, skipping the rows containing a character or & 
        two_theta = data[:, 0] #Set first column as two theta
        intensity = data[:, 1] #Set second column as intensity
    intensity = intensity/intensity.max() #Normalise intensity
    np.seterr(divide='ignore', invalid='ignore') #Ignore floating point errors
    return two_theta, intensity

def ExperimentalvPredicted_Clustering(file_path, file_path_pr, file_extension, EP_len_label, PP_len_label, Visualisations_Clust): #Control Loop
    file_path_c = file_path + "/Baseline_Removed/Best/Peaks"
    file_path_pp = file_path_pr + "/Pattern_Present"
    file_path_ppp = file_path_pr + "/Pattern_Present/Peaks"
    os.chdir(file_path_c) #Change directory to desired directory
    retval = os.getcwd() #Get current directory
    print ("Predicted_Similarities_Loop: Current working directory %s" % retval) #Show current directory
    filenames = [] #Create list for desired files
    for ex in file_extension: #For each given file extension
        filenames.extend(natsorted(glob.glob(ex))) #Identify files with set file extension
    os.chdir(file_path_ppp) #Change directory to desired directory
    retval = os.getcwd() #Get current directory
    print ("Current working directory %s" % retval) #Show current directory
    filenames_pr = []
    for ex in file_extension: #For each given file extension
        filenames_pr.extend(natsorted(glob.glob(ex))) #Identify files with set file extension
    writer = pd.ExcelWriter(file_path_pr + '/EvPr_Clustered_Data.xlsx', engine='openpyxl') #Create an excel spreadsheet
    writer.book = load_workbook(file_path_pr + '/EvPr_Clustered_Data.xlsx') #Set writer as a book
    Generate_Folders(file_path, file_path_c, file_path_pr, file_path_pp, file_path_ppp, file_extension, filenames, filenames_pr, EP_len_label, PP_len_label, writer, Visualisations_Clust)  #Generate folders for matched patterns
    print('Finished')
    return