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

def Read_in_DataFrame(file_path): #Read in DataFrame, DF1, containing extracted peaks
    loc = file_path + "/Clustered_Data.xlsx" #Locate the desired file
    mdf = pd.read_excel(loc, 'Position_Match') #Read in first sheet as % of matched peak positions (MP)
    mwdf = pd.read_excel(loc, 'Width_Match') #Read in second sheet as matched peak widths (MW)
    mdf = mdf.set_index('files') #Set first column (file names) as index
    mwdf = mwdf.set_index('files') #Set first column (file names) as index
    mdf.index = mdf.index.map(str) #Make all index elements strings
    mwdf.index = mwdf.index.map(str) #Make all index elements strings
    return mdf, mwdf

def Similarities(mdf, mwdf, Position_Match, Width_Match): #Identifying similar PXRD patterns
    fmdf = mdf >= Position_Match #Mask values >= Percentage Match
    fmdf =  fmdf & fmdf.T #Tranpose data and overlay both masks
    
    mmdf=pd.DataFrame(columns=['together', 'NoE', 'rec']) #Create Dataframe to contain matched files
    for ln,i in enumerate(fmdf.columns): #For each column in MP
        a = fmdf.index[fmdf[i]].tolist() #Identify all True masked elements
        a = natsorted(a) #Order elements
        mmdf.loc[ln,  'together'] = a #Add elements to new row in dataframe
        mmdf.loc[ln,  'NoE'] = len(a) #Add the number of elements 
    for ln, li in enumerate(mmdf['together']): #For each row of matched elements 
        a = mmdf[mmdf['together'].apply(lambda x: set(tuple(li)).issubset(tuple(x)))] #Identify if row is within another row
        mmdf.loc[ln, 'rec'] = len(a) #Add number of records of the row
        
    mmdf['together'] = mmdf['together'].apply(str) #Make column strings
    mmdf = mmdf.drop_duplicates() #Remove duplicated rows
    k = mmdf[mmdf['NoE'] != mmdf['rec']] #Identify rows where the number of elements does not equal the number of records
    edf = pd.DataFrame(columns=['together', 'NoE', 'rec']) #Create Dataframe to contain rematched files
    for ln, a in enumerate(k['together']): #For each row in k (where the number of elements does not equal the number of records)
        for ln1, b in enumerate(k['together']): #For each row in k (where the number of elements does not equal the number of records)
            if ln != ln1: #If the row a is not row b
                c = [] #Create new list ('elements are in row b')
                d = [] #Create new list ('elements are not in row b')
                for i in a.split(', '): #Split row a into elements (where ', ')
                    i = str(i).strip('[]')#.replace('"', '') #Strip brackets from elements
                    if i in b: #If element is in b
                        c.append(i) #Add to list c
                    if i not in b: #If element is not in b
                        d.append(i) #Add to list d    
                if len(c) > 1: #If c contains more than one elements
                    edf.loc[len(edf), 'together'] = c #Add elements to new row in dataframe
                    edf.loc[len(edf)-1, 'NoE'] = len(c) #Add the number of elements
                if len(d) > 1: #If d contains more than one elements
                    edf.loc[len(edf), 'together'] = d #Add elements to new row in dataframe
                    edf.loc[len(edf)-1, 'NoE'] = len(d) #Add the number of elements
    for ln, li in enumerate(edf['together']): #For each row of rematched elements 
        a = edf[edf['together'].apply(lambda x: set(tuple(li)).issubset(tuple(x)))] #Identify if row is within another row
        edf.loc[ln, 'rec'] = len(a) #Add number of records of the row
        
    edf['together'] = edf['together'].apply(str) #Make column strings
    edf=edf.groupby(edf.columns.tolist(),as_index=False).size().reset_index().rename(columns={0:'records'}) #Group together duplicated rows (rematched)
    edf = edf[edf['records']>1] #Identify rows which were duplicated (rematched)
    mmdf=mmdf.groupby(mmdf.columns.tolist(),as_index=False).size().reset_index().rename(columns={0:'records'}) #Group together duplicated rows (matched)
    mmdf = pd.concat([mmdf[(mmdf['NoE'] == mmdf['rec'])], edf]).reset_index(drop = True) #Combine matched and rematched dataframes (Clustered Positions)
    
    fmwdf = mwdf >= Width_Match #Mask values >= Width_Match
    fmwdf = fmwdf & fmwdf.T & fmdf #Tranpose data and overlay masks, including peak position match mask
    
    mmwdf=pd.DataFrame(columns=['together', 'NoE', 'rec']) #Create Dataframe to contain matched files
    for ln,i in enumerate(fmwdf.columns): #For each column in MW
        a = fmwdf.index[fmwdf[i]].tolist() #Identify all True masked elements
        a = natsorted(a) #Order elements
        mmwdf.loc[ln,  'together'] = a #Add elements to new row in dataframe
        mmwdf.loc[ln,  'NoE'] = len(a) #Add the number of elements 
    for ln, li in enumerate(mmwdf['together']): #For each row of matched elements 
        a = mmwdf[mmwdf['together'].apply(lambda x: set(tuple(li)).issubset(tuple(x)))] #Identify if row is within another row
        mmwdf.loc[ln, 'rec'] = len(a) #Add number of records of the row

    mmwdf['together'] = mmwdf['together'].apply(str) #Make column strings
    mmwdf = mmwdf.drop_duplicates() #Remove duplicated rows
    k = mmwdf[mmwdf['NoE'] != mmwdf['rec']] #Identify rows where the number of elements does not equal the number of records
    edf = pd.DataFrame(columns=['together', 'NoE', 'rec']) #Create Dataframe to contain rematched files
    for ln, a in enumerate(k['together']): #For each row in k (where the number of elements does not equal the number of records)
        for ln1, b in enumerate(k['together']): #For each row in k (where the number of elements does not equal the number of records)
            if ln != ln1: #If the row a is not row b
                c = [] #Create new list ('elements are in row b')
                d = [] #Create new list ('elements are not in row b')
                for i in a.split(', '):
                    i = str(i).strip('[]')#.replace('"', '') 
                    if i in b: #If element is in b
                        c.append(i) #Add to list c
                    if i not in b: #If element is not in b
                        d.append(i) #Add to list d 
                if len(c) > 1: #If c contains more than one elements  
                    edf.loc[len(edf), 'together'] = c #Add elements to new row in dataframe
                    edf.loc[len(edf)-1, 'NoE'] = len(c) #Add the number of elements
                if len(d) > 1: #If d contains more than one elements #str(a) != str(d).strip('[]').replace('"', '') and len(d) > 1:
                    edf.loc[len(edf), 'together'] = d #Add elements to new row in dataframe
                    edf.loc[len(edf)-1, 'NoE'] = len(d) #Add the number of elements
    for ln, li in enumerate(edf['together']): #For each row of rematched elements 
        a = edf[edf['together'].apply(lambda x: set(tuple(li)).issubset(tuple(x)))] #Identify if row is within another row
        edf.loc[ln, 'rec'] = len(a) #Add number of records of the row

    edf['together'] = edf['together'].apply(str) #Make column strings
    edf=edf.groupby(edf.columns.tolist(),as_index=False).size().reset_index().rename(columns={0:'records'}) #Group together duplicated rows (rematched)
    edf = edf[edf['records'] > 1] #Identify rows which were duplicated (rematched)
    mmwdf=mmwdf.groupby(mmwdf.columns.tolist(),as_index=False).size().reset_index().rename(columns={0:'records'}) #Group together duplicated rows (matched)
    mmwdf = pd.concat([mmwdf[(mmwdf['NoE'] == mmwdf['rec'])], edf]).reset_index(drop = True) #Combine matched and rematched dataframes (Clustered Widths)
    return mmdf, mmwdf

def Group_Assignment(mmdf, mmwdf, filenames, EP_len_label, Position_Match, Width_Match, writer):
    final = pd.DataFrame(columns=['Filename', 'File_name', 'PL', 'WL']) #New Dataframe to assign groups
    plen = len(mmdf) #Length of Clustered Positions dataframe
    wlen = len(mmwdf) #Length of Clustered Widths dataframe
    lenf = 0 #Length of Final Dataframe
    for fn, f in enumerate(filenames): #For each filename
        a = 0 
        b = 0
        for ln, p in enumerate(mmdf['together']): #For each Clustered Postions row
            if f[:EP_len_label] in p: #If filename is within row
                final.loc[lenf,  'Filename'] = f #Add filename
                final.loc[lenf,  'File_name'] = f[:EP_len_label] #Add reduced filename
                final.loc[lenf, 'PL'] = 'P!'+ str(ln) #Assign group number 
                a = a+1 #Filename has been assigned a group
                lenf = lenf+1 #Increase length of Final Dataframe by 1
         
        for ln, w in enumerate(mmwdf['together']): #For each Clustered Widths row
            if f[:EP_len_label] in w: #If filename is within row
                final.loc[lenf,  'Filename'] = f #Add filename
                final.loc[lenf,  'File_name'] = f[:EP_len_label] #Add reduced filename
                final.loc[lenf, 'WL'] = 'W!' + str(ln) #Assign group number 
                b = b+1 #Filename has been assigned a group
                lenf = lenf+1  #Increase length of Final Dataframe by 1
                
        if a == 0 : #If filename is not in Clustered Positions Dataframe
            final.loc[lenf,  'Filename'] = f #Add filename
            final.loc[lenf,  'File_name'] = f[:EP_len_label] #Add reduced filename
            final.loc[lenf, 'PL'] = 'P!'+ str(plen) #Assign group number (as length of Clustered Positions dataframe)
            plen = plen+1 #Increase length of Clustered Positions Dataframe by 1
            lenf = lenf+1 #Increase length of Final Dataframe by 1
            
        if b == 0 : #If filename is not in Clustered Positions Dataframe
            final.loc[lenf,  'Filename'] = f #Add filename
            final.loc[lenf,  'File_name'] = f[:EP_len_label] #Add reduced filename
            final.loc[lenf, 'WL'] = 'W!'+ str(wlen) #Assign group number (as length of Clustered Positions dataframe)
            wlen = wlen+1 #Increase length of Clustered Widths Dataframe by 1
            lenf = lenf+1 #Increase length of Final Dataframe by 1

    finall = final.fillna('').groupby(['Filename','File_name'], sort=False, as_index=False).agg({'PL': ''.join, 'WL': ''.join}) #Group rows by Filename, join PL and WL rows
    e = finall['PL'].str.split('P!', expand=True).rename(columns = lambda x: "PL"+str(x)) #Split PL column into separate columns at letter 'P' 
    f = finall['WL'].str.split('W!', expand=True).rename(columns = lambda x: "WL"+str(x)) #Split WL column into separate columns at letter 'W' 
    result = pd.concat([finall, e, f], axis=1, sort=False) #Combine all columns 
    result=result.replace('',np.nan) #Make nan values blank
    result = result.dropna(axis='columns', how='all').set_index(['Filename'])  #Drop first expanded columns (first P/W counts to give NAN columns)
    del result['PL'] #Delete 'PL' column
    del result['WL'] #Delete WL' column
    result = result.reindex(natsorted(result.index)).reset_index()

    result.style.applymap(color, subset=(result.columns[2:])).to_excel(writer, 'Final_Clustering_P' + str(Position_Match) + '%_W' + str(Width_Match) + '%', index=False) #Write result dataframe to execel
    writer.save() #Save excel file
    return final, result

def color(val): #Generate colour map
    cmap = plt.get_cmap('hsv', 30) #Generate colours from colourmap
    colours = cmap(np.arange(0,cmap.N)) #Fix colours  
    cmap = [rgb2hex(rgb) for rgb in colours] #Convert colours to hex values
    for i, j in zip(range(0, 30, 1), cmap): #For invervals of 1, up to 30
        if str(val) == 'None': #If cell is blank
            color = 'white' #Make background colour white
        else:
            if int(val) >= i: #If group assigned is >= i 
                color = j #Select representative colour
    return 'background-color: {}'.format(color)

def Generate_Clustered_Folders(filenames, file_path, file_path_c, EP_len_label, writer, final, result, Visualisations_P, Visualisations_W): #Generate clustered folders
    c_file_path = os.getcwd() #Get current directory
    if Visualisations_P == 'Yes':
        for p in result.columns[result.columns.str.startswith('P')]: #For columns starting with 'P'
            for name, i in zip(result['Filename'], result[str(p)]): #For each filename and assigned group
                if str(i) != 'None': #If cell is not 'None'
                    if not os.path.exists('Positions/Similar_Files' + '_' + str(i)): #If folder does not exist, make folder with column heading
                        os.makedirs('Positions/Similar_Files' + '_' + str(i)) #Make new folder
                    new_folder = c_file_path + '/' + 'Positions/Similar_Files' + '_' + str(i) #Combine current directory with new folder extension
                    shutil.copy(file_path + '/' + str(name) , new_folder + '/' + str(name)) #Copy original data from original file path to new folderto new folder
                    os.chdir(file_path_c) #Change directory to desired directory
    if Visualisations_W == 'Yes':
        for w in result.columns[result.columns.str.startswith('W')]: #For columns starting with 'W'
            for name, i in zip(result['Filename'], result[str(w)]): #For each filename and assigned group
                if str(i) != 'None':  #If cell is not 'None'
                    if not os.path.exists('Widths/Similar_Files' + '_' + str(i)): #If folder does not exist, make folder with column heading
                        os.makedirs('Widths/Similar_Files' + '_' + str(i)) #Make new folder
                    new_folder = c_file_path + '/' + 'Widths/Similar_Files' + '_' + str(i) #Combine current directory with new folder extension
                    shutil.copy(file_path + '/' + str(name) , new_folder + '/' + str(name)) #Copy original data from original file path to new folderto new folder
                    os.chdir(file_path_c) #Change directory to desired directory
    return

def twotheta_range(file_path, file_extension, file): #Determine two theta ranges of collected patterns
    max_tt = [] #Create new list of maximum collected two theta values
    min_tt = [] #Create new list of minimum collected two theta values
    os.chdir(file_path + '/' + file) #Change path to specified directory
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

def Visualise_Folder(file_path,  file_path_br, file_extension, EP_len_label): #Visualise the data in each folder
    df = pd.read_excel(io=file_path_br + "/Peak Picking.xlsx", sheet_name='Peak_Picking') #Obtain the Max to Mean Ratio file to locate the index of each file name (for colour assignment of pdfs)
    directories = [name for name in os.listdir(file_path) if os.path.isdir(os.path.join(file_path, name))] #Identify directories in the filepath
    for file in directories: #For each directory
        max_tt, min_tt = twotheta_range(file_path, file_extension, file)
        os.chdir(file_path + '/' + file) #Change path to specified directory
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
            plt.legend(prop={'size': 5}, loc='right') #Fix size of legend
            plt.yticks([]) #Remove y axis as all data is normalised
            if i != len(filenames): #If file is not the last in the directory
                plt.tick_params(axis='x', which='both', bottom=False, top=False, labelbottom= False) #Remove x axis
            i = i+1 #Loop through data
        plt.xlabel('2θ (°)', fontsize = 12) #Set x axis label
        plt.ylabel('Normalised Intensity (a.u.)', horizontalalignment='center', y=(length_of_directory/2.0), fontsize = 12) #Set y axis, central
        plt.savefig('2θ_Stacked_' + str(file)  + ".pdf",dpi=100, bbox_inches = "tight") #Save subplots
        plt.clf() #Clear current figure
        os.chdir(file_path) #Change directory back to orginal file path
    return

def Visualisation_Loop(file_path_c, file_path_br, file_extension, EP_len_label, Visualisations_P, Visualisations_W): #Control loop for data visualisation
    if Visualisations_P == 'Yes':
        file_path = file_path_c + '/Positions' #Change path to specified directory
        Visualise_Folder(file_path, file_path_br, file_extension, EP_len_label)  #Visulaise the data in each folder
    if Visualisations_W == 'Yes':
        file_path = file_path_c + '/Widths' #Change path to specified directory
        Visualise_Folder(file_path, file_path_br, file_extension, EP_len_label)  #Visulaise the data in each folder
    return 

def Clustering_Loop_3(file_path, file_extension, EP_len_label, Position_Match, Width_Match, Visualisations_P, Visualisations_W): #Control Loop
    file_path_c = file_path + "/Baseline_Removed/Best/Peaks"
    file_path_br = file_path + "/Baseline_Removed"
    os.chdir(file_path_c) #Change directory to desired directory
    retval = os.getcwd() #Get current directory
    print ("Clustering_Loop_3: Current working directory %s" % retval) #Show current directory
    filenames = [] #Create list for desired files
    for ex in file_extension: #For each given file extension
        filenames.extend(natsorted(glob.glob(ex))) #Identify files with set file extension
    writer = pd.ExcelWriter(file_path + '/Clustered_Data.xlsx', engine='openpyxl') #Create an excel spreadsheet
    writer.book = load_workbook(file_path + '/Clustered_Data.xlsx')
    mdf, mwdf = Read_in_DataFrame(file_path) #Read in DataFrame, DF1, containing extracted peaks
    mmdf, mmwdf = Similarities(mdf, mwdf, Position_Match, Width_Match) #Send to function Similarities (Identifying similarities between PXRD patterns)
    final, result = Group_Assignment(mmdf, mmwdf, filenames, EP_len_label, Position_Match, Width_Match, writer)
    Generate_Clustered_Folders(filenames, file_path, file_path_c, EP_len_label, writer, final, result, Visualisations_P, Visualisations_W) #Generate clustered folders
    Visualisation_Loop(file_path_c, file_path_br, file_extension, EP_len_label, Visualisations_P, Visualisations_W)  #Control loop for data visualisation
    print('Finished')
    return