import os
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import pandas as pd
import glob
import re
from natsort import natsorted
import peakutils
import shutil
from shutil import copyfile
import xrdtools
import math
import matplotlib.colors as cl
import lmfit.models
from openpyxl import load_workbook

def Read_File_Extension_Data(file_name, file_extension): #Read in data
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

def Read_Data(file_name): #Normalise pattern
    data = np.loadtxt(file_name, usecols = (0,1))  #Load in data from file
    max_theta = data[:, 0] #Set first column as two theta
    BR_intensity = data[:, 1] #Set second column as intensity
    BR_intensity =  BR_intensity/max(BR_intensity) #Normalise intensities
    np.seterr(divide='ignore', invalid='ignore') #Ignore floating point errors
    mean_BI = np.mean(BR_intensity) #Determine the mean intensity of the pattern after normalisation
    return max_theta, BR_intensity, mean_BI

def Baseline_Removal(file_path, file_name, two_theta, intensity, two_theta_limit): #Function to remove the background of the data
    mask = two_theta < two_theta_limit #Hightlight only values below the set upper two theta limit
    max_theta = two_theta[mask] #Obtain two_theta values up to the limit
    intensity = intensity[mask] #Obtain intensity values up to the limit
    baseline_values = peakutils.baseline(intensity, deg=5, max_it=10, tol = 1e-3) #Calculate the baseline
    BR_intensity = np.array(intensity-baseline_values) #Remove the calculated baseline from the data
    indexesToReplace = BR_intensity < 0.0 #Hightlight negative values from baseline removal
    BR_intensity[indexesToReplace] = 0 #Set negative values to 0
    if not os.path.exists('Baseline_Removed'): #If original data folder does not exist
        os.makedirs('Baseline_Removed') #Make original data folder
    os.chdir('Baseline_Removed') #Change directory crystalline folder
    n = open(file_name, "w+") #Write a new file (same extension as set)
    for number, line in zip(max_theta, BR_intensity): #For each line in max_theta and picked_i
        n.write(str(number) + ' ' + str(line) + '\n') #Write lines into new file
    n.close() #Close file
    os.chdir(file_path) #Change directory back to intial file path
    return

def select_patterns(file_path, file_name, two_theta_limit, group_size, file_extension, Local_search_threshold, MtMR_min, MtMR_max, EP_len_label): #Select desirable patterns based on their Max to Mean Ratio
    os.chdir(file_path +'/Baseline_Removed') #Change the directory to files which have the background removed
    filenames = [] #Create list for desired files
    for ex in file_extension: #For each given file extension
        filenames.extend(natsorted(glob.glob(ex))) #Identify files with set file extension
    DF = pd.DataFrame(columns={'file_name', 'MtMR'}) #Create new dataframe DF, to contain Max to Mean ratios
    DF = DF[['file_name', 'MtMR']] #Fix the column order
    for l_n, file_name in enumerate(filenames): #For each file
        max_theta, BR_intensity, mean_BI = Read_Data(file_name) #Read in data
        MtMR= (np.max(BR_intensity))/(np.mean(BR_intensity)) #Calculate the maximum to mean ratio (MtMR)
        DF = DF.append({'file_name': file_name[:EP_len_label], 'MtMR':MtMR}, ignore_index=True) #Add all values to new row in DF
    MDF=DF.sort_values('MtMR', ascending=True) #Sort values with respect to MtMR column
    cs = plt.cm.brg(np.linspace(0,1,len(filenames)+1)) #Make a colourmap based on the number of files in the folder

    col = [] #Create new list to store order of colours with respect to files
    for c in MDF['file_name']: #For each file name
        col.append(cs[DF.set_index('file_name').index.get_loc(c)]) #Identify the corresponding colour code
    MDF.plot(x='file_name', y='MtMR', kind='bar', color=col, legend=None, fontsize=5) #Plot all MtMR values
    plt.xlabel('Filename') #Set x axis as filenames
    plt.ylabel('Maximum to Mean Ratio') #Set y axis as MtMR values
    plt.savefig('Calculated Maximum to Mean Ratio' + ".pdf", bbox_inches = "tight") #Save MtMR figure
    
    mask = (MtMR_max >= DF['MtMR']) & (DF['MtMR'] >= MtMR_min) #Locate all files with MtMR values within the inputted desired range
    BF = DF['file_name'].values[mask] #Where the mask is true, idenitfy files (BF)
    OF = DF['file_name'].values[~mask] #Where the mask is flase, identify files (OF)
    DF['Folder_Place'] = mask #Create new column to store mask results
    DF = DF.replace({True: 'Best', False: 'Other'}) #If DF value is 'True', replace with 'Best' (folder type), and vise versa
    DF.MtMR = DF.MtMR.apply(int) # round MtMR values
    DF['Number_of_Peaks'] = 0
    Number_of_peaks_to_excel(file_path, DF)
    
    for i in BF: #For files in BF
        for j in filenames: #For all files
            if str(i) in str(j): #Identify full file name
                Best_Folder(file_path, file_extension, j)  #Send to Best_Folder function
    for i in OF: #For files in OF
        for j in filenames: #For all files
            if str(i) in str(j): #Identify full file name
                Other_Folder(file_path, file_extension, j) #Send to Other_Folder function       
    return DF

def Best_Folder(file_path, file_extension, j): #Create crystalline folder, if more peaks picked than set limit
    file_path = os.getcwd() #Get set file path
    if not os.path.exists('Best'): #If crystalline folder does not exist
        os.makedirs('Best') #Make crystalline directory
    new_folder = file_path + '/Best' #Make new folder called 'Best'
    shutil.copy(file_path + '/' + j , new_folder + '/' + j) #Copy original data from original file path to new folder
    os.chdir(file_path) #Change directory back to intial file path
    return

def Other_Folder(file_path, file_extension, j): #Create crystalline folder, if more peaks picked than set limit
    file_path = os.getcwd() #Get set file path
    if not os.path.exists('Other'): #If crystalline folder does not exist
        os.makedirs('Other') #Make crystalline directory
    new_folder = file_path + '/Other' #Make new folder called 'Other'
    shutil.copy(file_path + '/' + j , new_folder + '/' + j) #Copy original data from original file path to new folder
    os.chdir(file_path) #Change directory back to intial file path
    return 

def Create_Groups_np(max_theta, BR_intensity, group_size): #Divide data into groups
    Number_of_groups = len(max_theta) / group_size #Calculate the number of groups required depending on the set group size
    two_theta_groups = np.array_split(max_theta, Number_of_groups) #Split the two theta data into groups
    BR_intensity_groups = np.array_split(BR_intensity, Number_of_groups) #Split the intensity data into groups
    mean_BR_intensity_groups = [] #New list for two theta groups
    mean_two_theta_groups =[] #New list for intensity groups
    for i in range(int(Number_of_groups)): # for each group determined
        mean_BR_intensity_groups.append(np.mean(BR_intensity_groups[i]))#Calulate the mean of each intensity group, add to new list
        mean_two_theta_groups.append(two_theta_groups[i][0]) #Determine the first value of each group, add to new list
    mean_two_theta_groups = np.asarray(mean_two_theta_groups) #Make mean two theta groups list an array 
    mean_BR_intensity_groups = np.asarray(mean_BR_intensity_groups) #Make mean intensity groups an array 
    return mean_two_theta_groups, mean_BR_intensity_groups, group_size

def Derivative(max_theta, BR_intensity, mean_two_theta_groups, mean_BR_intensity_groups, thres_factor, mean_BI): #Calculate derivatives in the data
    dx = mean_two_theta_groups[1:] - mean_two_theta_groups[:-1] #Calculate dx of neighbouring two theta groups
    dy = mean_BR_intensity_groups[1:] - mean_BR_intensity_groups[:-1] #Calculate dy of neighbouring intensity groups
    derivative = dy/dx #Calculate derivavives
    indexesToReplace = derivative < 0 #Highlight negative derivative values
    derivative[indexesToReplace] = 0 #Make negative derivatives equal 0
    factor = thres_factor / np.mean(dx) #Calculate a suitable factor (from given thres_factor and mean of dx, as derivatives include dx)
    threshold= mean_BI*factor #Multiply the mean intensity of the data by the factor to give a desirable threshold (to identify peaks)
    derivative_mask = derivative > threshold #Create a True/False mask to indicate all derivatives greater than the threshold (to identify peak positions)
    return derivative, derivative_mask, threshold

def Derivative_2(file_name, max_theta, BR_intensity, derivative, derivative_mask, mean_two_theta_groups, mean_BR_intensity_groups): #Collect all two theta indexes which may correspond to a peak
    peak_index = [i for i, x in enumerate(derivative_mask) if x] #Create an index of all True statements in derivative_mask
    last_peak_index = peak_index[-1] #Identify the last True statement in derivative mask (lpi)
    peak_index.append(last_peak_index + 1) #Add the index succeeding lpi to the list peak_index
    for line in peak_index: #For each index in peak_index
        if line != peak_index[-1] and line!= peak_index[-2]: #If the index is not the last or penultimate in peak_index
            derivative_mask[line+1] = True #Ensure all succeeding lines are True in the derivative mask
    derivative_mask[0] = True #Make first line True
    Only_derivative = derivative[derivative_mask] #Obtain all True derivative values
    Only_two_theta = (mean_two_theta_groups[:-1])[derivative_mask] #Obtain corresponding two theta values
    return derivative, Only_derivative, Only_two_theta

def Pick_Peaks(max_theta, BR_intensity, Only_derivative, Only_two_theta): #Pick peaks from derivative values
    indicies = peakutils.indexes(Only_derivative, thres=(np.mean(BR_intensity)), min_dist=1.0) #Pick maxima of each indentified peak from dervative search
    derivative_peaks = Only_derivative[indicies] #Identify derivative values determined as peak maxima
    two_theta_position = Only_two_theta[indicies] #Ientify two theta values corresponding to peak maxima
    mask = np.in1d(max_theta, two_theta_position) #Identify peak maxima (corresponding two theta values) in max theta list
    peak_position = [i for i, x in enumerate(mask) if x] #Identify peak positions corresponding to each peak maxima
    return peak_position

def Local_Search(file_name, max_theta, BR_intensity, peak_position, Local_search_threshold, DF, EP_len_label): #Allow picked peaks to move towards the local maximum
    local_search_i = [] #Create new list for intensities maxima
    local_search_t = [] #Create new list for corresponding two theta values 
    for line in peak_position: #For each two theta position indexed as a peak (peak_position)
        max_peak_position = line + Local_search_threshold #Determine peak range allowed allowed
        if max_peak_position >= len(BR_intensity): #If the range goes beyond or is equal to the length of the data
            max_peak_position = len(BR_intensity) #Make the range maxima equal to the length of the data
        for i in range(max_peak_position): #For each number in the allowed range
            if BR_intensity[i] == BR_intensity[line:max_peak_position].max(): #Determine the max intensity value in the range
                local_search_i.append(BR_intensity[i]) #Add intensity value to list
                local_search_t.append(max_theta[i]) #Add corresponding two theta value to list       
    mask = np.in1d(max_theta, local_search_t) #Identify the new max_theta value corresponding to the peak mixima (from the local search)
    peak_position = [i for i, x in enumerate(mask) if x] #Identify the indexes of local search peaks 
    picked_i = np.where(mask==False, 0, BR_intensity) #Make all intensity values 0 which do not correspond to peak maxima from local search
    Number_of_peaks = (len(picked_i[picked_i>0])) #Determine the number of picked peaks)
    result = list(*np.where(DF['file_name'] == str(file_name[:EP_len_label])))[0] #Idenitify row number in DF which contains the named file
    DF.loc[result, 'Number_of_Peaks'] = (Number_of_peaks) #Change the indexed Number_of_Peaks row to number of picked peaks
    return picked_i, DF

def Number_of_peaks_to_excel(file_path, DF): #Write an excel spreadsheet containing file names, Max to Mean Ratios, folder location and number of peaks picked
    os.chdir(file_path + '/Baseline_Removed') #Change directory
    writer = pd.ExcelWriter('Peak Picking.xlsx', engine='openpyxl') #Create an excel spreadsheet
    DF.to_excel(writer, 'Peak_Picking', index=True) #Write dataframe, DF, to excel
    writer.save() #Save excel spreadsheet
    return

def Peak_Folder(file_name, max_theta, picked_i): #Create crystalline folder, if more peaks picked than set limit
    file_path = os.getcwd() #Get set file path
    if not os.path.exists('Peaks'): #If crystalline folder does not exist
        os.makedirs('Peaks') #Make crystalline directory       
    os.chdir('Peaks') #Change directory crystalline folder
    retval = os.getcwd() #Get current directory
    n = open(file_name, "w+") #Write a new file (same extension as set)
    for number, line in zip(max_theta, picked_i): #For each line in max_theta and picked_i
        n.write(str(number) + ' ' + str(line) + '\n') #Write lines into new file
    n.close() #Close file
    os.chdir(file_path) #Change directory back to original
    return 

def Visualise_Folders_og(file_extension, EP_len_label): #Create a visual image of data
    retval = os.getcwd() #Get current directory
    filenames = []
    for ex in file_extension:
        filenames.extend(natsorted(glob.glob(ex))) #Identify files with set file extension
    cs = plt.cm.brg(np.linspace(0,1,len(filenames)+1)) #Generate a colourmap
    i = 1 #Set as 1 for first file
    for file_name, c in zip(filenames, cs): #For each filename
        length_of_directory = (len(filenames)) #Determine number of files with the extension
        two_theta, intensity = Read_File_Extension_Data(file_name, file_extension) #Read in data from file
        if i == 0:
            ax0 = plt.subplot(length_of_directory, 1, i)
        plt.subplot(length_of_directory, 1, i, sharex=ax0) #Create subplots, rows=no.of.files, columns=1, current subplot = i 
        plt.plot(two_theta, intensity, label=file_name[:EP_len_label], color=c, linewidth=0.3) #Plot data with a label, random colour and a linewidth
        plt.legend(prop={'size': 5}, loc='right') #Fix size of legend
        plt.yticks([]) #Remove y axis as all data is normalised
        if i != len(filenames): #If file is not the last in the directory
            plt.tick_params(axis='x', which='both', bottom=False, top=False, labelbottom= False) #Remove x axis
        i = i+1 #Loop through data
    plt.xlabel('2θ (°)') #Set x axis label
    plt.ylabel('Normalised Intensity (a.u.)', horizontalalignment='center', y=(len(filenames)/2.0)) #Set y axis, central
    plt.savefig('Stacked' + ".pdf", bbox_inches = "tight") #Save subplots
    plt.clf() #Clear current figure
    return retval

def Visualise_Folders_br(file_extension, EP_len_label): #Create a visual image of data
    retval = os.getcwd() #Get current directory
    filenames = []
    for ex in file_extension:
        filenames.extend(natsorted(glob.glob(ex))) #Identify files with set file extension
    cs = plt.cm.brg(np.linspace(0,1,len(filenames)+1)) #Generate a colourmap
    i = 1 #Set as 1 for first file
    for file_name, c in zip(filenames, cs): #For each filename
        length_of_directory = (len(filenames)) #Determine number of files with the extension
        max_theta, BR_intensity, mean_BI = Read_Data(file_name) #Read in data from file
        if i == 0:
            ax0 = plt.subplot(length_of_directory, 1, i)
        plt.subplot(length_of_directory, 1, i, sharex=ax0) #Create subplots, rows=no.of.files, columns=1, current subplot = i 
        plt.plot(max_theta, BR_intensity, label=file_name[:EP_len_label], color=c, linewidth=0.3) #Plot data with a label, random colour and a linewidth
        plt.legend(prop={'size': 5}, loc='right') #Fix size of legend
        plt.yticks([]) #Remove y axis as all data is normalised
        if i != len(filenames): #If file is not the last in the directory
            plt.tick_params(axis='x', which='both', bottom=False, top=False, labelbottom= False) #Remove x axis
        i = i+1 #Loop through data
    plt.xlabel('2θ (°)') #Set x axis label
    plt.ylabel('Normalised Intensity (a.u.)', horizontalalignment='center', y=(len(filenames)/2.0)) #Set y axis, central
    plt.savefig('Stacked' + ".pdf", bbox_inches = "tight") #Save subplots
    plt.clf() #Clear current figure
    return retval

def Visualise_Folders(file_path_br, file_extension, EP_len_label): #Create a visual image of data
    df = pd.read_excel(io=file_path_br + "/Peak Picking.xlsx", sheet_name='Peak_Picking') #Obtain the Max to Mean Ratio file to locate the index of each file name (for colour assignment of pdfs)
    retval = os.getcwd() #Get current directory
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
        if i == 1:
            ax0 = plt.subplot(length_of_directory,1,i, sharex=ax0)
        if x == 0: #If sum is zero
            flabel = int(flabel) #Save it as an integer
            mask = (df['file_name'] == flabel) #Identify the file name in df
        mask2 = df.index[mask] #Identify the index of the file name in df
        max_theta, BR_intensity, mean_BI = Read_Data(file_name) #Read in data from file
        length_of_directory = (len(filenames)) #Determine number of files with the extension 
        plt.subplot(length_of_directory, 1, i) #Create subplots, rows=no.of.files, columns=1, current subplot = i 
        plt.plot(max_theta, BR_intensity, label=file_name[:EP_len_label], color=cs[mask2[0]], linewidth=0.5) #Plot data with a label, random colour and a linewidth
        plt.legend(prop={'size': 5}, loc='right') #Fix size of legend
        plt.yticks([]) #Remove y axis as all data is normalised
        if i != len(filenames): #If file is not the last in the directory
            plt.tick_params(axis='x', which='both', bottom=False, top=False, labelbottom= False) #Remove x axis
        i = i+1 #Loop through data
    plt.xlabel('2θ (°)') #Set x axis label
    plt.ylabel('Normalised Intensity (a.u.)', horizontalalignment='center', y=(len(filenames)/2.0)) #Set y axis, central
    plt.savefig('Stacked' + ".pdf", bbox_inches = "tight") #Save subplots
    plt.clf() #Clear current figure
    return retval

def Create_Groups(file_name, max_theta, BR_intensity, file_path_b, file_path_c, group_size): #Normalise data and divide data into groups
    mask = BR_intensity > 0 #Locate intensity peaks picked
    p_intensity = BR_intensity[mask] #Collect all intenisties picked
    p_two_theta = max_theta[mask] #Collect all relative positions picked
    os.chdir(file_path_b) #Change directory to file with baseline removed
    b_file = np.loadtxt(file_name) #Load in all files
    b_two_theta = b_file[:, 0] #Set first column as two theta
    b_intensity = b_file[:, 1] #Set second column as intensity
    b_intensity = b_intensity/b_intensity.max() #Normalise all intensiies
    os.chdir(file_path_c) #Change directory to file with baseline removed
    Number_of_groups = len(b_two_theta) / group_size #Calculate the number of groups required depending on the set group size
    b_two_theta_groups = np.split(b_two_theta, Number_of_groups) #Split the two theta data into groups
    b_intensity_groups = np.split(b_intensity, Number_of_groups) #Split the intensity data into groups
    mean_b_intensity_groups = [] #New list for two theta groups
    mean_b_two_theta_groups =[] #New list for intensity groups
    for i in range(int(Number_of_groups)): # for each group determined
        mean_b_intensity_groups.append(np.mean(b_intensity_groups[i]))#Calulate the mean of each intensity group, add to new list
        mean_b_two_theta_groups.append(np.mean(b_two_theta_groups[i])) #Determine the first value of each group, add to new list
    mean_b_two_theta_groups = np.asarray(mean_b_two_theta_groups) #Make mean two theta groups list an array 
    mean_b_intensity_groups = np.asarray(mean_b_intensity_groups) #Make mean intensity groups an array
    return b_two_theta, b_intensity, p_two_theta, p_intensity, b_two_theta_groups, b_intensity_groups, mean_b_two_theta_groups, mean_b_intensity_groups

def Peak_Range(b_two_theta, b_intensity, p_two_theta, b_two_theta_groups, b_intensity_groups, mean_b_two_theta_groups, mean_b_intensity_groups): #Identify the two theta and itensity range of each peak
    picked_peaks_in_b_two_theta = [] #Create new list for baseline removed picked potions
    minimum = [] #Create new list to collect position of the 'start' of the peak
    min_moved = [] #Create new list to collect position of the start of the peak
    for i in p_two_theta: #For each picked peak position
        for line_number, j in enumerate(b_two_theta_groups): #For each group in baseline removed positions
            if i in j: #Identify basline removed picked peak position
                picked_peaks_in_b_two_theta.append(line_number) #Collect position line number
    for j in picked_peaks_in_b_two_theta: #For each collected position line number
        while mean_b_intensity_groups[j-1] > mean_b_intensity_groups[j]: #While the previous intenisty is greater, move to the previous position
            j -=1
        while mean_b_intensity_groups[j-1] < mean_b_intensity_groups[j]: #While the previous intenisty is lower, move to the previous position
            j -= 1
        if j <= 0: #If j is less than or equal to zero
            minimum.append(0) #Make j equal to zero
        else:
            minimum.append(j) #Collect value of j
    for line_number,j in enumerate(minimum): #For each peak minimum value
        while np.isclose(mean_b_intensity_groups[j], mean_b_intensity_groups[j+1], atol=1e-3) == True: #While succeeding value is within a tolerance
            if j == len(mean_b_intensity_groups)-2: #If j is the penultimate value in the data
                min_moved.append(j) #Collect j
            else:
                j += 1 #Increase j by 1
        min_moved.append(j) #Collect j value (start of peak)
    
    maximum = [] #Create new list to collect position of the 'end' of the peak
    max_moved = [] #Create new list to collect position of the end of the peak
    for j in picked_peaks_in_b_two_theta: #For each collected position line number
        while mean_b_intensity_groups[j] < mean_b_intensity_groups[j+1]: #While the succeeding intenisty is greater, move to the succeeding position
            j +=1
        while mean_b_intensity_groups[j] > mean_b_intensity_groups[j+1]: #While the succeeding intenisty is lower, move to the succeeding position
            j += 1
        if j >= len(b_two_theta_groups):
            maximum.append(b_two_theta_groups[-1])
        else:
            maximum.append(j) #Collect value of j
    for line_number,j in enumerate(maximum): #For each peak maximum value
        while np.isclose(mean_b_intensity_groups[j-1], mean_b_intensity_groups[j], atol=1e-3) == True: #While previous value is within a tolerance
            j -= 1
        max_moved.append(j) #Collect j value (end of peak)
    
    min_i = [] #Create new list for start peak intenisty group
    min_t = [] #Create new list for start peak two theta group
    max_i = [] #Create new list for end peak intenisty group
    max_t = [] #Create new list for end peak two theta group
    for i, j in zip(min_moved, max_moved): #For collected values of the start and end of peaks
        min_i.append(mean_b_intensity_groups[i]) #Collect start peak intenisty group
        min_t.append(mean_b_two_theta_groups[i]) #Collect start peak two theta group
        max_i.append(mean_b_intensity_groups[j]) #Collect end peak intensity group
        max_t.append(mean_b_two_theta_groups[j]) #Collect end two theta group

    a_min_t = [] #Create new list for initial two theta values
    a_max_t = [] #Create new list for final two theta values
    for i, j in zip(min_t, max_t): #For collected start and end groups
        x = (min(b_two_theta, key=lambda x:abs(x-i))) #Identify closest two theta value to the start group mean
        y = (min(b_two_theta, key=lambda x:abs(x-j))) #Identify closest two theta value to the end group mean
        ll_x = list(b_two_theta).index(x) #Identify the line number of collected start two theta value
        ll_y = list(b_two_theta).index(y) #Identify the line number of collect end two theta value
        a_min_t.append(ll_x) #Collect all line numbers of inital two theta values
        a_max_t.append(ll_y) #Collect all line numbers of final two theta values
    return a_min_t, a_max_t, min_i, min_t, max_i, max_t

def Model_Fitting(file_name, EP_len_label, b_two_theta, b_intensity, p_two_theta, p_intensity, a_min_t, a_max_t, min_i, min_t, max_i, max_t): #Fit a model to the identified peaks
    DF = pd.DataFrame(columns={'file_name','two_theta', 'maxI', 'redchi', 'amplitude', 'centre', 'fwhm', 'height', 'step_size'}) #Create new dataframe (DF) to contain fitting parameters
    DF = DF[['file_name','two_theta', 'centre', 'maxI', 'height', 'redchi', 'fwhm', 'amplitude', 'step_size']] #Keep columns in set order
    for a, b, i, tt, mi in zip(a_min_t, a_max_t, range(0, len(min_i)), p_two_theta, p_intensity): #For collected values correposponding to each peak the the pattern 
        x = b_two_theta[a:b] #Identify two theta range of the peak
        y = b_intensity[a:b] #Identify intensity range of the peak
        step_size = b_two_theta[1]-b_two_theta[0]
        mask = y > 0
        x = x[mask]
        y = y[mask]
        mod = lmfit.models.PseudoVoigtModel() #Select the PseudoVoightModel()
        pars = mod.guess(y, x=x) #Guess the parameters of the peak
        out = mod.fit(y, pars, x=x, weights=np.sqrt(1.0/y)) #Fit the parameters to the peak
        redchi = out.redchi #Collect the reduced chi value
        amplitude = out.params['amplitude'].value #Collect the amplitude of the peak
        centre = out.params['center'].value #Collect the centre of the peak
        fwhm = out.params['fwhm'].value #Collect the full width half max of the peak
        height = out.params['height'].value #Collect the height of the peak
        DF = DF.append({'file_name': file_name, 'two_theta':tt, 'maxI':mi, 'redchi':redchi, 'amplitude':amplitude, 'centre':centre, 'fwhm':fwhm, 'height':height, 'step_size':step_size}, ignore_index=True) #Add all collected values to the relevant columns
    return DF

def Create_DF(file_name): #Create a new dataframe, DF1, to hold all of the fitted peak data from all files
    DF1 = pd.DataFrame(columns={'file_name','two_theta', 'maxI', 'redchi', 'amplitude', 'centre', 'fwhm', 'height', 'step_size'}) #Create relevant columns
    return

def Combine_DFs(file_name, DF1, DF): #Add data in DF to DF1
    frames = [DF1, DF] #Add data in DF to DF1
    DF1 = pd.concat(frames) #Save frames as DF1
    return DF1

def DF_to_excel(DF1): #Save data to an excel spreadsheet
    writer = pd.ExcelWriter('Fitted_Picked_Peaks.xlsx', engine='openpyxl') #Create an excel spreadsheet
    DF1.to_excel(writer, 'Fitted_Picked_Peaks', index=False) #Write dataframe into sheet 1
    writer.save() #Save spreadsheet
    DF1 = DF1.drop_duplicates(['file_name', 'centre'], keep='last')
    writer = pd.ExcelWriter('Unique_Fitted_Picked_Peaks.xlsx', engine='openpyxl') #Create an excel spreadsheet
    DF1.to_excel(writer, 'Unique_Fitted_Picked_Peaks', index=False) #Write dataframe into sheet 1
    writer.save() #Save spreadsheet  
    return writer

def main(file_path, file_name, two_theta_limit, group_size, file_extension, Local_search_threshold, ln): #Remove the baseline from all patterns
    two_theta, intensity = Read_File_Extension_Data(file_name, file_extension) #Read in data
    Baseline_Removal(file_path, file_name, two_theta, intensity, two_theta_limit) #Remove baseline
    return

def main_2(file_path, file_name, two_theta_limit, group_size, file_extension, Local_search_threshold, ln, max_theta, BR_intensity, thres_factor, mean_BI, DF, EP_len_label): #Pick Peaks
    mean_two_theta_groups, mean_BR_intensity_groups, group_size = Create_Groups_np(max_theta, BR_intensity, group_size) #Calculate two theta and intensity groups
    derivative, derivative_mask, threshold = Derivative(max_theta, BR_intensity, mean_two_theta_groups, mean_BR_intensity_groups, thres_factor, mean_BI) #Calculate the derivative of consecutive groups
    if sum(derivative_mask) > 0: #If the number of 'picked peaks' is greater than 0
        derivative, Only_derivative, Only_two_theta = Derivative_2(file_name, max_theta, BR_intensity, derivative, derivative_mask, mean_two_theta_groups, mean_BR_intensity_groups) #Collect all two theta indexes which may correspond to a peak
        peak_position = Pick_Peaks(max_theta, BR_intensity, Only_derivative, Only_two_theta) #Pick peaks from derivative values
        picked_i, Number_of_peaks = Local_Search(file_name, max_theta, BR_intensity, peak_position, Local_search_threshold, DF, EP_len_label) #Allow picked peaks to move towards the local maximum
    return picked_i, Number_of_peaks

def main3(file_path_c, file_name, file_path_b,  group_size, EP_len_label, DF1): #Fit a model to all picked peaks
    max_theta, BR_intensity, mean_BI = Read_Data(file_name)
    b_two_theta, b_intensity, p_two_theta, p_intensity, b_two_theta_groups, b_intensity_groups, mean_b_two_theta_groups, mean_b_intensity_groups = Create_Groups(file_name, max_theta, BR_intensity, file_path_b, file_path_c, group_size) #Normalise data and divide data into groups
    a_min_t, a_max_t, min_i, min_t, max_i, max_t = Peak_Range(b_two_theta, b_intensity, p_two_theta, b_two_theta_groups, b_intensity_groups, mean_b_two_theta_groups, mean_b_intensity_groups) #Identify the two theta and itensity range of each peak
    DF = Model_Fitting(file_name, EP_len_label, b_two_theta, b_intensity, p_two_theta, p_intensity, a_min_t, a_max_t, min_i, min_t, max_i, max_t) #Fit a model to the identified peaks
    DF1 = Combine_DFs(file_name, DF1, DF) #Add data in DF to DF1
    return DF1

def Peak_Picking_Loop(file_path, two_theta_limit, group_size, file_extension, Local_search_threshold, EP_len_label, MtMR_min, MtMR_max, thres_factor): #Control Loop
    file_path_c = file_path + "/Baseline_Removed/Best/Peaks"
    file_path_b = file_path + "/Baseline_Removed/Best"
    file_path_br = file_path + "/Baseline_Removed"
    os.chdir(file_path) #Change directory to desired directory
    retval = os.getcwd() #Get current directory
    print ("Peak_Picking_Loop: Current working directory %s" % retval) #Show current directory
    filenames = [] #Create list for desired files
    for ex in file_extension: #For each given file extension
        filenames.extend(natsorted(glob.glob(ex))) #Identify files with set file extension
    nf = len(filenames)-1 #Determine number of files (enummerate starts at 0, hence -1)
    Visualise_Folders_og(file_extension, EP_len_label) ###########
    for ln, f in enumerate(filenames): #For each filename
        file_name= str(f)  #Make the filename a string
        #print ("File name is %s" % file_name) #Show filename
        main(file_path, file_name, two_theta_limit, group_size, file_extension, Local_search_threshold, ln) #Send to function main (Remove the baseline from all patterns)
        os.chdir(file_path_br) #Change directory
        Visualise_Folders_br(file_extension, EP_len_label) #Send to function Visualise_Folders (Create a visual image of data)
        os.chdir(file_path) #Change directory
        if ln == nf: #If all files have had their baselines removed
            DF = select_patterns(file_path, file_name, two_theta_limit, group_size, file_extension, Local_search_threshold, MtMR_min, MtMR_max, EP_len_label) #Send to function select_patterns (Select desirable patterns based on their Max to Mean Ratio)
    os.chdir(file_path_b) #Change directory
    Visualise_Folders(file_path_br, file_extension, EP_len_label) #Send to function Visualise_Folders (Create a visual image of data)
    retval = os.getcwd() #Get current directory
    print ("Current working directory %s" % retval) #Show current directory
    filenames = [] #Create list for desired files
    for ex in file_extension: #For each given file extension
        filenames.extend(natsorted(glob.glob(ex))) #Identify files with set file extension
    for ln, f in enumerate(filenames): #For each filename
        file_name= str(f)  #Make the filename a string
        #print ("File name is %s" % file_name) #Show filename 
        max_theta, BR_intensity, mean_BI = Read_Data(file_name) #Send to function Read_Data (Normalise all patterns)
        picked_i, DF = main_2(file_path, file_name, two_theta_limit, group_size, file_extension, Local_search_threshold, ln, max_theta, BR_intensity, thres_factor, mean_BI, DF, EP_len_label) #Send to function main_2 (Pick Peaks)
        Peak_Folder(file_name, max_theta, picked_i) #Create crystalline folder, if more peaks picked than set limit
    Number_of_peaks_to_excel(file_path, DF) # Send to function Number_of_peaks_to_excel (Write an excel spreadsheet containing file names, Max to Mean Ratios, folder location and number of peaks picked)
    os.chdir(file_path_c) #Change directory to desired directory
    Visualise_Folders(file_path_br, file_extension, EP_len_label) #Send to function Visualise_Folders (Create a visual image of data)
    os.chdir(file_path_c) #Change directory to desired directory
    retval = os.getcwd() #Get current directory
    print ("Current working directory %s" % retval) #Show current directory
    DF1 = Create_DF(file_path_c) #Send to function Create_DF (Create a new dataframe, DF1, to hold all of the fitted peak data from all files)
    filenames = [] #Create list for desired files
    for ex in file_extension: #For each given file extension
        filenames.extend(natsorted(glob.glob(ex))) #Identify files with set file extension
    for f in filenames: #For each filename
        file_name= str(f)  #Make the filename a string
        #print ("File name is %s" % file_name) #Show filename
        DF1 = main3(file_path_c, file_name, file_path_b,  group_size, EP_len_label, DF1) #Send to function main (Fit a model to all picked peaks)
    writer = DF_to_excel(DF1) #Send to function DF_to_excel (Save data to an excel spreadsheet)
    print('Finished')
    return
