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

def Read_File_Extension_Data(file_name, file_extension, two_theta_limit): #Read in data
    retval = os.getcwd() #Get current directory
    data = np.loadtxt(file_name, skiprows=1, usecols = (0,1))  #Load in data from file, skipping the rows containing a character or &
    two_theta = data[:, 0] #Set first column as two theta
    intensity = data[:, 1] #Set second column as intensity
    mask = two_theta < two_theta_limit #Hightlight only values below the set upper two theta limit
    max_theta = two_theta[mask] #Obtain two_theta values up to the limit
    intensity = intensity[mask] #Obtain intensity values up to the limit
    if np.count_nonzero(intensity) > 0:
        intensity = intensity/intensity.max() #Normalise intensity
        if not os.path.exists('Pattern_Present'): #If original data folder does not exist
            os.makedirs('Pattern_Present') #Make original data folder
        os.chdir('Pattern_Present') #Change directory crystalline folder
        n = open(file_name, "w+") #Write a new file (same extension as set)
        for number, line in zip(max_theta, intensity): #For each line in max_theta and picked_i
            n.write(str(number) + ' ' + str(line) + '\n') #Write lines into new file
        n.close() #Close file
        os.chdir(retval) #Change directory back to intial file path
    np.seterr(divide='ignore', invalid='ignore') #Ignore floating point errors
    return intensity

def Read_Data(file_name): #Normalise pattern
    data = np.loadtxt(file_name, skiprows= 1, usecols = (0,1))  #Load in data from file
    max_theta = data[:, 0] #Set first column as two theta
    intensity = data[:, 1] #Set second column as intensity
    intensity = intensity/intensity.max() #Normalise intensity
    np.seterr(divide='ignore', invalid='ignore') #Ignore floating point errors
    mean_BI = np.mean(intensity) #Determine the mean intensity of the pattern after normalisation
    return max_theta, intensity, mean_BI

def Create_Groups_np(max_theta, intensity, group_size): #Divide data into groups
    Number_of_groups = len(max_theta) / group_size #Calculate the number of groups required depending on the set group size
    two_theta_groups = np.array_split(max_theta, Number_of_groups) #Split the two theta data into groups
    intensity_groups = np.array_split(intensity, Number_of_groups) #Split the intensity data into groups
    mean_intensity_groups = [] #New list for two theta groups
    mean_two_theta_groups =[] #New list for intensity groups
    for i in range(int(Number_of_groups)): # for each group determined
        mean_intensity_groups.append(np.mean(intensity_groups[i]))#Calulate the mean of each intensity group, add to new list
        mean_two_theta_groups.append(two_theta_groups[i][0]) #Determine the first value of each group, add to new list
    mean_two_theta_groups = np.asarray(mean_two_theta_groups) #Make mean two theta groups list an array 
    mean_intensity_groups = np.asarray(mean_intensity_groups) #Make mean intensity groups an array 
    return mean_two_theta_groups, mean_intensity_groups, group_size

def Derivative(max_theta, intensity, mean_two_theta_groups, mean_intensity_groups, thres_factor, mean_BI): #Calculate derivatives in the data
    dx = mean_two_theta_groups[1:] - mean_two_theta_groups[:-1] #Calculate dx of neighbouring two theta groups
    dy = mean_intensity_groups[1:] - mean_intensity_groups[:-1] #Calculate dy of neighbouring intensity groups
    derivative = dy/dx #Calculate derivavives
    indexesToReplace = derivative < 0 #Highlight negative derivative values
    derivative[indexesToReplace] = 0 #Make negative derivatives equal 0
    factor = thres_factor / np.mean(dx) #Calculate a suitable factor (from given thres_factor and mean of dx, as derivatives include dx)
    threshold= mean_BI*factor #Multiply the mean intensity of the data by the factor to give a desirable threshold (to identify peaks)
    derivative_mask = derivative > threshold #Create a True/False mask to indicate all derivatives greater than the threshold (to identify peak positions)
    return derivative, derivative_mask, threshold

def Derivative_2(file_name, max_theta, intensity, derivative, derivative_mask, mean_two_theta_groups, mean_intensity_groups): #Collect all two theta indexes which may correspond to a peak
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

def Pick_Peaks(max_theta, intensity, Only_derivative, Only_two_theta): #Pick peaks from derivative values
    indicies = peakutils.indexes(Only_derivative, thres=(np.mean(intensity)), min_dist=1.0) #Pick maxima of each indentified peak from dervative search
    derivative_peaks = Only_derivative[indicies] #Identify derivative values determined as peak maxima
    two_theta_position = Only_two_theta[indicies] #Ientify two theta values corresponding to peak maxima
    mask = np.in1d(max_theta, two_theta_position) #Identify peak maxima (corresponding two theta values) in max theta list
    peak_position = [i for i, x in enumerate(mask) if x] #Identify peak positions corresponding to each peak maxima
    return peak_position

def Local_Search(file_name, max_theta, intensity, peak_position, Local_search_threshold, P_DF, PP_len_label): #Allow picked peaks to move towards the local maximum
    local_search_i = [] #Create new list for intensities maxima
    local_search_t = [] #Create new list for corresponding two theta values 
    for line in peak_position: #For each two theta position indexed as a peak (peak_position)
        max_peak_position = line + Local_search_threshold #Determine peak range allowed allowed
        if max_peak_position >= len(intensity): #If the range goes beyond or is equal to the length of the data
            max_peak_position = len(intensity) #Make the range maxima equal to the length of the data
        for i in range(max_peak_position): #For each number in the allowed range
            if intensity[i] == intensity[line:max_peak_position].max(): #Determine the max intensity value in the range
                local_search_i.append(intensity[i]) #Add intensity value to list
                local_search_t.append(max_theta[i]) #Add corresponding two theta value to list       
    mask = np.in1d(max_theta, local_search_t) #Identify the new max_theta value corresponding to the peak mixima (from the local search)
    peak_position = [i for i, x in enumerate(mask) if x] #Identify the indexes of local search peaks 
    picked_i = np.where(mask==False, 0, intensity) #Make all intensity values 0 which do not correspond to peak maxima from local search
    Number_of_peaks = (len(picked_i[picked_i>0])) #Determine the number of picked peaks)
    result = list(*np.where(P_DF['file_name'] == str(file_name[:PP_len_label])))[0] #Idenitify row number in P_DF which contains the named file
    P_DF.loc[result, 'Number_of_Peaks'] = (Number_of_peaks) #Change the indexed Number_of_Peaks row to number of picked peaks
    return picked_i, P_DF

def Number_of_peaks_to_excel(P_DF): #Write an excel spreadsheet containing file names, Max to Mean Ratios, folder location and number of peaks picked
    writer = pd.ExcelWriter('Peak Picking.xlsx', engine='openpyxl') #Create an excel spreadsheet
    P_DF.to_excel(writer, 'Peak_Picking', index=True) #Write dataframe, P_DF, to excel
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

def Visualise_Folders(file_path_pp, file_extension, PP_len_label): #Create a visual image of data
    p_df = pd.read_excel(io=file_path_pp + "/Peak Picking.xlsx", sheet_name='Peak_Picking') #Obtain the Max to Mean Ratio file to locate the index of each file name (for colour assignment of pdfs)
    retval = os.getcwd() #Get current directory
    filenames = [] #Create list for desired files
    for ex in file_extension: #For each given file extension
        filenames.extend(natsorted(glob.glob(ex))) #Identify files with set file extension
    cs = plt.cm.brg(np.linspace(0,1,len(p_df)+1)) #Generate a colourmap
    i = 1 #Set as 1 for first file
    for f in filenames: #For each filename
        file_name= str(f) #Make the filename a string
        flabel = file_name[:PP_len_label] #Identify the files label
        if file_name[:PP_len_label].isdigit(): #If the file label is only digits
            flabel = int(flabel) #Save it as an integer
        mask = (p_df['file_name'] == flabel) #Identify the file name in p_df
        if sum(mask) > 0:
            mask2 = p_df.index[mask] #Identify the index of the file name in p_df
            max_theta, intensity, mean_BI = Read_Data(file_name) #Read in data from file
            length_of_directory = (len(filenames)) #Determine number of files with the extension 
            plt.subplot(length_of_directory, 1, i) #Create subplots, rows=no.of.files, columns=1, current subplot = i 
            plt.plot(max_theta, intensity, label=file_name[:PP_len_label], color=cs[mask2[0]], linewidth=0.5) #Plot data with a label, random colour and a linewidth
            plt.legend(prop={'size': 5}) #Fix size of legend
            plt.yticks([]) #Remove y axis as all data is normalised
            if i != len(filenames): #If file is not the last in the directory
                plt.tick_params(axis='x', which='both', bottom=False, top=False, labelbottom= False) #Remove x axis
            i = i+1 #Loop through data
    plt.xlabel('2θ (°)') #Set x axis label
    plt.ylabel('Normalised Intensity (a.u.)', horizontalalignment='center', y=(len(filenames)/2.0)) #Set y axis, central
    plt.savefig('Stacked' + ".pdf") #Save subplots
    plt.clf() #Clear current figure
    return retval

def Create_Groups(file_name, max_theta, intensity, file_path_pp, file_path_ppp, group_size): #Normalise data and divide data into groups
    mask = intensity > 0 #Locate intensity peaks picked
    p_intensity = intensity[mask] #Collect all intenisties picked
    p_two_theta = max_theta[mask] #Collect all relative positions picked
    os.chdir(file_path_pp) #Change directory to file with baseline removed
    b_file = np.loadtxt(file_name) #Load in all files
    b_two_theta = b_file[:, 0] #Set first column as two theta
    b_intensity = b_file[:, 1] #Set second column as intensity
    b_intensity = b_intensity/b_intensity.max() #Normalise all intensiies
    os.chdir(file_path_ppp) #Change directory to file with baseline removed
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
    P_DF = pd.DataFrame(columns={'file_name','two_theta', 'maxI', 'redchi', 'amplitude', 'centre', 'fwhm', 'height', 'step_size'}) #Create new dataframe (DF) to contain fitting parameters
    P_DF = P_DF[['file_name','two_theta', 'centre', 'maxI', 'height', 'redchi', 'fwhm', 'amplitude', 'step_size']] #Keep columns in set order
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
        P_DF = P_DF.append({'file_name': file_name, 'two_theta':tt, 'maxI':mi, 'redchi':redchi, 'amplitude':amplitude, 'centre':centre, 'fwhm':fwhm, 'height':height, 'step_size':step_size}, ignore_index=True) #Add all collected values to the relevant columns
    return P_DF

def Create_DF(file_name): #Create a new dataframe, P_DF1, to hold all of the fitted peak data from all files
    P_DF1 = pd.DataFrame(columns={'file_name','two_theta', 'maxI', 'redchi', 'amplitude', 'centre', 'fwhm', 'height', 'step_size'}) #Create relevant columns
    return

def Combine_DFs(file_name, P_DF1, P_DF): #Add data in P_DF to P_DF1
    frames = [P_DF1, P_DF] #Add data in P_DF to P_DF1
    P_DF1 = pd.concat(frames) #Save frames as P_DF1
    return P_DF1

def DF_to_excel(P_DF1): #Save data to an excel spreadsheet
    writer = pd.ExcelWriter('Fitted_Picked_Peaks.xlsx', engine='openpyxl') #Create an excel spreadsheet
    P_DF1.to_excel(writer, 'Fitted_Picked_Peaks', index=False) #Write dataframe into sheet 1
    writer.save() #Save spreadsheet
    return writer

def main(file_path_pr, file_name, two_theta_limit, group_size, file_extension, Local_search_threshold, ln): #Remove the baseline from all patterns
    intensity = Read_File_Extension_Data(file_name, file_extension, two_theta_limit) #Read in data
    return intensity

def main_2(file_path_pr, file_name, two_theta_limit, group_size, file_extension, Local_search_threshold, ln, max_theta, intensity, thres_factor, mean_BI, P_DF, PP_len_label): #Pick Peaks
    mean_two_theta_groups, mean_intensity_groups, group_size = Create_Groups_np(max_theta, intensity, group_size) #Calculate two theta and intensity groups
    derivative, derivative_mask, threshold = Derivative(max_theta, intensity, mean_two_theta_groups, mean_intensity_groups, thres_factor, mean_BI) #Calculate the derivative of consecutive groups
    if sum(derivative_mask) > 0: #If the number of 'picked peaks' is greater than 0
        derivative, Only_derivative, Only_two_theta = Derivative_2(file_name, max_theta, intensity, derivative, derivative_mask, mean_two_theta_groups, mean_intensity_groups) #Collect all two theta indexes which may correspond to a peak
        peak_position = Pick_Peaks(max_theta, intensity, Only_derivative, Only_two_theta) #Pick peaks from derivative values
        picked_i, Number_of_peaks = Local_Search(file_name, max_theta, intensity, peak_position, Local_search_threshold, P_DF, PP_len_label) #Allow picked peaks to move towards the local maximum
    return picked_i, Number_of_peaks

def main3(file_path_ppp, file_name, file_path_pp,  group_size, PP_len_label, P_DF1): #Fit a model to all picked peaks
    max_theta, intensity, mean_BI = Read_Data(file_name)
    b_two_theta, b_intensity, p_two_theta, p_intensity, b_two_theta_groups, b_intensity_groups, mean_b_two_theta_groups, mean_b_intensity_groups = Create_Groups(file_name, max_theta, intensity, file_path_pp, file_path_ppp,  group_size) #Normalise data and divide data into groups
    a_min_t, a_max_t, min_i, min_t, max_i, max_t = Peak_Range(b_two_theta, b_intensity, p_two_theta, b_two_theta_groups, b_intensity_groups, mean_b_two_theta_groups, mean_b_intensity_groups) #Identify the two theta and itensity range of each peak
    P_DF = Model_Fitting(file_name, PP_len_label, b_two_theta, b_intensity, p_two_theta, p_intensity, a_min_t, a_max_t, min_i, min_t, max_i, max_t) #Fit a model to the identified peaks
    P_DF1 = Combine_DFs(file_name, P_DF1, P_DF) #Add data in P_DF to P_DF1
    return P_DF1

def Pr_Peak_Picking_Loop(file_path_pr, two_theta_limit, group_size, file_extension, Local_search_threshold, PP_len_label, MtMR_min, MtMR_max, thres_factor): #Control Loop
    file_path_pp = file_path_pr + "/Pattern_Present"
    file_path_ppp = file_path_pr + "/Pattern_Present/Peaks"
    os.chdir(file_path_pr) #Change directory to desired directory
    retval = os.getcwd() #Get current directory
    print ("Predicted_Peak_Picking_Loop: Current working directory %s" % retval) #Show current directory
    filenames = [] #Create list for desired files
    for ex in file_extension: #For each given file extension
        filenames.extend(natsorted(glob.glob(ex))) #Identify files with set file extension
    nf = len(filenames)-1 #Determine number of files (enummerate starts at 0, hence -1)
    P_DF = pd.DataFrame(columns={'file_name', 'Number_of_Peaks'}) #Create new dataframe P_DF, to contain Max to Mean ratios
    for ln, f in enumerate(filenames): #For each filename
        file_name= str(f)  #Make the filename a string
        #print ("File name is %s" % file_name) #Show filename
        intensity = main(file_path_pr, file_name, two_theta_limit, group_size, file_extension, Local_search_threshold, ln) #Send to function main (Remove the baseline from all patterns)
        if np.count_nonzero(intensity) > 0:
            P_DF = P_DF.append({'file_name': file_name[:PP_len_label], 'Number_of_Peaks': str(0)}, ignore_index=True) #Add all values to new row in P_DF
    os.chdir(file_path_pp) #Change directory
    Number_of_peaks_to_excel(P_DF)
    Visualise_Folders(file_path_pp, file_extension, PP_len_label) #Send to function Visualise_Folders (Create a visual image of data)
    retval = os.getcwd() #Get current directory
    print ("Current working directory %s" % retval) #Show current directory
    filenames = [] #Create list for desired files
    for ex in file_extension: #For each given file extension
        filenames.extend(natsorted(glob.glob(ex))) #Identify files with set file extension
    for ln, f in enumerate(filenames): #For each filename
        file_name= str(f)  #Make the filename a string
        #print ("File name is %s" % file_name) #Show filename 
        max_theta, intensity, mean_BI = Read_Data(file_name) #Send to function Read_Data (Normalise all patterns)
        picked_i, P_DF = main_2(file_path_pr, file_name, two_theta_limit, group_size, file_extension, Local_search_threshold, ln, max_theta, intensity, thres_factor, mean_BI, P_DF, PP_len_label) #Send to function main_2 (Pick Peaks)
        Peak_Folder(file_name, max_theta, picked_i) #Create crystalline folder, if more peaks picked than set limit
    Number_of_peaks_to_excel(P_DF) # Send to function Number_of_peaks_to_excel (Write an excel spreadsheet containing file names, Max to Mean Ratios, folder location and number of peaks picked)
    os.chdir(file_path_ppp) #Change directory to desired directory
    Visualise_Folders(file_path_pp, file_extension, PP_len_label) #Send to function Visualise_Folders (Create a visual image of data)
    retval = os.getcwd() #Get current directory
    print ("Current working directory %s" % retval) #Show current directory
    P_DF1 = Create_DF(file_name) #Send to function Create_DF (Create a new dataframe, P_DF1, to hold all of the fitted peak data from all files)
    filenames = [] #Create list for desired files
    for ex in file_extension: #For each given file extension
        filenames.extend(natsorted(glob.glob(ex))) #Identify files with set file extension
    for f in filenames: #For each filename
        file_name= str(f)  #Make the filename a string
        #print ("File name is %s" % file_name) #Show filename
        P_DF1 = main3(file_path_ppp, file_name, file_path_pp,  group_size, PP_len_label, P_DF1) #Send to function main (Fit a model to all picked peaks)
    writer = DF_to_excel(P_DF1) #Send to function DF_to_excel (Save data to an excel spreadsheet)
    print('Finished')
    return