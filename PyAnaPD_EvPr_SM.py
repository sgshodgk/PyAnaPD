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

def Read_in_DataFrame(file_path_ppp): #Read in DataFrame, DF1, containing extracted peaks
    sdf = pd.read_excel(file_path_ppp + "/Matched Peak Positions(EP vs. PP).xlsx") #Locate the desired file
    sdf = sdf.set_index(sdf.columns[0]) #Set first column (file names) as index
    pm = pd.read_excel(file_path_ppp + "/Percentage Match(EP vs. PP).xlsx") #Locate the desired file
    pm = pm.set_index(pm.columns[0]) #Set first column (file names) as index
    return sdf, pm

def Visualise_Similarities(sdf, pm, EvP_Match_Pr, EvP_Match_Ex, Intensity_difference_tol, I_Match, Visualisations_SM): #Visualise similaries between predicted and experimental data
    retval = os.getcwd()
    newpath = retval + '/EvPr/Visualisations' #Create 'visualisations' folder
    if Visualisations_SM == 'Yes':
        if not os.path.exists(newpath): #If folder does not exist
            os.makedirs(newpath) #Make folder
        os.chdir(newpath) 
    cs = plt.cm.nipy_spectral(np.linspace(0,1,len(pm['PP'].unique())+1)) #Generate a colourmap
    
    if Visualisations_SM == 'Yes':
        pm.groupby(['EP', 'PP']).sum().unstack().plot(y='PM', kind='bar', stacked=False, color=cs, fontsize=5).legend(title = 'Predicted Patterns', loc='center left',bbox_to_anchor=(1.0, 0.5))
        plt.title('Percentage Number of Peaks within Predicted Patterns observed in Experimental Patterns')
        plt.xlabel('Collected Patterns')
        plt.ylabel('Percentage Match (%)')
        plt.savefig('Percentage Number of Peaks within Predicted Patterns observed in Experimental Patterns.pdf',dpi=100, bbox_inches = "tight")
        plt.clf()
        plt.close()
    
    if Visualisations_SM == 'Yes':
        pm.groupby(['EP', 'PP']).sum().unstack().plot(y='EPM', kind='bar', stacked=False, color=cs, fontsize=5).legend(title = 'Predicted Patterns', loc='center left',bbox_to_anchor=(1.0, 0.5))
        plt.title('Percentage Number of Peaks within Experimental Patterns observed in Predicted Patterns')
        plt.xlabel('Collected Patterns')
        plt.ylabel('Percentage Match (%)')
        plt.savefig('Percentage Number of Peaks within Experimental Patterns observed in Predicted Patterns.pdf',dpi=100, bbox_inches = "tight")
        plt.clf()
        plt.close()
    
    spm = pm[pm.PM >= EvP_Match_Pr]
    if Visualisations_SM == 'Yes':
        spm.groupby(['EP', 'PP']).sum().unstack().plot(y='PM', kind='bar', stacked=False, color=cs).legend(title = 'Predicted Patterns', loc='center left',bbox_to_anchor=(1.0, 0.5))
        plt.title('Percentage Number of Peaks within Predicted Patterns observed in Experimental Patterns (Equal to or Greater than ' + str(EvP_Match_Pr) + ').pdf')
        plt.xlabel('Collected Patterns')
        plt.ylabel('Percentage Match (%)')
        plt.savefig('Percentage Number of Peaks within Predicted Patterns observed in Experimental Patterns (Equal to or Greater than ' + str(EvP_Match_Pr) + ').pdf',dpi=100, bbox_inches = "tight")
        plt.clf()
        plt.close()
    
    spm = spm[spm.EPM >= EvP_Match_Ex]
    if Visualisations_SM == 'Yes':
        spm.groupby(['EP', 'PP']).sum().unstack().plot(y='EPM', kind='bar', stacked=False, color=cs).legend(title = 'Predicted Patterns', loc='center left',bbox_to_anchor=(1.0, 0.5))
        plt.title('Percentage Number of Peaks within Experimental Patterns observed in Predicted Patterns (Equal to or Greater than ' + str(EvP_Match_Ex) + ').pdf')
        plt.xlabel('Collected Patterns')
        plt.ylabel('Percentage Match (%)')
        plt.savefig('Percentage Number of Peaks within Predicted Patterns observed in Experimental Patterns (Equal to or Greater than ' + str(EvP_Match_Ex) + ').pdf',dpi=100, bbox_inches = "tight")
        plt.clf()
        plt.close()
    
    xdf = pd.merge(spm, sdf, on=['EP','PP'])
    scdf = pd.DataFrame({'count' : xdf.groupby(['EP', 'PP']).size()}).reset_index() #Determine number of matched peaks
    cdf = xdf[abs(xdf.D_Int) <= Intensity_difference_tol].groupby(['EP', 'PP']).size().reset_index()
    df = pd.merge(scdf, cdf, on=['EP','PP'])
    df['count_IT/count (%)']=((df.iloc[:,-1]/df.iloc[:,-2])*100) #Determine % number of peaks that match within Ex and Pr

    df = df[df['count_IT/count (%)'] >= I_Match] #Determine Ex where intensities match over a set threshold (%) to Pr
    if Visualisations_SM == 'Yes':
        df.groupby(['EP', 'PP']).sum().unstack().plot(y='count_IT/count (%)', kind='bar', stacked=False, color=cs).legend(title = 'Predicted Patterns', loc='center left',bbox_to_anchor=(1.0, 0.5))
        plt.title('Percentage Intensity Match')
        plt.xlabel('Collected Patterns')
        plt.ylabel('Percentage Intensity Match (%)')
        plt.savefig('Percentage Intensity Match.pdf',dpi=100, bbox_inches = "tight")
        plt.clf()
        plt.close()
        
    xxdf = pd.merge(spm, df, on=['EP','PP'])
    if Visualisations_SM == 'Yes':
        xxdf.groupby(['EP', 'PP']).sum().unstack().plot(y='PS', kind='bar', stacked=False, color=cs).legend(title = 'Predicted Patterns', loc='center left',bbox_to_anchor=(1.0, 0.5))
        plt.title('Percentage Shift of Peaks')
        plt.xlabel('Collected Patterns')
        plt.ylabel('Percentage Shift (%)')
        plt.savefig('Percentage Shift of Peaks.pdf',dpi=100, bbox_inches = "tight") #Save subplots
        plt.clf()
        plt.close()
        
    os.chdir(retval) 
    spm = spm.merge(df.iloc[:, :2].assign(a='key'),how='left').dropna() #Drop any data which does not meet the intensities requirement
    writer = pd.ExcelWriter('Selected_Matching.xlsx', engine='openpyxl') #Create an excel spreadsheet
    spm.to_excel(writer, 'Selected_Matching', index=True) #Write dataframe, DF, to excel
    df.to_excel(writer, 'Intensities', index=True) #Write dataframe, DF, to excel
    writer.save() #Save excel spreadsheet
    
    return spm, newpath

def Visualise_shifts(file_path_c, file_path_ppp, filenames, filenames_pr, EP_len_label, PP_len_label, spm, lamda1, Max_d, newpath): #Visualise peak shifts
    pi = math.pi
    loc_3 = file_path_ppp + "/Matched Peak Positions(EP vs. PP).xlsx" #Locate the excel file (Matched Peak Positions)
    MF1 = pd.read_excel(loc_3) #Read in file as a dataframe, MF1
    
    for ln, (cp, pp) in enumerate(zip(spm['EP'], spm['PP'])): #For matched experimental and predicted patterns (match >= EvP_Match_Pr)
        cp_files = [i for i in os.listdir(file_path_c) if os.path.isfile(os.path.join(file_path_c,i)) and str(cp) in i[:EP_len_label]] #Identify full experimental file name
        file_name = str(cp_files[0]) #full experimental file name
        os.chdir(file_path_c)
        two_theta, intensity = Read_Data(file_name) #Read in picked peaks file
        cpt_d= lamda1/(2*np.sin((two_theta/2)*pi/180)) #Convert to d space
        mask = cpt_d < Max_d #Indentify d space below Max_d (mask)
        cpt_d = cpt_d[mask] #Cut d space above Max_d
        e_intensity = intensity[mask] #Cut intensities above mask
        
        ff = pd.DataFrame({'d_space':cpt_d, 'intensity':e_intensity}) #Collect masked data in a Dataframe
        se = ff[ff['intensity'] > 0].copy() #Identify intensities above 0
        ff['intensity']= 0 #Set all intensties as 0
        a = spm.loc[spm.index[ln], 'PS'] #Identify calculated shift
        se['d_space'] = se['d_space'] - (se['d_space']*(a/100)) #Shift experimental data
        result = pd.concat([ff, se]) #Combine shited pattern with orginal pattern framework
        res = result.sort_values('d_space', ascending=True).reset_index() #Sort values with respect to d space
        
        pp_files = [i for i in os.listdir(file_path_ppp) if os.path.isfile(os.path.join(file_path_ppp,i)) and str(pp) in i[:PP_len_label]] #Identify full predicted file name
        file_name = str(pp_files[0]) #full experimental file name
        os.chdir(file_path_ppp)
        two_theta, intensity = Read_Data(file_name) #Read in picked peaks file
        ppt_d= lamda1/(2*np.sin((two_theta/2)*pi/180)) #Convert to d space
        mask = ppt_d < Max_d #Indentify d space below Max_d (mask)
        ppt_d = ppt_d[mask] #Cut d space above Max_d
        pr_intensity = intensity[mask] #Cut intensities above mask      

        smm = MF1[(MF1['EP'] == (cp)) & (MF1['PP'] == (pp))] #Identify matched data corresponding to experimental and predicted pattern
        sm = smm.iloc[:, 3:].copy()
        sm = sm[sm < Max_d]
        
        os.chdir(newpath)
        f, (ax1, ax2, ax3) = plt.subplots(3, 1, sharex=True) #Generate 3 plots
        ax1.plot(cpt_d, e_intensity, color='r', label=str(cp)) #Plot peak picked experimental data
        ax1.scatter(sm['EP_d'], sm['EP_Int'], color = 'k', s = 8) #Highlight matched peaks
        ax1.legend(loc="upper right")
        ax2.plot(res['d_space'], res['intensity'], color='g', label=str(cp) + ' Shifted') #Plot shifted peak picked experimental data
        ax2.scatter(sm['SEP_d'], sm['EP_Int'], color = 'k', s = 8) #Highlight matched peaks
        ax2.legend(loc="upper right")
        ax3.plot(ppt_d, pr_intensity, color='b', label = str(pp)) #Plot peak picked predicted data
        ax3.scatter(sm['PP_d'], sm['PP_Int'], color = 'k', s = 8) #Highlight matched peaks
        ax3.legend(loc="upper right")
        f.text(0.5, 0.04, 'd (Å)', ha='center', va='center')
        f.text(0.06, 0.5, 'Normalised Intensity (a.u.)', ha='center', va='center', rotation='vertical')
        plt.savefig('Shifted ' + str(cp) + ' against ' +  str(pp)  + ".pdf") #Save subplots
        plt.clf()
        plt.close()
    return

def Read_Data(file_name): #Read in data
    data = np.loadtxt(file_name, usecols = (0,1))  #Load in data from file, skipping the rows containing a character or &
    two_theta = data[:, 0] #Set first column as two theta
    intensity = data[:, 1] #Set second column as intensity
    intensity = intensity/intensity.max() #Normalise intensity
    np.seterr(divide='ignore', invalid='ignore') #Ignore floating point errors
    return two_theta, intensity
    
def ExperimentalvPredicted_Selective_Matching(file_path, file_path_pr, file_extension,  EP_len_label, PP_len_label, lamda1, Intensity_difference_tol, I_Match, EvP_Match_Pr, EvP_Match_Ex, Max_d, Visualisations_SM): #Control Loop
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
    sdf, pm = Read_in_DataFrame(file_path_ppp) #Read in DataFrame, DF1, containing extracted peaks
    spm, newpath= Visualise_Similarities(sdf, pm, EvP_Match_Pr, EvP_Match_Ex, Intensity_difference_tol, I_Match, Visualisations_SM) #Visualise similaries between predicted and experimental data
    if Visualisations_SM == 'Yes':
        Visualise_shifts(file_path_c, file_path_ppp, filenames, filenames_pr, EP_len_label, PP_len_label, spm, lamda1, Max_d, newpath) #Visualise peak shifts
    print('Finished')
    return