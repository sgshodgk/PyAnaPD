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

def Read_in_DataFrame(file_path_c, file_path_ppp, lamda1): #Read in DataFrames containing extracted peaks (experimental and predicted)
    loc_1 = file_path_c + "/Unique_Fitted_Picked_Peaks.xlsx" #Locate the excel file (experimental data)
    DF1 = pd.read_excel(loc_1) #Read in file as a dataframe, DF1
    loc_2 = file_path_ppp + "/Fitted_Picked_Peaks.xlsx" #Locate the excel file (predicted data)
    P_DF1 = pd.read_excel(loc_2) #Read in file as a dataframe, P_DF1
    pi = math.pi
    cd_s = lamda1/(2*np.sin((DF1['two_theta']/2)*pi/180)) #Convert experimental to d space
    DF1['d'] = cd_s #Add result to DF1
    pd_s = lamda1/(2*np.sin((P_DF1['two_theta']/2)*pi/180)) #Convert predicted to d space
    P_DF1['d'] = pd_s #Add result to P_DF1
    return DF1, P_DF1

def Similarities(DF1, P_DF1, filenames, filenames_pr, EP_len_label, PP_len_label, Sig_peaks, Inital_difference_tol, Difference_tol, writer): #Identifying similarities between experimental and predicted patterns
    e_filenames = DF1['file_name'].unique() #Determine unique filenames in DF1
    p_filenames = P_DF1['file_name'].unique() #Determine unique filenames in DF1
    ppdf = pd.DataFrame(columns=[i[:PP_len_label] for i in p_filenames])#Set all filenames as columns
    ppdf['files']= [i[:EP_len_label] for i in e_filenames] #Set column of filenames as index (rows) #Add a column of filenames
    ppdf = ppdf.set_index('files') #Set column of filenames as index (rows)
    epdf = pd.DataFrame(columns=[i[:PP_len_label] for i in p_filenames])#Set all filenames as columns
    epdf['files']= [i[:EP_len_label] for i in e_filenames] #Set column o filenames as index (rows) #Add a column of filenames
    epdf = epdf.set_index('files') #Set column of filenames as index (rows)
    
    d = dict(tuple(DF1.groupby('file_name'))) #Group data by filenames in DF1
    e = dict(tuple(P_DF1.groupby('file_name'))) #Group data by filenames in DF1
    sdf = pd.DataFrame() #New Dataframe for Matched Peak Postions (Ex vs Pr)
    en = [] #Experimental file name
    pn = [] #Predicted file name
    d_e = [] #Similar experimental peak
    d_p = [] #Similar predicted peak
    i_e = [] #Corresponding experimental intensity 
    i_p = [] #Corresponding predicted intensity 
    shifted = [] #Shifted experimental peak position
    enn = [] #Experimental file name
    pnn = [] #Predicted file name
    pm = [] #Percentage match
    epm = []
    s= [] #Shift relative to experimental peak postion (x axis)
    ps = [] #Percentage shift
    for linei, i in enumerate(filenames): #For each experimental pattern
        for linej, j in enumerate(filenames_pr): #For each predicted pattern
            b = 0 #Number of matched peaks
            a = d[i].sort_values(by=['maxI'], ascending=False)
            for linek, k in enumerate(a['d']):
                c = e[j].sort_values(by=['maxI'], ascending=False)
                for linel, l in enumerate(c['d']):
                    if linek < Sig_peaks and b==0:
                        if linel < Sig_peaks and b==0:
                            x = math.isclose(k, l, abs_tol=(k * Inital_difference_tol)) #If difference is less than set tolerence
                            if x == True and b==0:
                                dif = k-l #Determine difference between peaks
                                cdif = dif/k #Difference relative to experimental peak positions
                                b = b+1 #Increase number of matched peaks by 1
                            if linel & linek == Sig_peaks and b==0:
                                dif = 0
                                cdif = 0
                                b = b+1 #Increase number of matched peaks by 1
            if b >= 1: #For all other peaks
                for linek, k in enumerate(a['d']):
                    for linel, l in enumerate(c['d']):
                        x = math.isclose(k-(k*cdif), l, abs_tol=((k-(k*cdif))*Difference_tol)) #If difference between shifted peak and experimental peak is less than 0.05
                        if x == True:
                            x = DF1.loc[(DF1['file_name'] == i) & (DF1['d'] == k)] #Identify peak row in DF1
                            d_e.append(k) #Collect peak position
                            shifted.append(k-(k*cdif)) #Collect shifted peak position
                            en.append(str(i[:EP_len_label])) #Collect experimental file name
                            i_e.append(x['maxI'].values[0]) #Collect corresponding intensity
                            y = P_DF1[(P_DF1['file_name'] == j) & (P_DF1['d'] == l)] #Identify peak row in P_DF1
                            d_p.append(l) #Collect peak position
                            pn.append(str(j[:PP_len_label])) #Collect predicted file name
                            i_p.append(y['maxI'].values[0]) #Collect corresponding intensity
                            b = b+1 #Increase number of matched peaks by 1
                            pass
            
            pdf = pd.DataFrame({'EP':en, 'PP':pn, 'EP_d':d_e, 'PP_d': d_p, 'SEP_d': shifted, 'EP_Int':i_e, 'PP_Int':i_p}) #Add similar positions to dataframe pdf
            pdf['ED_d'] = pdf['EP_d'] - pdf['PP_d'] #Determine the difference between experimental and predicted peak positions
            pdf['SED_d'] = pdf['SEP_d'] - pdf['PP_d'] #Determine the difference between shifted experimental and predicted peak positions
            pdf['ASED_d'] = pdf['SED_d'].abs() #Absolute difference between shifted experimental and predicted peak positions
            pdf['D_Int'] = pdf['EP_Int'] - pdf['PP_Int'] #Determine the difference between experimental and predicted peak intensities
            pdf = pdf.sort_values('ASED_d', ascending=True).drop_duplicates(['EP_d']) #Sort by abs difference, drop experimental peak positions duplicates
            pdf = pdf.sort_values('ASED_d', ascending=True).drop_duplicates(['PP_d']) #Sort by abs difference, drop predicted peak positions duplicates
            pdf = pdf.sort_values('EP_d', ascending=False).reset_index(drop = True) #Sort by experimental peak positions
            sdf = pd.concat([sdf, pdf]) #Collect matched data 
            del en[:] #Delete experimental file name
            del pn[:] #Delete predicted file nam
            del d_e[:] #Delete experimental peak positions
            del d_p[:] #Delete predicted peak postions
            del i_e[:] #Delete experimental peak intensities
            del i_p[:] #Delete predicted peak intensities
            del shifted[:] #Delete shifted experimental peak positions
            enn.append(str(i[:EP_len_label])) #Collect experimental file name
            pnn.append(str(j[:PP_len_label])) #Collect predicted file name
            pm.append((len(pdf)/len(e[j])*100)) #Collect percentage match
            epm.append((len(pdf)/len(d[i])*100))
            s.append(dif) #Collect shift relative to experimental peak postion (x axis)
            if type(cdif) != str:
                ps.append(cdif*100) #Collect percentage shift
            else:
                ps.append(cdif) #Collect percentage shift
            
            ppdf.loc[str(i[:EP_len_label]), str(j[:PP_len_label])] = (len(pdf)/len(e[j])*100) #Determine % number of matched peaks in file i
            epdf.loc[str(i[:EP_len_label]), str(j[:PP_len_label])] = (len(pdf)/len(d[i])*100) #Determine % number of matched peaks in file i
    
    ppdf.style.applymap(color).to_excel(writer, 'Predicted_Match', index=True) #Write dataframe, mdf, to excel
    epdf.style.applymap(color).to_excel(writer, 'Experimental_Match', index=True) #Write dataframe, mwdf, to excel
    writer.save() #Save excel file
    
    pm = pd.DataFrame({'EP':enn, 'PP':pnn, 'PM':pm, 'EPM':epm, 'S':s, 'PS':ps})
    writer = pd.ExcelWriter('Matched Peak Positions(EP vs. PP).xlsx', engine='openpyxl') #Create an excel spreadsheet
    sdf.to_excel(writer, 'Peak Positions', index=True) #Write dataframe, DF, to excel
    writer.save() #Save excel spreadsheet
    
    writer = pd.ExcelWriter('Percentage Match(EP vs. PP).xlsx', engine='openpyxl') #Create an excel spreadsheet
    pm.to_excel(writer, 'Percentage Match', index=True) #Write dataframe, DF, to excel
    writer.save() #Save excel spreadsheet
    return 

def color(val): #Generate colour map for matched peak positions
    cmap = plt.get_cmap('RdYlGn', 10) #Generate 10 colours from colourmap
    colours = cmap(np.arange(0,cmap.N)) #Fix colours 
    cmap = [rgb2hex(rgb) for rgb in colours] #Convert colours to hex values
    for i, j in zip(range(0, 100, 10), cmap): #For % invervals of 10, up to 100
        if val >= i: #If value >= % interval
            color = j #Select representative colour
    return 'background-color: {}'.format(color)

def ExperimentalvPredicted_Similarities(file_path, file_path_pr, file_extension,  EP_len_label, PP_len_label, lamda1, Sig_peaks, Inital_difference_tol, Difference_tol): #Control Loop
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
    DF1, P_DF1 = Read_in_DataFrame(file_path_c, file_path_ppp, lamda1) #Read in DataFrame, DF1, containing extracted peaks
    Similarities(DF1, P_DF1, filenames, filenames_pr, EP_len_label, PP_len_label, Sig_peaks, Inital_difference_tol, Difference_tol, writer) #Identifying similarities between experimental and predicted patterns
    print('Finished')
    return