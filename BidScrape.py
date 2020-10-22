import threading
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.chrome.options import Options
from time import sleep
import credentials
import glob
import os
import datetime

# ISF
wallsISF = 'https://app.isqft.com/services/go/advancedsearch/Jr1I2M4hco'
fiberglassISF = 'https://app.isqft.com/services/go/advancedsearch/9iARkDeblv'
louversISF = 'https://app.isqft.com/services/go/advancedsearch/_pS4oqxmFN'
polycarbonateISF = 'https://app.isqft.com/services/go/advancedsearch/-cHZf7hZhX'
skylightsISF = 'https://app.isqft.com/services/go/advancedsearch/I5cXOp94JG'
tubesISF = 'https://app.isqft.com/services/go/advancedsearch/LQjEQkNwuI'
sleep5 = 5
sleep2 = 2
sleep1 = 1

# Insight CMD
wallsCMD = 'CV - Folding Partitions'
fiberglassCMD = 'CV - Fiberglass Sandwich Panels 2018'
louversCMD = 'CV - Louvers 2018'
polycarbonateCMD = 'CV - Polycarbonate Translucent 2018'
skylightsCMD = 'CV - Skylights 2018'
tubesCMD = 'CV - Tube Skylights 2018'

# Virtual Builders Exchange
foldingpartitionsVBX = 'Folding Partitions'
fiberglassVBX = 'Translucent Panels'
louversVBX = 'Louvers'
polycarbonateVBX = 'Polycarbonate Misc'
skylightsVBX = 'Skylights'
tubesVBX = 'Tubular Skylights'

# Dodge Pipeline
wallsDPL = 'CV - Operable Walls'
fiberglassDPL = 'CV - Translucent Panels'
louversDPL = 'CV - Louvers'
polyDPL = 'CV - Polycarbonate'
skylightsDPL = 'CV - Skylights'
tubesDPL = 'CV - Tubular Skylights'

def ISF():
    driverISF = webdriver.Chrome()
    driverISF.get("https://app.isqft.com/login/index.html?returnurl=/")
    sleep(sleep1)
    usernameISF = driverISF.find_element_by_id("username")
    usernameISF.clear()
    usernameISF.send_keys(credentials.unISF)
    passwordISF = driverISF.find_element_by_id("password")
    passwordISF.clear()
    passwordISF.send_keys(credentials.pwISF)
    passwordISF.send_keys(Keys.RETURN)
    sleep(sleep1)

# Folding Partitions
    driverISF.get(wallsISF)
    sleep(sleep5)
    driverISF.find_element_by_id("ext-gen959").click()
    sleep(sleep1)
    driverISF.find_element_by_link_text(
        "Export List - All Projects to - Excel").click()
# Fiberglass    
    driverISF.get(fiberglassISF)
    sleep(sleep5)
    driverISF.find_element_by_id("ext-gen959").click()
    sleep(sleep1)
    driverISF.find_element_by_link_text(
        "Export List - All Projects to - Excel").click()
# Louvers   
    driverISF.get(louversISF)
    sleep(sleep5)
    driverISF.find_element_by_id("ext-gen959").click()
    sleep(sleep1)
    driverISF.find_element_by_link_text(
        "Export List - All Projects to - Excel").click()
# Polycarbonate   
    driverISF.get(polycarbonateISF)
    sleep(sleep5)
    driverISF.find_element_by_id("ext-gen959").click()
    sleep(sleep1)
    driverISF.find_element_by_link_text(
        "Export List - All Projects to - Excel").click()
# Skylights 
    driverISF.get(skylightsISF)
    sleep(sleep5)
    driverISF.find_element_by_id("ext-gen959").click()
    sleep(sleep1)
    driverISF.find_element_by_link_text(
        "Export List - All Projects to - Excel").click()
# Tubes    
    driverISF.get(tubesISF)
    sleep(sleep5)
    driverISF.find_element_by_id("ext-gen959").click()
    sleep(sleep1)
    driverISF.find_element_by_link_text(
        "Export List - All Projects to - Excel").click()
    driverISF.close()

def ISFRename():
    Current_Date = datetime.datetime.today().strftime ('%d-%b-%Y')
    Morpheus_ISqFt = r'C:\Users\clint\Desktop\Morpheus\ISqFt'
    if not os.path.exists(Morpheus_ISqFt):
        os.makedirs(Morpheus_ISqFt)
    S1 = r'C:\Users\clint\Downloads\Advanced_Search.xlsx'
    S2 = r'C:\Users\clint\Downloads\Advanced_Search (1).xlsx'
    S3 = r'C:\Users\clint\Downloads\Advanced_Search (2).xlsx'
    S4 = r'C:\Users\clint\Downloads\Advanced_Search (3).xlsx'
    S5 = r'C:\Users\clint\Downloads\Advanced_Search (4).xlsx'
    S6 = r'C:\Users\clint\Downloads\Advanced_Search (5).xlsx'

    S1_Current = r'C:\Users\clint\Desktop\Morpheus\ISqFt\Folding Partitions_ISF_' + str(Current_Date) + '.xlsx'
    S2_Current = r'C:\Users\clint\Desktop\Morpheus\ISqFt\Fiberglass_ISF_' + str(Current_Date) + '.xlsx'
    S3_Current = r'C:\Users\clint\Desktop\Morpheus\ISqFt\Louvers_ISF_' + str(Current_Date) + '.xlsx'
    S4_Current = r'C:\Users\clint\Desktop\Morpheus\ISqFt\Polycarbonate_ISF_' + str(Current_Date) + '.xlsx'
    S5_Current = r'C:\Users\clint\Desktop\Morpheus\ISqFt\Skylights_ISF_' + str(Current_Date) + '.xlsx'
    S6_Current = r'C:\Users\clint\Desktop\Morpheus\ISqFt\Tubes_ISF_' + str(Current_Date) + '.xlsx'
    
    if os.path.exists(S1_Current):
        os.remove(S1)
    else: 
        os.rename(S1, S1_Current)
    if os.path.exists(S2_Current):
        os.remove(S2)
    else: 
        os.rename(S2, S2_Current)
    if os.path.exists(S3_Current):
        os.remove(S3)
    else: 
        os.rename(S3, S3_Current)
    if os.path.exists(S4_Current):
        os.remove(S4)
    else: 
        os.rename(S4, S4_Current)
    if os.path.exists(S5_Current):
        os.remove(S5)
    else: 
        os.rename(S5, S5_Current)
    if os.path.exists(S6_Current):
        os.remove(S6)
    else: 
        os.rename(S6, S6_Current)
    

def CMD():
    driverCMD = webdriver.Chrome()
    driverCMD.get('https://login.cmdgroup.com/Account/Login')
    sleep(sleep2)
    usernameCMD = driverCMD.find_element_by_id('txtUserName')
    passwordCMD = driverCMD.find_element_by_id('txtPassword')
    usernameCMD.clear()
    usernameCMD.send_keys(credentials.unCMD)
    passwordCMD.clear()
    passwordCMD.send_keys(credentials.pwCMD)
    passwordCMD.send_keys(Keys.RETURN)
    sleep(sleep5)
    driverCMD.find_element_by_id('btnReset').click()
    sleep(sleep1)
    
    driverCMD.find_element_by_id('drpSearch-trigger-picker').click()
    driverCMD.find_element_by_id(wallsCMD).click()
    sleep(sleep2)
    driverCMD.find_element_by_id('btnOpenExportPopup-btnInnerEl').click()
    sleep(sleep1)
    driverCMD.find_element_by_id('btnExportPopup-btnInnerEl').click()
    sleep(sleep5)
    list_of_files = glob.glob(r'C:\Users\clint\Downloads\Project List*.xls')
    latest_file = max(list_of_files, key=os.path.getctime)
    S1 = str(latest_file)
    sleep(sleep1)

    

    driverCMD.find_element_by_id('drpSearch-trigger-picker').click()
    driverCMD.find_element_by_id(fiberglassCMD).click()
    sleep(sleep2)
    driverCMD.find_element_by_id('btnOpenExportPopup-btnInnerEl').click()
    sleep(sleep1)
    driverCMD.find_element_by_id('btnExportPopup-btnInnerEl').click()
    sleep(sleep5)
    list_of_files = glob.glob(r'C:\Users\clint\Downloads\Project List*.xls')
    latest_file = max(list_of_files, key=os.path.getctime)
    S2 = str(latest_file)
    sleep(sleep1)

    

    driverCMD.find_element_by_id('drpSearch-trigger-picker').click()
    driverCMD.find_element_by_id(louversCMD).click()
    sleep(sleep2)
    driverCMD.find_element_by_id('btnOpenExportPopup-btnInnerEl').click()
    sleep(sleep1)
    driverCMD.find_element_by_id('btnExportPopup-btnInnerEl').click()
    sleep(sleep5)
    list_of_files = glob.glob(r'C:\Users\clint\Downloads\Project List*.xls')
    latest_file = max(list_of_files, key=os.path.getctime)
    S3 = str(latest_file)
    sleep(sleep1)

    

    driverCMD.find_element_by_id('drpSearch-trigger-picker').click()
    driverCMD.find_element_by_id(polycarbonateCMD).click()
    sleep(sleep2)
    driverCMD.find_element_by_id('btnOpenExportPopup-btnInnerEl').click()
    sleep(sleep1)
    driverCMD.find_element_by_id('btnExportPopup-btnInnerEl').click()
    sleep(sleep5)
    list_of_files = glob.glob(r'C:\Users\clint\Downloads\Project List*.xls')
    latest_file = max(list_of_files, key=os.path.getctime)
    S4 = str(latest_file)
    sleep(sleep1)

    

    driverCMD.find_element_by_id('drpSearch-trigger-picker').click()
    driverCMD.find_element_by_id(skylightsCMD).click()
    sleep(sleep2)
    driverCMD.find_element_by_id('btnOpenExportPopup-btnInnerEl').click()
    sleep(sleep1)
    driverCMD.find_element_by_id('btnExportPopup-btnInnerEl').click()
    sleep(sleep5)
    list_of_files = glob.glob(r'C:\Users\clint\Downloads\Project List*.xls')
    latest_file = max(list_of_files, key=os.path.getctime)
    S5 = str(latest_file)
    sleep(sleep1)

    

    driverCMD.find_element_by_id('drpSearch-trigger-picker').click()
    driverCMD.find_element_by_id(tubesCMD).click()
    sleep(sleep2)
    driverCMD.find_element_by_id('btnOpenExportPopup-btnInnerEl').click()
    sleep(sleep1)
    driverCMD.find_element_by_id('btnExportPopup-btnInnerEl').click()
    sleep(sleep5)
    list_of_files = glob.glob(r'C:\Users\clint\Downloads\Project List*.xls')
    latest_file = max(list_of_files, key=os.path.getctime)
    S6 = str(latest_file)
    sleep(sleep1)
    driverCMD.close()
    
    
    Current_Date = datetime.datetime.today().strftime ('%d-%b-%Y')
    Morpheus_InsightCMD = r'C:\Users\clint\Desktop\Morpheus\InsightCMD'
    if not os.path.exists(Morpheus_InsightCMD):
        os.makedirs(Morpheus_InsightCMD)
    
    S1_Current = r'C:\Users\clint\Desktop\Morpheus\InsightCMD\Folding Partitions_CMD_' + str(Current_Date) + '.xls'
    S2_Current = r'C:\Users\clint\Desktop\Morpheus\InsightCMD\Fiberglass_CMD_' + str(Current_Date) + '.xls'
    S3_Current = r'C:\Users\clint\Desktop\Morpheus\InsightCMD\Louvers_CMD_' + str(Current_Date) + '.xls'
    S4_Current = r'C:\Users\clint\Desktop\Morpheus\InsightCMD\Polycarbonate_CMD_' + str(Current_Date) + '.xls'
    S5_Current = r'C:\Users\clint\Desktop\Morpheus\InsightCMD\Skylights_CMD_' + str(Current_Date) + '.xls'
    S6_Current = r'C:\Users\clint\Desktop\Morpheus\InsightCMD\Tubes_CMD_' + str(Current_Date) + '.xls'
    
    if os.path.exists(S1_Current):
        os.remove(S1)
    else: 
        os.rename(S1, S1_Current)
    if os.path.exists(S2_Current):
        os.remove(S2)
    else: 
        os.rename(S2, S2_Current)
    if os.path.exists(S3_Current):
        os.remove(S3)
    else: 
        os.rename(S3, S3_Current)
    if os.path.exists(S4_Current):
        os.remove(S4)
    else: 
        os.rename(S4, S4_Current)
    if os.path.exists(S5_Current):
        os.remove(S5)
    else: 
        os.rename(S5, S5_Current)
    if os.path.exists(S6_Current):
        os.remove(S6)
    else: 
        os.rename(S6, S6_Current)

    
def DPL():
    Current_Date = datetime.datetime.today().strftime ('%d-%b-%Y')
    Morpheus_DodgePL = r'C:\Users\clint\Desktop\Morpheus\DodgePL'
    if not os.path.exists(Morpheus_DodgePL):
        os.makedirs(Morpheus_DodgePL)
    S1_Current = r'C:\Users\clint\Desktop\Morpheus\DodgePL\Folding Partitions_DPL_' + str(Current_Date) + '.xlsx'
    S2_Current = r'C:\Users\clint\Desktop\Morpheus\DodgePL\Fiberglass_DPL_' + str(Current_Date) + '.xlsx'
    S3_Current = r'C:\Users\clint\Desktop\Morpheus\DodgePL\Louvers_DPL_' + str(Current_Date) + '.xlsx'
    S4_Current = r'C:\Users\clint\Desktop\Morpheus\DodgePL\Polycarbonate_DPL_' + str(Current_Date) + '.xlsx'
    S5_Current = r'C:\Users\clint\Desktop\Morpheus\DodgePL\Skylights_DPL_' + str(Current_Date) + '.xlsx'
    S6_Current = r'C:\Users\clint\Desktop\Morpheus\DodgePL\Tubes_DPL_' + str(Current_Date) + '.xlsx'
    
    driverDPL = webdriver.Chrome()
    driverDPL.get('https://connect.dodgepipeline.com')
    sleep(2)
    usernameDPL = driverDPL.find_element_by_id('UserName')
    usernameDPL.clear()
    usernameDPL.send_keys(credentials.unDPL)
    passwordDPL = driverDPL.find_element_by_id('Password')
    passwordDPL.clear()
    passwordDPL.send_keys(credentials.pwDPL)
    driverDPL.find_element_by_id('LoginButton').click()
    sleep(2)
    

    driverDPL.find_element_by_link_text(wallsDPL).click()
    sleep(3)
    if driverDPL.find_elements_by_id('excelExport'):
        driverDPL.find_element_by_id('excelExport').click()
    sleep(10)
    if driverDPL.find_elements_by_link_text('Download file'):
        driverDPL.find_element_by_link_text('Download file').click()
        sleep(2)
        list_of_files = glob.glob(r'C:\Users\clint\Downloads\*_ProjectList_*.xlsx')
        latest_file = max(list_of_files, key=os.path.getctime)
        S1 = str(latest_file)
        if os.path.exists(S1_Current):
            os.remove(S1)
        else: 
            os.rename(S1, S1_Current)
    else:
        sleep(2)     
    

    driverDPL.find_element_by_link_text('Saved Searches').click()
    sleep(3)
    driverDPL.find_element_by_link_text(fiberglassDPL).click()
    sleep(2)
    if driverDPL.find_elements_by_id('excelExport'):
        driverDPL.find_element_by_id('excelExport').click()
    sleep(10)
    if driverDPL.find_elements_by_link_text('Download file'):
        driverDPL.find_element_by_link_text('Download file').click()
        sleep(2)
        list_of_files = glob.glob(r'C:\Users\clint\Downloads\*_ProjectList_*.xlsx')
        latest_file = max(list_of_files, key=os.path.getctime)
        S2 = str(latest_file)
        if os.path.exists(S2_Current):
            os.remove(S2)
        else: 
            os.rename(S2, S2_Current)
    else:
        sleep(2)
    

    driverDPL.find_element_by_link_text('Saved Searches').click()
    sleep(3)
    driverDPL.find_element_by_link_text(louversDPL).click()
    sleep(2)
    if driverDPL.find_elements_by_id('excelExport'):
        driverDPL.find_element_by_id('excelExport').click()
    sleep(10)
    if driverDPL.find_elements_by_link_text('Download file'):
        driverDPL.find_element_by_link_text('Download file').click()
        sleep(2)
        list_of_files = glob.glob(r'C:\Users\clint\Downloads\*_ProjectList_*.xlsx')
        latest_file = max(list_of_files, key=os.path.getctime)
        S3 = str(latest_file)
        if os.path.exists(S3_Current):
            os.remove(S3)
        else: 
            os.rename(S3, S3_Current)
    else:
        sleep(2)
    

    driverDPL.find_element_by_link_text('Saved Searches').click()
    sleep(3)
    driverDPL.find_element_by_link_text(polyDPL).click()
    sleep(2)
    if driverDPL.find_elements_by_id('excelExport'):
        driverDPL.find_element_by_id('excelExport').click()
    sleep(10)
    if driverDPL.find_elements_by_link_text('Download file'):
        driverDPL.find_element_by_link_text('Download file').click()
        sleep(2)
        list_of_files = glob.glob(r'C:\Users\clint\Downloads\*_ProjectList_*.xlsx')
        latest_file = max(list_of_files, key=os.path.getctime)
        S4 = str(latest_file)
        if os.path.exists(S4_Current):
            os.remove(S4)
        else: 
            os.rename(S4, S4_Current)
    else:
        sleep(2)

    

    driverDPL.find_element_by_link_text('Saved Searches').click()
    sleep(3)
    driverDPL.find_element_by_link_text(skylightsDPL).click()
    sleep(2)
    if driverDPL.find_elements_by_id('excelExport'):
        driverDPL.find_element_by_id('excelExport').click()
    sleep(10)
    if driverDPL.find_elements_by_link_text('Download file'):
        driverDPL.find_element_by_link_text('Download file').click()
        sleep(2)
        list_of_files = glob.glob(r'C:\Users\clint\Downloads\*_ProjectList_*.xlsx')
        latest_file = max(list_of_files, key=os.path.getctime)
        S5 = str(latest_file)
        if os.path.exists(S5_Current):
            os.remove(S5)
        else: 
            os.rename(S5, S5_Current)
    else:
        sleep(2)
    

    driverDPL.find_element_by_link_text('Saved Searches').click()
    sleep(3)
    driverDPL.find_element_by_link_text(tubesDPL).click()
    sleep(2)
    if driverDPL.find_elements_by_id('excelExport'):
        driverDPL.find_element_by_id('excelExport').click()
    sleep(10)
    if driverDPL.find_elements_by_link_text('Download file'):
        driverDPL.find_element_by_link_text('Download file').click()
        sleep(2)
        list_of_files = glob.glob(r'C:\Users\clint\Downloads\*_ProjectList_*.xlsx')
        latest_file = max(list_of_files, key=os.path.getctime)
        S6 = str(latest_file)
        if os.path.exists(S6_Current):
            os.remove(S6)
        else: 
            os.rename(S6, S6_Current)
    else:
        driverDPL.close()




ISF()
ISFRename()
CMD()
DPL()






# search.ISF(wallsISF)


# t1 = threading.Thread(target=search.ISF, args=(1))
# t2 = threading.Thread(target=search.DPL, args=(1))

# t1.start()
# t2.start()


# filters.DPLsearch(louversDPL)


# z = a.get()
# print(z)
# app = App(root)
# app.mainloop()
