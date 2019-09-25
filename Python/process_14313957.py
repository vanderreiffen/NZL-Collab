import os
# website------------------------------------------------------------------------------------
from pkg_resources import working_set
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
### for waiting website loading------
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

# from selenium.common.exceptions import * #optional
import shutil
import datetime
import time
from urllib.request import Request, urlopen, urlretrieve
# update-------------------------------------------------------------------------------------
import win32com.client as xlz
import win32com.client as xlx
import xlwings as xw
import xlwings.constants
# -------------------------------------------------------------------------------------------
import sys
import glob

sys.path.insert(0, str(os.path.abspath(os.path.join('..', 'Config'))))
from getInfoObs import *
from all_function import *
import requests
import re
import collections
requests.packages.urllib3.disable_warnings()



def proceed_14313957(dictfiletorun, obsfullpath, wbmain, source_path, master_path):
    # global dictfiletorun
    methodID = "";
    status = ""
    obs = xlz.gencache.EnsureDispatch("Excel.Application")
    obsBk = obs.Workbooks.Open(str(obsfullpath), ReadOnly=True)
    obs.Visible = True
    obsSHT = None
    freq = str(dictfiletorun["Freq."])
    for list_sht in obsBk.Sheets:
        if str(list_sht.Name).find(freq) > -1:
            obsSHT = obsBk.Worksheets[list_sht.Name]
            break

    # extract information from obsfile to proceed website
    if obsSHT is not None:
        DictheadObs = getDictAllInfo(obsSHT)
        methodID = str(dictfiletorun["SourceMethodID"])
        pub_updt = str(dictfiletorun["SourceMethodID"]) + ":" + str(dictfiletorun["EdgePublication"])
        pub_updt = pub_updt.lower()
        DictFrmObs = dict();
        dictMD = dict();
        dictRoundg = dict()
        DictFrmObs = getDictUpdate(obsSHT, DictheadObs, pub_updt)
        dictMD, dictRoundg = getDictMultDiv(obsSHT, DictheadObs)
        # closed obs
        obsBk.Close(SaveChanges=False)
        if obs.Workbooks.Count == 0:
            obs.Quit()

    # select based on methodID
    if not (dictfiletorun["DownloadedFilePaths"] == 'None'):
        dictfiletorun["SaveFileName"] = dictfiletorun["DownloadedFilePaths"]
        try:
            scdownload, status = update_14313957(str(dictfiletorun["DownloadedFilePaths"]), obsfullpath,
                                            str(dictfiletorun["Freq."]), str(dictfiletorun["QCpath"]),
                                            DictFrmObs[pub_updt][3], dictMD, dictRoundg)
        except Exception as e:
            str(e)
            status = "Failed - <Source File Layout/Content Change>"
    else:
        # attached error
        try:
            scdownload, status = goWebsite_14313957(DictFrmObs, dictfiletorun, source_path, master_path, pub_updt, methodID)
        except Exception as e:
            status = str(e)
            # status = "Failed - <Unable to Download Source File>"
            scdownload = ""
        if scdownload != "" and status == "":
            # strremarks = "Completed (Save source only"
            try:
                scdownload, status = update_14313957(str(scdownload), obsfullpath,
                                                str(dictfiletorun["Freq."]), str(dictfiletorun["QCpath"]),
                                                DictFrmObs[pub_updt][3], dictMD, dictRoundg)
            except:
                status = "Failed - <Source File Layout/Content Change>"

        # here will insert "Completed (Save source only) just incase
        # if scdownload != "" and status == "":
        #     status = "Completed (Save source only)"
        ## end-------

        # replace with the latest source file name
        dictfiletorun["SaveFileName"] = scdownload  # str(os.path.basename(scdownload))

    return status, dictfiletorun


def goWebsite_14313957(DictFrmObs, dictfiletorun, scpath, master_path, publication, smid):

    finscname = ""
    SFilename = ""
    strerror = ""
    bok = False
    # ----------------------------------------------------------------------------------------------
    publication = publication.lower()
    url = str(DictFrmObs[publication][1])
    mID = str(DictFrmObs[publication][0])
    pubname = str(DictFrmObs[publication][2])
    arr_mID = mID.split(":")
    filenm = dictfiletorun["SaveFileName"]
    parent_el = None

    saveName = arr_mID[0].strip() + "_" + filenm + "_" + datetime.datetime.now().strftime("%Y%m%d") \
               + arr_mID[(len(arr_mID) - 1)].strip()

    # check file at source file folder exist or not
    scpathTemp = scpath + "/" + "Temp_" + datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    DestinationFile = scpath + "/" + saveName
    SFilename = scpathTemp + "/" + saveName

    if os.path.exists(scpathTemp):
        shutil.rmtree(scpathTemp)
        os.mkdir(scpathTemp)
    else:
        os.mkdir(scpathTemp)

    if os.path.exists(SFilename):
        os.remove(SFilename)

    if os.path.exists(DestinationFile):
        os.remove(DestinationFile)

    chrome_options = webdriver.ChromeOptions()
    scpathTemp = str(scpathTemp).replace('/', '\\').strip()
    preferences = {"directory_upgrade": True,
                   "download.default_directory": scpathTemp,
                   "plugins.always_open_pdf_externally": True,
                   "download.prompt_for_download": False,
                   "download.directory_upgrade": True,
                   "safebrowsing.enabled": True}
    chrome_options.add_experimental_option("prefs", preferences)

    try:
        driver = webdriver.Chrome(chrome_options=chrome_options, executable_path=master_path + "/chromedriver.exe")
    except:
        strerror = "Failed - <Webdriver Error>"
        driver.close()
        return finscname, strerror

    try:

        driver.get(url)
        driver.implicitly_wait(90)

        action = ActionChains(driver)
        firstLevelMenu = driver.find_element_by_xpath("//ul[@id='menubar-list']/li[2]")
        action.move_to_element(firstLevelMenu).perform()
        secondLevelMenu = driver.find_element_by_xpath("//a[contains(text(),'Excel')]")
        action.move_to_element(secondLevelMenu).perform()
        secondLevelMenu.click()
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//div[@class='ui-dialog ui-widget ui-widget-content ui-corner-all ui-draggable']/div[@id='dialog-modal']/div[@id = 'dialog-content']/iframe")))
        driver.implicitly_wait(90)

        driver.switch_to.frame(driver.find_element_by_xpath("//div[@class='ui-dialog ui-widget ui-widget-content ui-corner-all ui-draggable']/div[@id='dialog-modal']/div[@id = 'dialog-content']/iframe"))
        driver.find_element_by_xpath("//input[@id='btnExportToExcel']").click()
        driver.implicitly_wait(160)
        wait_file_complete(scpathTemp)
        bok = True


        if bok == True:
            time.sleep(3)
            FileList = glob.glob(scpathTemp + "\*")
            if len(FileList) > 0:
                LatestFile = max(FileList, key=os.path.getctime)
                os.rename(LatestFile, SFilename)
                shutil.move(SFilename, DestinationFile)
                time.sleep(3)
                bcomplete = True
        if bcomplete == False:
            strerror = "Failed - <Unable to Download Source File>"

    except Exception as e:
        strerror = str(e)
        strerror = "Failed - <Unable to Download Source File>"
    if not os.path.exists(DestinationFile):
        strerror = "Failed - <Unable to Download Source File>"

    driver.close()
    shutil.rmtree(scpathTemp)
    return DestinationFile, strerror




def update_14313957(scdownloaded, obsfilepath, freq, path_to_save, dictUpdateIDsfrmSc, dictMD, dictRoundg):
    print('start update..')
    # global link_folder
    global dictSCDate
    dictcol = {}
    obs = xlz.gencache.EnsureDispatch("Excel.Application")
    sf = xlx.gencache.EnsureDispatch("Excel.Application")

    sourceFile = ""
    updateStatus = ""
    obsFile = str(obsfilepath)
    sourceFile = str(scdownloaded)

    if sourceFile != "" and obsFile != "":
        scBk = sf.Workbooks.Open(sourceFile, ReadOnly=True)

        filenmstr = obsFile.split("/")[len(obsFile.split("/")) - 1]
        finalobspath = path_to_save + "/" + filenmstr

        if os.path.exists(finalobspath):
            obsBk = obs.Workbooks.Open(finalobspath, ReadOnly=False)
        else:
            obsBk = obs.Workbooks.Open(obsFile, ReadOnly=True)

        obs.Visible = True
        sf.Visible = True
        obsSHT = obsBk.ActiveSheet
        scsht = None

        for list_sht in obsBk.Sheets:
            if str(list_sht.Name).find(freq) > -1:
                obsSHT = obsBk.Worksheets[list_sht.Name]
                break

        bcomplete = False
        filename = scBk.Name
        icount = 0

        try:
            scsht = scBk.Worksheets[1]
        except:
            scsht = None

        if obsSHT is not None and scsht is not None:



            for ik in dictUpdateIDsfrmSc:
                ref_rng = None
                dictSCDate = {}
                res = False
                str_find = str(dictUpdateIDsfrmSc[ik]).strip().lower()
                tbl_data = str_find.split('|')[0].strip().lower()
                to_find = str_find.split('|')[1].strip()

                scsht_find = scBk.Worksheets(tbl_data)
                lrow = scsht_find.Cells.Find(What="*", After=scsht_find.Cells(1, 1),
                                        SearchOrder=xlwings.constants.SearchOrder.xlByRows,
                                        SearchDirection=xlwings.constants.SearchDirection.xlPrevious).Row
                lcol = scsht_find.Cells.Find(What="*", After=scsht_find.Cells(1, 1),
                                        SearchOrder=xlwings.constants.SearchOrder.xlByColumns,
                                        SearchDirection=xlwings.constants.SearchDirection.xlPrevious).Column

                for scsht in scBk.Worksheets:
                    tmpdata = str(scsht.Name).lower().strip()
                    if tmpdata.find(tbl_data) >= 0:
                        try:
                            res = get_data_date(scsht, to_find,lrow,lcol)
                        except Exception as e:
                            str(e)
                        break

                if res and len(dictSCDate) > 0:
                    try:
                        for varDate in dictSCDate:
                            scData = ''
                            scData = (dictSCDate[varDate])
                            mydate = datetime.datetime(varDate.year, varDate.month, 1)
                            if is_number(scData):
                                if len(dictMD) > 0:
                                    if dictMD.get(ik) is not None:
                                        scData = float(scData) / float(dictMD[ik])

                                if len(dictRoundg) > 0:
                                    if dictRoundg.get(ik) is not None:
                                        scData = round(float(scData), int(dictRoundg[ik]))

                            UpdateOneDateOneValueOptimized(scData, ik, mydate, obsSHT, dictcol)
                    except:
                        str(varDate)
                else:
                    rangeObj = obsSHT.Range('A' + str(ik) + ':' + 'M' + str(ik))
                    rangeObj.Interior.ColorIndex = 27
                    ##add counter here
                    icount = int(icount) + 1

            # remove highlight in time points
            lastDateCol = obsSHT.UsedRange.Columns.Count
            rangeObj = obsSHT.Range(
                'N' + str(2) + ':' + convertColSTR(obsSHT, lastDateCol) + str(obsSHT.UsedRange.Rows.Count))
            rangeObj.Interior.ColorIndex = None

            if icount == len(dictUpdateIDsfrmSc):
                updateStatus = "Failed - <Source File Layout/Content Change>"
            else:
                updateStatus = "Completed"
                filenm = os.path.basename(obsFile)
                finalobspath = path_to_save + "/" + filenm

                if os.path.exists(finalobspath):
                    obsBk.Save()
                else:
                    obsBk.SaveAs(finalobspath)
        else:
            if obsSHT is None:  updateStatus = "Failed - <Could not find freq sheet>"

        obsBk.Close(SaveChanges=False)
        scBk.Close(SaveChanges=False)

    if obs.Workbooks.Count == 0:
        obs.Quit()
    if sf.Workbooks.Count == 0:
        sf.Quit()

    return (scdownloaded, updateStatus)




def get_data_date(scsht, tofind,lrow,lcol):
    print('start get_data_date')
    tempdict = {}
    res = False
    col = 1
    print ("lrow = {}, lcol = {}".format(lrow,lcol))
    findcol = find_col(scsht,tofind,lrow,lcol)

    for row in range(1,lrow + 1):
        tempstr = str(scsht.Cells(row,col).Value).strip()
        if tempstr is not None and len(tempstr) == 10 and 'YEJun' in tempstr:
            tempstr = tempstr.replace('YE','')
            tempstr = tempstr.replace(' ','-')
            date = datetime.datetime.strptime(tempstr,'%b-%Y')
            mydate = datetime.datetime(date.year,date.month,1   )
            val = str(scsht.Cells(row,findcol)).strip()
            val = val.replace('..','')
            print('date = {}, val = {}'.format(mydate,val))
            dictSCDate.update({date:val})
            res = True


    print('=======================================================')
    return res


def find_col(scsht,strFind,lrow,lcol):
    print('Finding Column')
    rngFind = scsht.Range(scsht.Cells(1, 1), scsht.Cells(lrow, lcol)).Find(
        What=strFind.strip(), LookAt=xlwings.constants.LookAt.xlWhole,
        LookIn=xlwings.constants.FindLookIn.xlValues,
        SearchOrder=xlwings.constants.SearchOrder.xlByRows, MatchCase=False)
    if rngFind is None:
        rngFind = scsht.Range(scsht.Cells(1, 1), scsht.Cells(lrow, lcol)).Find(
            What=strFind.strip(), LookAt=xlwings.constants.LookAt.xlPart,
            LookIn=xlwings.constants.FindLookIn.xlValues,
            SearchOrder=xlwings.constants.SearchOrder.xlByRows, MatchCase=False)
    print ('{} is in column : {}'.format(strFind,rngFind.Column))
    return rngFind.Column



def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False


def is_date(iyr, im):
    try:
        if (int(iyr) >= 1900 and int(iyr) <= 2500) and (int(im) >= 1 and int(im) <= 12):
            return True
        else:
            return False
    except ValueError:
        pass
