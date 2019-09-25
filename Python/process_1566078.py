import os
import  re
# website------------------------------------------------------------------------------------
from pkg_resources import working_set
from selenium import webdriver
# from selenium.common.exceptions import * #optional
import shutil
import datetime
from datetime import timedelta
import time
from urllib.request import Request, urlopen, urlretrieve
# update-------------------------------------------------------------------------------------
import win32com.client as xlz
import win32com.client as xlx
import xlwings.constants
# -------------------------------------------------------------------------------------------
import sys
import glob

sys.path.insert(0, str(os.path.abspath(os.path.join('..', 'Config'))))
from getInfoObs import *
from all_function import *
from pdfconvert import *
import requests

requests.packages.urllib3.disable_warnings()

def proceed_1566078(dictfiletorun, obsfullpath, wbmain, source_path, master_path):
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
            scdownload, status = update_1566078_10001(str(dictfiletorun["DownloadedFilePaths"]), obsfullpath,
                                            str(dictfiletorun["Freq."]), str(dictfiletorun["QCpath"]),
                                            DictFrmObs[pub_updt][3], dictMD, dictRoundg)
        except Exception as e:
            str(e)
            status = "Failed - <Source File Layout/Content Change>"
        dictfiletorun["SaveFileName"] = scdownload
    else:
        # attached error
        try:
            scdownload = ""
            scdownload, status = goWebsite_1566078_10001(DictFrmObs, dictfiletorun, source_path, master_path,
                                                              pub_updt, methodID)
        except Exception as e:
            status = str(e)
            # status = "Failed - <Unable to Download Source File>"
            scdownload = ""
        if scdownload != "" and status == "":
            # strremarks = "Completed (Save source only"
            try:
                scdownload, status = update_1566078_10001(str(scdownload), obsfullpath,
                                                 str(dictfiletorun["Freq."]), str(dictfiletorun["QCpath"]),
                                                DictFrmObs[pub_updt][3], dictMD, dictRoundg)
            except:
                status = "Failed - <Source File Layout/Content Change>"

        dictfiletorun["SaveFileName"] = scdownload  # str(os.path.basename(scdownload))
    return status, dictfiletorun


def goWebsite_1566078_10001(DictFrmObs, dictfiletorun, scpath, master_path, publication, smid):
    finscname = "";
    strerror = ""
    res = False
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

    if os.path.exists(SFilename):
        os.remove(SFilename)

    if os.path.exists(DestinationFile):
        os.remove(DestinationFile)

    chrome_options = webdriver.ChromeOptions()

    preferences = {"directory_upgrade": True,
                   "download.default_directory": str(scpathTemp).replace('/','\\'),
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

    driver.get(url)
    try:
        try:
            # 1st page
            for el in  driver.find_elements_by_tag_name('a'):
                if str(el.get_attribute('href')).strip().lower().find(xt(pubname))>=0 and \
                        str(el.get_attribute('href')).strip().lower().find(xt(".pdf")) >= 0 and \
                        not (str(el.get_attribute('innerText')).strip().lower().find(xt("survey")) >= 0):
                    el.click()
                    time.sleep(3)
                    wait_file_complete(scpathTemp)
                    time.sleep(3)
                    bok = True
                    res = True
                    break

            if bok == True:
                time.sleep(3)
                FileList = glob.glob(scpathTemp + "\*")
                if len(FileList) > 0:
                    LatestFile = max(FileList, key=os.path.getctime)
                    os.rename(LatestFile, SFilename)
                    shutil.move(SFilename, DestinationFile)
                    time.sleep(3)
                    res = True
        except Exception as e:
            strerror = str(e)
            strerror = "Failed - <Unable to Download Source File>"
    except Exception as e:
        strerror = str(e)

    if res == False:
        strerror = "Failed - <Unable to Download Source File>"
    if not os.path.exists(DestinationFile):
        strerror = "Failed - <Unable to Download Source File>"

    driver.close()
    if os.path.exists(scpathTemp):
        shutil.rmtree(scpathTemp)
    return DestinationFile, strerror


def update_1566078_10001(scdownloaded, obsfilepath, freq, path_to_save, dictUpdateIDsfrmSc, dictMD, dictRoundg):
    # global link_folder
    global dictSCDate
    dictcol = {}
    obs = xlz.gencache.EnsureDispatch("Excel.Application")
    sf = xlx.gencache.EnsureDispatch("Excel.Application")

    sourceFile = ""
    updateStatus = ""
    obsFile = str(obsfilepath)
    if str(scdownloaded).find(".pdf") > 0:
        result = convertpdftoexcel(str(scdownloaded))
        if len(result['error']) > 0:
            ##input the error message to metadata
            updateStatus = str(result['error'])
        else:
            sourceFile = str(result['filepath'])
            scdownloaded = scdownloaded + ";" + sourceFile
    else:
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
        obsSHT = None
        scsht = None

        for list_sht in obsBk.Sheets:
            if str(list_sht.Name).find(freq) > -1:
                obsSHT = obsBk.Worksheets[list_sht.Name]
                # fill_obs(obsSHT,freq)
                break

        bcomplete = False
        filename = scBk.Name
        icount = 0

        try:
            scsht = scBk.Worksheets[1]
            set_sheet(scsht)
        except Exception as e:
            str(e)
            scsht = None

        if obsSHT is not None and scsht is not None:
            for ik in dictUpdateIDsfrmSc:
                dictSCDate = {}
                ref_rng = None
                res = False
                str_find = str(dictUpdateIDsfrmSc[ik]).strip().lower()
                tofind = str(str_find).strip().split('|')[0].strip()
                try:
                    if str_find.find('2|')>=0:
                        res = get_data_date_102(scsht, str_find)
                    else:
                        res = get_data_date_101(scsht, str_find, scBk)
                except:
                    res = False

                if res and len(dictSCDate) > 0:
                    try:
                        for varDate in sorted(dictSCDate):
                            scData = ''
                            scData = dictSCDate[varDate]
                            # scData = only_digits(dictSCDate[varDate])
                            if scData is not None:
                                if is_number(scData):
                                    if len(dictMD) > 0:
                                        if dictMD.get(ik) is not None:
                                            scData = float(scData) / float(dictMD[ik])

                                    if len(dictRoundg) > 0:
                                        if dictRoundg.get(ik) is not None:
                                            scData = round(float(scData), int(dictRoundg[ik]))

                            UpdateOneDateOneValueOptimized(scData, ik, varDate, obsSHT, dictcol)
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

        obsBk.Close(SaveChanges=False)
        scBk.Close(SaveChanges=False)
    else:
        if obsht is nothing:  updateStatus = "Failed - <Could not find freq sheet>"


    if obs.Workbooks.Count == 0:
        obs.Quit()
    if sf.Workbooks.Count == 0:
        sf.Quit()
    return (scdownloaded, updateStatus)


def get_data_date_101(sh, tofind, scbk):
    res = False
    dc = None
    dc_row = None
    rng_tbl = None
    rng_row = None
    rng_col = None
    irow = 0 ; icol = 0 ; prow = 0
    initdata = 0 ; scdata =0
    # table | rows  | columns
    iyr =0 ; im =0
    tmpdata = tofind.strip()
    tmpdata = tmpdata.replace('1|', '')
    ch = xt(tmpdata).split('|')
    try:
        # cycle sheet
        set_sheet(sh)
        rng_tbl = find_mapping(sh,xt(ch[0]),':',10)

        if rng_tbl is not None:
            rng_row = find_mapping(sh, xt(ch[1]),':',rng_tbl.Row)
            rng_col = find_mapping(sh, xt(ch[2]),':',rng_tbl.Row)
            if rng_row is not None:
                irow = rng_row.Row
                endcol  = rng_col.Column +1
                iyr = only_digits(rng_tbl.Text)
                im = ex_month(rng_tbl.Text)
                # col / date
                dc =find_mapping(sh,'survey results',':',rng_tbl.Row)
                while rng_col.Column <= endcol:
                    icol = rng_col.Column
                    if is_date(iyr,im):
                        mydate = datetime.datetime(int(iyr),int(im),1)
                        if mydate not in dictSCDate:
                            scdata = sh.Cells(irow,icol).Value
                            if (scdata) == '…':
                                scdata =0
                            dictSCDate.update({mydate: scdata})
                            res = True
                    rng_col = sh.Cells(rng_col.Row , rng_col.Column + 1)
                    im= im -1
    except Exception as e:
        str(e)
        return False
    return res

def get_data_date_102(sh, tofind):
    res = False
    dc = None
    dc_row = None
    rng_tbl = None
    rng_row = None
    rng_col = None
    irow = 0 ; icol = 0 ; prow = 0
    initdata = 0 ; scdata =0
    # table | rows  | columns
    iyr =0 ; im =0
    tmpdata = tofind.strip()
    tmpdata = tmpdata.replace('2|', '')
    ch = xt(tmpdata).split('|')
    try:
        # cycle sheet
        set_sheet(sh)
        rng_tbl = find_mapping(sh,xt(ch[0]),':',10)

        if rng_tbl is not None:
            rng_row = find_mapping(sh, xt(ch[1]),':',rng_tbl.Row)
            rng_col = find_mapping(sh, xt(ch[2]),':',rng_tbl.Row)
            if rng_row is not None:
                irow = rng_row.Row
                endcol  = rng_col.Column +1
                iyr = only_digits(rng_tbl.Text)
                im = ex_month(rng_tbl.Text)
                # col / date
                dc =find_mapping(sh,'survey results',':',rng_tbl.Row)
                icol = rng_col.Column
                if is_date(iyr,im):
                    mydate = datetime.datetime(int(iyr),int(im),1)
                    if mydate not in dictSCDate:
                        scdata = sh.Cells(irow,icol).Value
                        if scdata == '…':
                            scdata = 0
                        dictSCDate.update({mydate: scdata})
                        res = True
                rng_col = sh.Cells(rng_col.Row , rng_col.Column + 1)
    except Exception as e:
        str(e)
        return False
    return res

def find_mapping(scsht, tofind, delim, st_row=1, ind = 1):
    dc = None
    dc = scsht.Cells(st_row, 1)
    if ind == 1:
        look_at = xlwings.constants.LookAt.xlPart
    else:
        look_at = xlwings.constants.LookAt.xlWhole


    for tmpdata in tofind.split(delim):
        if dc is not None:
            dc = scsht.Cells.Find(What=tmpdata.strip(), LookAt=look_at,
                                  SearchOrder=xlwings.constants.SearchOrder.xlByRows,
                                  MatchCase=False, After=dc, SearchDirection= xlwings.constants.SearchDirection.xlNext)

    outputRng = None
    if dc is not None:
        outputRng = dc
    return outputRng

def find_mapping_2(scsht, tofind, delim, st_row=1, srch_order=xlwings.constants.SearchOrder.xlByColumns, look_at=xlwings.constants.LookAt.xlPart,
                 srch_direction=xlwings.constants.SearchDirection.xlNext):
    dc = None
    dc = scsht.Cells(st_row, 1)
    for tmpdata in tofind.split(delim):
        if dc is not None:
            dc = scsht.Cells.Find(What=tmpdata.strip(), LookAt=look_at,
                                  SearchOrder=srch_order,
                                  MatchCase=False, After=dc, SearchDirection= srch_direction)

    outputRng = None
    if dc is not None:
        outputRng = dc
    return outputRng


def rev_srch(sh, tofind, st=1):
    dc = None
    dc = sh.Cells(st, 1)
    dc = sh.Cells.Find(What=tofind.strip(), LookAt=xlwings.constants.LookAt.xlPart,
                          SearchOrder=xlwings.constants.SearchOrder.xlByRows,
                          MatchCase=False, After=dc, SearchDirection= xlwings.constants.SearchDirection.xlPrevious)
    return dc

def only_digits(s):
    tmpstr = str(s)
    scdata = ""
    for ctr in range(0, len(tmpstr)):
        if is_number(str(tmpstr)[ctr]):
            scdata = scdata + s[ctr]

    return scdata


def only_letters(s):
    scdata = ""
    for ctr in range(0, len(s)):
        if isalpha(str(s)[ctr]) or (str(s)[ctr]) == '/':
            scdata = scdata + s[ctr]

    return scdata

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
        return False


def is_year(iyr):
    try:
        if len((only_digits(iyr))) == 4:
            return True
        else:
            return  False
    except ValueError:
        return  False

def find_month(sh,st_row, row_lim = 5):
    rngfind = None
    limrow = st_row + row_lim
    for i in range(1, 12):
        strfind = int_to_month(i)
        rngfind = find_mapping(sh, strfind,':', st_row)
        if rngfind is not None:
            if not rngfind.Row > limrow and not (str(rngfind.Text).lower().find('end of') >= 0) and not rngfind.Row == st_row:
                return rngfind
    return None


def find_month_Q(sh,st_row):
    rngfind = None
    limrow = st_row + 5
    for i in range(1,4):
        strfind = 'q' + str(i)
        rngfind = find_mapping(sh, strfind,':', st_row)
        if rngfind is not None:
            if not rngfind.Row > limrow and rngfind.Row > st_row:
                return rngfind

    return None


def find_year(sh,st_row, rng = None):
    rngfind = None
    iyr = datetime.date.today().year
    limyr = iyr - 20
    limrow = st_row + 10
    for x in range(iyr,limyr,-1):
        rngfind = find_mapping(sh,str(x),':',st_row,2)
        if rngfind is not None:
            tmpdata = only_digits(str(rngfind.Text).strip())
            if len(tmpdata) == 4:
                if rngfind.Row < limrow:
                    return rngfind

    return None


def rev_year(sh, row, col):
    lim = row - 5
    for x in range(row -1, lim, -1):
        for y in range(1, 20):
            txt = str(sh.Cells(x,y).Text).strip()
            if txt.find('/') >=0 and is_year(left(txt,4)):
                return sh.Cells(x,y)
            if is_year(txt):
                return sh.Cells(x, y)

def findcol_range(sh, tofind):
    dc= None
    rng = None

    lrow = 0; lcol =0
    ch = xt(tofind).split(':')
    dc =find_mapping(sh, xt(ch[0]), '|')
    if dc is not None:

        lrow = dc.Row +4; lcol = dc.Column +2
        frng = convertColSTR(sh,dc.Column -1) + str(dc.Row)
        erng = convertColSTR(sh,lcol) + str(lrow)

        rng= sh.Range(frng + ':' + erng)
        dc = rng.Find(What=xt(ch[1]), LookAt=xlwings.constants.LookAt.xlPart,
                              SearchOrder=xlwings.constants.SearchOrder.xlByColumns,
                              MatchCase=False)

    try:
        if dc.Row <= lrow and dc.Column <= lcol:
            return (dc,rng)
    except:
        return None,None


def left(s, amount):
    return s[:amount]

def right(s, amount):
    return s[-amount:]

def mid(s, offset, amount):
    return s[offset:offset+amount]

def fill_obs(obsht, freq):
    if str(freq).lower().find('m') >=0:
        fill_m(obsht)
    if str(freq).lower().find('a3') >=0:
        fill_a3(obsht)
    if str(freq).lower().find('q') >=0:
        fill_q(obsht)
    if str(freq).lower().find('d') >=0:
        fill_d(obsht)


def fill_m(obsht):
    dc = obsht.Range("N1")
    st_row = dc.Column
    endcol = dc.Column + 20
    iyr = int(datetime.datetime.today().year)
    im = int(datetime.datetime.today().month)

    for x in range(st_row, endcol):
        rangeObj = obsht.Range("N:N")
        rangeObj.EntireColumn.Insert()

    for x in range(st_row, endcol):
        # rangeObj = obsht.Range("N:N")
        # rangeObj.EntireColumn.Insert()
        mydate = datetime.datetime(iyr,im , 1)
        dc = obsht.Cells(1,x)
        dc.Value = mydate.strftime('%m-%d-%Y')
        dc.NumberFormat = "mmm-yyyy"
        obsht.Columns(convertColSTR(obsht, dc.Column) + ":"  + convertColSTR(obsht, dc.Column)).EntireColumn.AutoFit()
        if im == 1:
            im = 12
            iyr = iyr - 1
        else:
            im = im -1

def fill_d(obsht):
    dc = obsht.Range("N1")
    st_row = dc.Column
    endcol = dc.Column + 100
    iyr = int(datetime.datetime.today().year)
    im = int(datetime.datetime.today().month)
    iday = int(datetime.datetime.today().day)

    mydate = datetime.datetime(iyr, im, iday)

    for x in range(st_row, endcol):
        rangeObj = obsht.Range("N:N")
        rangeObj.EntireColumn.Insert()

    for x in range(st_row, endcol):
        # rangeObj = obsht.Range("N:N")
        # rangeObj.EntireColumn.Insert()
        # mydate = datetime.datetime(iyr,im , iday)
        dc = obsht.Cells(1,x)
        dc.Value = mydate.strftime('%m-%d-%Y')
        dc.NumberFormat = "mmm-yyyy"
        obsht.Columns(convertColSTR(obsht, dc.Column) + ":"  + convertColSTR(obsht, dc.Column)).EntireColumn.AutoFit()
        mydate = mydate - timedelta(days=1)


def fill_a3(obsht):
    dc = obsht.Range("N1")
    st_row = dc.Column
    endcol = dc.Column + 20
    iyr = int(datetime.datetime.today().year)
    im = 6

    for x in range(st_row, endcol):
        rangeObj = obsht.Range("N:N")
        rangeObj.EntireColumn.Insert()

    for x in range(st_row, endcol):
        # rangeObj = obsht.Range("N:N")
        # rangeObj.EntireColumn.Insert()
        mydate = datetime.datetime(iyr,im , 1)
        dc = obsht.Cells(1,x)
        dc.Value = mydate.strftime('%m-%d-%Y')
        dc.NumberFormat = "mmm-yyyy"
        obsht.Columns(convertColSTR(obsht, dc.Column) + ":"  + convertColSTR(obsht, dc.Column)).EntireColumn.AutoFit()
        iyr = iyr - 1


def fill_q(obsht):
    dc = obsht.Range("N1")
    st_row = dc.Column
    endcol = dc.Column + 20
    iyr = int(datetime.datetime.today().year)
    im = int(datetime.datetime.today().month)
    im = int(im / 3) * 3

    for x in range(st_row, endcol):
        rangeObj = obsht.Range("N:N")
        rangeObj.EntireColumn.Insert()

    for x in range(st_row, endcol):
        # rangeObj = obsht.Range("N:N")
        # rangeObj.EntireColumn.Insert()
        mydate = datetime.datetime(iyr,im , 1)
        dc = obsht.Cells(1, x)
        dc.Value = mydate.strftime('%m-%d-%Y')
        dc.NumberFormat = "mmm-yyyy"
        obsht.Columns(convertColSTR(obsht, dc.Column) + ":"  + convertColSTR(obsht, dc.Column)).EntireColumn.AutoFit()
        if im == 3:
            im = 12
            iyr = iyr - 1
        else:
            im = im -3


def letters(input):
    return ''.join(filter(str.isalpha, input))


def roman_numeral_quarter(input):
    txt = str(input).strip().lower()
    if txt == 'i':
        return 3
    if txt == 'ii':
        return 6
    if txt == 'iii':
        return 9
    if txt == 'iv':
        return 12
    else:
        return  0

def words_quarter(input):
    txt = str(input).strip().lower()
    if txt == 'first':
        return 3
    if txt == 'second':
        return 6
    if txt == 'third':
        return 9
    if txt == 'fourth':
        return 12
    else:
        return  0


def has_date(input):
    if input.find(' en ') > 0:
        txt = str(input).strip().lower()
        txt = txt.split('en')[1].strip()
        var = []
        var = txt.split(' ')
        for val in var:
            if getmonth_french(val.strip().lower()) > 0:
                return  True
    return False

def xt(xx):
    return  str(xx).strip().lower()


def set_sheet(sh):
    rng = sh.Range(sh.Cells(1, 1), sh.Cells(1000, 100))
    rng.MergeCells = False
    rng.WrapText = False
    sh.Range("A:Z").EntireColumn.Hidden = False

def find_sheet(scbk, toFind):
    for sh in scbk.Sheets:
        if xt(sh.Name).find(xt(toFind))>=0:
            return  sh
    return None


def ex_month(ss):
    for x in str(ss).split(' '):
        if getmonth(x.strip()) > 0:
            return getmonth(x.strip())
    return 0
