import pandas as pd
import xlwings as xw
# from tabula import read_pdf
import os
import configparser 

pd.options.mode.chained_assignment = None  # default='warn'

bank_codes = {"UOB": "7375", 
    "POSB": "7171", 
    "DBS":"7171", 
    "OCBC":"7339", 
    "MAYBANK":"7302", 
    "STANDARD CHARTERED": "9496"}

presetcolumns = ['Company Code',
            'Document Date',
            'Posting Date',
            'Fiscal Period',
            'Document Type',
            'Identifier to Get Line Items',
            'Vendor Account',
            'Curr Key in Doc',
            'Payment Method',
            'Payment Terms',
            'Assignment Number',
            'Text',
            'To Park(PK) or Post (PT)?',
            'Building Name / Unit No/City',
            'Postal Cd',
            'Bank Country',
            'Transaction Type']

def readSettings(path):
    config = configparser.ConfigParser()
    config.read(path)
    d = {}
    for i in config["MAIN"]:
        d[i] = config["MAIN"][i]
    return d

def saveSettings(path, new):
    config = configparser.ConfigParser()
    config.read(path)
    for k, v in new.items():
        print(k)
        config["MAIN"][k] = v
    with open(path, "w") as f:
        config.write(f)

def fixbranch(branch):
    if type(branch) == str:
        return branch
    elif len(str(int(branch))) == 2:
        branch = "0" + str(int(branch))
        return branch
    else:
        return str(int(branch))

def check_if_float(bankacc):
    if isinstance(bankacc, float):
        s = str(bankacc)[:-2]
        return s
    return bankacc

def readFreshExcel(PATH):
    wb = xw.Book(PATH)
    sheet = wb.sheets[0]
    
    excel = sheet.range('A4').options(pd.DataFrame, 
                                header=1,
                                index=False, 
                                expand="table").value
    # excel = df.loc[df["checked?"] == 1]
#     excel = excel.drop(columns=["checked?"])
    # excel[["S/No", "Contact"]] = excel[["S/No", "Contact"]].astype("int64")
    excel["AN"] = excel.apply(lambda x: check_if_float(x["AN"]), axis=1)
#     excel = excel.astype(str)

    return excel

def readExcel(PATH):
    wb = xw.Book(PATH)
    sheet = wb.sheets[0]

    df = sheet.range('A4').options(pd.DataFrame, 
                                header=1,
                                index=False, 
                                expand='table').value

    excel = df.loc[df["checker"] == True]
    # excel[["S/No", "Bank Code"]] = excel[["S/No", "Bank Code"]].astype(int)
#     excel["Branch Code"] = excel.apply(lambda x: fixbranch(x["Branch Code"]), axis=1)
    # excel = excel.astype(str)
    return excel

def populateFields(excel):
    excel["Bank Code"] = excel["Bank Name"].map(bank_codes)

    uob_ = excel[(excel["Bank Name"] == "UOB") & (excel["AN"].str.len() == 10)]
    uob_["temp"] = uob_["AN"].str[:3]
    uob_.loc[uob_["AN"].str[:4] == "3013", ["temp"]] = "3013"
    uob_.loc[uob_["AN"].str[:4] == "3011", ["temp"]] = "3011"
    uob_["Branch Code"] = uob_["temp"].map(uob_branch_codes)
    uob_ = uob_.drop(columns=["temp"])
    excel[(excel["Bank Name"] == "UOB") & (excel["AN"].str.len() == 10)] = uob_

    dbs = excel[(excel["Bank Name"] == "DBS")]
    dbs["Branch Code"] = dbs["AN"].str[:3]
    excel.loc[excel["Bank Name"] == "DBS"] = dbs

    posb = excel.loc[excel["Bank Name"] == "POSB"]
    posb["Branch Code"] = "081"
    excel.loc[excel["Bank Name"] == "POSB"] = posb

    sc = excel.loc[excel["Bank Name"] == "STANDARD CHARTERED"]
    sc["Branch Code"] = "0" + sc["AN"].str[:2]
    excel.loc[excel["Bank Name"] == "STANDARD CHARTERED"] = sc

    empty = excel.loc[excel["Branch Code"].isnull()]
    empty["Branch Code"] = "000"
    excel.loc[excel["Branch Code"].isnull()] = empty

    excel["Masked_Num"] = "XXXX" + excel["Num"].str[-5:]

    excel["Bank Key (Bank and Branch Code)"] = excel["Bank Code"] + excel["Branch Code"]
    
#     excel.loc[excel['Branch Code'] == "000", 'Error'] = 'yes'
    excel[["Bank Name", "Bank Code", "Branch Code", "AN", "BAH","Name ", "Num", "Masked_Num", "Bank Key (Bank and Branch Code)"]] = excel[["Bank Name", "Bank Code", "Branch Code", "AN", "BAH","Name ", "Num", "Masked_Num", "Bank Key (Bank and Branch Code)"]].astype(str)
    return excel 

def writetoexcel(PATH, populated_fields):
    wb = xw.Book(PATH)
    sheet = wb.sheets[0]
    sheet.range('A4').options(pd.DataFrame,
                                index=False,
                                headers=True,
                                expand="table").value = populated_fields
    rng = xw.Range('A4').options(expand="table").current_region
    rng.autofit()
    for border_id in range(7,13):
        rng.api.Borders(border_id).LineStyle=1
        rng.api.Borders(border_id).Weight=2
    rng.api.Font.ColorIndex = 1

def generateUOB():
    pdf = read_pdf("bank codes.pdf", pages="3-13", encoding="utf-8", stream=True)
    uob = pd.DataFrame(columns=["branch name", "acc no", "branch code"])
    for i in range(11):
        temp = pdf[i].loc[3:]
        temp.columns= (["branch name", "acc no", "branch code"])
        uob = uob.append(temp)
    uob = uob.reset_index()
    uob = uob.drop(["index"], axis=1)

    uob.at[248, "branch name"] = "UOB Wealth Banking Scotts Square"
    uob.at[248, "acc no"] = 633
    uob.at[248, "branch code"] = "7375 632"


    uob.at[288, "branch name"] = "UOB Wealth Banking Scotts Square"
    uob.at[288, "acc no"] = 722
    uob.at[288, "branch code"] = "7375 632"

    uob.at[85, "acc no"] = 301
    uob.at[85, "branch code"] = "7375 001"

    uob.at[88, "acc no"] = 301
    uob.at[88, "branch code"] = "7375 046"

    
    uob.dropna(inplace=True)
    
    
    temp = uob["branch code"].str.split(" ", expand = True)
    uob["branch code"] = temp[1]
    uob["bank code"] = temp[0]

    uob = uob.astype(str)
    
    uob.loc[(uob["acc no"] == "301") & (uob["branch code"] == "001"), ["acc no"]] = "3013"
    uob.loc[(uob["acc no"] == "301") & (uob["branch code"] == "046"), ["acc no"]] = "3011"
    
    return uob

def readUOB(PATH):
    uob = pd.read_csv(PATH, dtype={'branch code':'string'})
    uob = uob.astype(str)
    return uob
    
def main_filter(bank_name, bank_code, branch_code, acc_no, nric, masked_nric, bankkey):
    m_nric = "XXXX" + nric[-5:]
    if m_nric != masked_nric:
        return False
    elif bankkey != bank_code + branch_code:
        return False
    elif bank_name == "UOB" and bank_code == bank_codes[bank_name] and len(acc_no) == 10:
        digits = acc_no[:3]
        if digits == "301":
            digits = acc_no[:4]
        result = uob[uob["acc no"] == digits]
        if len(result) == 0:
            return False
        elif result["branch code"].values[0] == branch_code:
            return True
    elif bank_name == "UOB" and bank_code == bank_codes[bank_name] and len(acc_no) in [7, 9, 11, 12, 13, 14, 17, 18] and branch_code == "001":
        return True
    elif bank_name == "DBS" and bank_code == bank_codes[bank_name] and len(acc_no) == 10 and branch_code == acc_no[:3]:
        return True
    elif bank_name == "POSB" and bank_code == bank_codes[bank_name] and len(acc_no) == 9 and branch_code == "081":
        return True
    elif bank_name == "STANDARD CHARTERED" and bank_code == bank_codes[bank_name] and branch_code == "0" + acc_no[:2]:
        return True
    elif bank_name == "MAYBANK" and bank_code == bank_codes[bank_name]:
        return True
    else:
        return False

def check(excel):
    excel[["Bank Name", "Bank Code", "Branch Code", "AN", "BAH","Name ", "Num", "Masked_Num", "Bank Key (Bank and Branch Code)"]] = excel[["Bank Name", "Bank Code", "Branch Code", "AN", "BAH","Name ", "Num", "Masked_Num", "Bank Key (Bank and Branch Code)"]].astype(str)

    excel["checker"] = excel.apply(lambda x: main_filter(x["Bank Name"], x["Bank Code"], x["Branch Code"], x["AN"], x["Num"], x["Masked_Num"], x["Bank Key (Bank and Branch Code)"]), axis=1)
    
    return excel

def workingdf(PATH, checked):
    wb = xw.Book(PATH)
    df = wb.sheets[0].range('A1').options(pd.DataFrame, 
                                    header=1,
                                    index=False, 
                                    expand='right').value
    
    df[["Reference Document Number", "Vendor Name 1", "Bank Key", "Bank Account", "Amt in Doc Curr."]] = checked[["Masked_Num", "BAH", "Bank Key (Bank and Branch Code)", "AN", "Amount Disbursed \nto Member "]]
    return df

def writetotemplate(presetcolumns, working):
    wb = xw.Book(settings["template"])
    row = wb.sheets[0].range('A' + str(wb.sheets[0].cells.last_cell.row)).end("up").row
    sheet = wb.sheets[0]

    for i in range(len(presetcolumns)):
        column = presetcolumns[i]
        if column == 'Identifier to Get Line Items':
            working[column] = "hello"
            continue
        if column == "Text":
            date_list = settings["date"].split(".")
            date = "/".join(date_list)
            text = to_add[i].format(settings["assignmentnumber"] ,date)
            working[column] = text
            continue
        if column == "Identifier to Get Line Items":
            continue
        working[column] = to_add[i]

    working['Identifier to Get Line Items'] = [f'T{x}' for x in range(row, row+len(working))]

    sheet.range("A{}".format(row+1)).options(index=False, header=False).value = working

if __name__ == "__main__":
    # --- part 1 ---
    settings = readSettings("checker/settings.ini")

    to_add = [settings["companycode"], 
                settings["date"], 
                settings["date"], 
                settings["fp"], 
                settings["documenttype"], 
                "placeholder", 
                settings["vendoraccount"], 
                settings["currkeyindoc"], 
                settings["paymentmethod"], 
                settings["paymentterms"], 
                settings["assignmentnumber"], 
                settings["text"], 
                settings["pp"], 
                settings["building"],  
                settings["postal"], 
                settings["bank_country"], 
                settings["transaction"]]

    uob = readUOB(settings["uob"])
    uob_branch_codes = dict(zip(uob["acc no"], uob["branch code"]))

    # excel = readFreshExcel(settings["path"])
    # fields = populateFields(excel)

    # results = check(fields)
    
    # writetoexcel(settings["path"], results)

    # --- part 2 ----
    results = readExcel(settings["path"])
    working = workingdf(os.path.join(settings["path"], settings["template"]), results)

    writetotemplate(presetcolumns, working)