import os.path, sys, time, xlwings, win32gui

sys.path.append(r"C:\Users\englandt\AppData\Local\Continuum\anaconda3\Lib\site-packages\win32\\")
sys.path.append(r"C:\Users\englandt\AppData\Local\Continuum\anaconda3\Lib\site-packages\win32\lib\\")
import win32com.client  # for some reason this can't be imported until the paths ^^^ are appended

# hardcoded values
macro_wb_path = r"\\PSACLW02\PROJDATA\EnglandT\MISC\SCRIPTS\MacroBook.xlsm"
rfq_macro = r"MacroBook.xlsm!GetEmailInfo"
quote_wb_path = r"\\PSACLW02\PROJDATA\EnglandT\MATEER\QUOTES\Quote_Auto.xlsm"


def getQuoteNumber(opp_url, price, opp_typ):  # opp_type: NM / AM / CP
    from selenium import webdriver
    driver = webdriver.Chrome()  # use Chrome

    time.sleep(1)
    driver.get(opp_url)  # go to CRM opportunity
    time.sleep(1)  # allow CRM to load

    config_button = driver.find_element_by_id('gsliguidedselling')  # find 'Configure' button
    config_button.click()  # click the button
    time.sleep(5)  # allow CPQ to load

    driver.switch_to.frame(driver.find_element_by_id('gsiframe'))  # switch to iframe "gsiframe"

    try:  # if a quote has already been done & a new one is not in progress
        quote_number = driver.find_element_by_css_selector(
            "div[data-property='QuoteNumber']").text  # get text of quoteNum
        driver.switch_to.default_content()  # switch from iframe back to default content
        return quote_number
    except Exception:
        pass

    try:  # if a quote has never been started
        test_bool = True  # see if the 'try' was successful
        catalog_button = driver.find_element_by_name('SelectCatalog')  # find 'Select Catalog' button
        catalog_button.click()  # click the button
        time.sleep(5)  # allow time for new buttons to load
        if opp_typ.upper() == "NM":
            opptyp_button = driver.find_element_by_xpath(
                "//a[contains(@href,'#outerkey0')]")  # find button for "new machine"
        else:
            opptyp_button = driver.find_element_by_xpath(
                "//a[contains(@href,'#outerkey2')]")  # find button for "aftermarket"
        opptyp_button.click()  # click the button
        time.sleep(5)  # allow time for the menu to load
        if opp_typ.upper() == "NM":  # new machine
            smart_button = driver.find_element_by_id('headinginnerkey1')  # find button for "Smart Catalog"
        elif opp_typ.upper() == "AM":  # aftermarket/rebuild
            smart_button = driver.find_element_by_id('headinginnerkey5')  # find button for "Smart Catalog"
        else:  # aftermarket/change parts
            smart_button = driver.find_element_by_id('headinginnerkey3')  # find button for "Smart Catalog"
        smart_button.click()  # click the button
        time.sleep(5)  # allow time for menu to load
        if opp_typ.upper() == "NM":  # new machine
            mach_button = driver.find_element_by_link_text('Filler Dry')  # find button for "Filler Dry"
        elif opp_typ.upper() == "AM":  # aftermarket/rebuild
            mach_button = driver.find_element_by_link_text('Clw - Mateer Rebuild')  # find button for "Mateer Rebuild"
        else:  # aftermarket/change parts
            mach_button = driver.find_element_by_link_text(
                'Clw - Mateer Change Parts')  # find button for "Mateer Change Parts"
        mach_button.click()  # click the button -> go to CPQ
        time.sleep(5)  # allow time to load
        qty_box = driver.find_element_by_id('spp-2')  # text box to fill in (configuration quantity)
        qty_box.send_keys('1')  # fill in quantity as 1
        time.sleep(2)  # allow time to load
        price_box = driver.find_element_by_id('spp-4')  # find price textbox
        price_box.click()  # need to click it so the page reloads
        time.sleep(2)  # allow time to load
        price_box = driver.find_element_by_id('spp-4')  # find price textbox (again)
        price_box.click()  # click price text box (required?)
        price_box.send_keys(price)  # fill in price
        finish_button = driver.find_element_by_id('Finish')  # find "Finish" button
        finish_button.click()  # click the button
        time.sleep(5)  # allow time to load
    except Exception:  # a quote was started but is still in progress
        test_bool = False
        catalog_button = driver.find_element_by_xpath("//button[contains(@onclick,'true')]")  # find "Yes" button
        catalog_button.click()  # click the button

    time.sleep(1)  # allow opp to load
    # driver.switch_to.frame(driver.find_element_by_id('gsiframe'))  # switch to iframe "gsiframe"
    try:
        quote_number = driver.find_element_by_css_selector("div[data-property='QuoteNumber']").text  # get text
        driver.switch_to.default_content()  # switch from iframe back to default content
        driver.close()  # close chrome
        return quote_number
    except:
        try:
            quote_number = driver.find_element_by_css_selector("div[id='Summary-3']").text  # get text
            driver.switch_to.default_content()  # switch from iframe back to default content
            driver.close()  # close chrome
            return quote_number
        except:
            if test_bool:  # unknown error occurred
                print("Unknown error occurred")
            else:  # quote was started, but is in progress (can't tell where in the process it is)
                print("A quote is in progress -- must be finished manually")


def typeDialogBox():  # TODO: Make this look less ugly
    import tkinter as tk

    global answer
    answer = "x"
    root = tk.Tk()

    def assignValue(x):
        global answer
        answer = x
        root.destroy()

    tk.Label(root, text="Select quote type").grid(row=0)
    tk.Radiobutton(root,
                   text="New Machine",
                   indicatoron=0,
                   command=lambda: assignValue("NM")).grid(row=1)
    tk.Radiobutton(root,
                   text="Aftermarket",
                   indicatoron=0,
                   command=lambda: assignValue("AM")).grid(row=2)
    tk.Radiobutton(root,
                   text="Change Parts",
                   indicatoron=0,
                   command=lambda: assignValue("CP")).grid(row=3)

    root.mainloop()
    return answer


# set default values
cust_name = ""
url = ""
mach_model = ""
st_address = ""
zip_address = ""
country_add = ""
cont_name = ""
sales_exec = ""

quote_type = typeDialogBox().upper()  # get the quote type (NM, AM, CP)
if quote_type.upper() == "NM":  # aftermarket or CP ==> not new machine price
    mach_price = 0
else:
    mach_price = 1  # to be manually assessed & changed

xl = win32com.client.DispatchEx("Excel.Application")  # xl = Excel program
if os.path.exists(macro_wb_path):  # scratch macro wb exists
    # open wb and run macro
    xl.Visible = True
    wkbk = xl.Workbooks.Open(Filename=macro_wb_path)  # , ReadOnly=1) # <-- uncomment for read only
    xl.Application.Run(rfq_macro)  # run the macro to pull info out of the email
    # wkbk.Save()  # required because xlrd needs to re-open it (I don't understand why this is necessary)

    # get info out
    wkbk = xlwings.Book(macro_wb_path)  # re-opens wkbk
    # wsheet = wkbk.sheet_by_index(0)  # set worksheet (first sheet)
    wsheet = wkbk.sheets[0]  # set worksheet (first sheet)
    url = wsheet.range("A1").value  # get url
    if url.upper() == "NOTHING":
        print("No email found")
        wkbk.close()
        quit()
    cust_name = wsheet.range("A2").value  # get customer/company name
    st_address = wsheet.range("A3").value  # get street address
    zip_address = wsheet.range("A4").value  # get city/state/zip
    country_add = wsheet.range("A5").value  # get country
    cont_name = wsheet.range("A6").value  # get contact name
    mach_model = wsheet.range("A7").value  # get machine model
    sales_exec = wsheet.range("A8").value  # get sales exec
    wkbk.app.quit()  # close wkbk

if mach_model != None:  # a model value was taken out of the email
    xl = win32com.client.DispatchEx("Excel.Application")  # not redundant
    if os.path.exists(quote_wb_path):  # check for AutoQuote workbook
        xl.Visible = True
        wkbk = xlwings.Book(quote_wb_path)  # re-opens wkbk
        wsheet = wkbk.sheets[0]  # set worksheet (first sheet)
        wsheet.range("D6").value = mach_model  # fill in model info
        machprice_string = wsheet.range("E6").value  # get price
        wkbk.close()  # close window
        try:  # not sure if the model from the email is correct / in the list
            if quote_type.upper() == "NM":
                mach_price = machprice_string[2:len(machprice_string) - 1]  # trim off parentheses & dollar sign
            if mach_price == 0:  # no price (rotary?) or some other bad value
                mach_price = 1  # set to 1 for CRM
        except Exception:  # model value was no good
            mach_price = 1  # set to 1 for CRM
    else:  # model found but no AutoQuote workbook
        print("Error with Auto Quote workbook path")
        quit()

num = url.find('<')  # see if weird bracket thing is happening (only sometimes)
if num > 0:  # -1 = not found, but position of 0 is useless anyway
    url = url[:num]  # cut off content contained in <> at the end

quote_number = getQuoteNumber(str(url), mach_price, quote_type)  # get quote number

xl = win32com.client.DispatchEx("Excel.Application")  # not redundant
if os.path.exists(quote_wb_path):  # check for AutoQuote workbook
    xl.Visible = True
    wkbk = xlwings.Book(quote_wb_path)  # re-opens wkbk
    wsheet = wkbk.sheets[0]  # set worksheet (first sheet)
    wsheet.range("D4").value = quote_number  # quote number
    wsheet.range("D5").value = quote_type  # new machine, aftermarket, change parts
    wsheet.range("D6").value = mach_model  # fill in model info
    wsheet.range("D7").value = sales_exec  # sales exec
    wsheet.range("D8").value = cust_name  # company name
    wsheet.range("D9").value = st_address  # address1
    wsheet.range("D10").value = zip_address  # address2
    wsheet.range("D11").value = country_add  # country
    wsheet.range("D12").value = cont_name  # contact person


# bring AutoQuote workbook to foreground
def windowEnumeration(handle, windows_list):  # function to create list of windows
    windows_list.append((handle, win32gui.GetWindowText(handle)))  # create the list


windows_list = []  # declaring the list (outside the function)
win32gui.EnumWindows(windowEnumeration, windows_list)  # enumerate the windows & fill in list
for __ in windows_list:  # for each window in the list
    if "QUOTE_AUTO" in __[1].upper():  # if statement to identify the workbook
        win32gui.ShowWindow(__[0], 5)  # show the window
        win32gui.SetForegroundWindow(__[0])  # bring it to the top
        break
