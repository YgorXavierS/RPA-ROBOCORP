from robocorp.tasks import task
from robocorp import browser as br
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files

@task
def main_executer():
    '''Execute the main actions on this class'''
    br.configure(
        slowmo=1000,
    )
    open_the_intranet_website()
    login()
    downLoad_excel_file()
    input_data_row()
    img_downLoad()
    log_out()
    
def open_the_intranet_website():
    br.goto("https://robotsparebinindustries.com/")

def login():
    page = br.page()
    page.fill('#username','maria')
    ###ACEITA XPATH
    page.fill('//*[@id="username"]','maria')

    page.fill('//*[@id="password"]','thoushallnotpass')
    page.click('//*[@id="root"]/div/div/div/div[1]/form/button')

def fill_and_submit_sales_form(row):
    """Fills in the sales data and click the 'Submit' button"""
    page = br.page()
    page.fill("#firstname", row['First Name'])
    page.fill("#lastname", row['Last Name'])
    page.select_option("#salestarget", str(row['Sales Target']))
    page.fill("#salesresult", str(row['Sales']))
    page.click("text=Submit")

def downLoad_excel_file():
    http = HTTP()
    http.download(url='https://robotsparebinindustries.com/SalesData.xlsx',overwrite=True)


def input_data_row():
    excel = Files()
    ##OPEN EXCEL
    excel.open_workbook('SalesData.xlsx')
    worsheet =excel.read_worksheet_as_table('data',header=True)
    for row in worsheet:
        fill_and_submit_sales_form(row)
    excel.close_workbook()

def img_downLoad():
    page = br.page()
    #path
    'img/sales_summary.png'
    page.screenshot(path='img/sales_summary.png')


def log_out():
    """Presses the 'Log out' button"""
    page = br.page()  
    page.click("text=Log out")