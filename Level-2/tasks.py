from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
import csv
from RPA.PDF import PDF
import shutil
import os

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=100,
    )
    open_robot_order_website()
    orders = get_orders()
    for row in orders:
        fill_the_form(row)
    archive_receipts()
        
def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    close_annoying_modal()

def download_orders_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def get_orders():
    """Get Orders data from orders CSV file"""
    download_orders_file()
    with open('orders.csv', newline='') as csvfile:
        reader = csv.DictReader(csvfile)
        data = [row for row in reader]
    return data

def close_annoying_modal():
    """Close the popup"""
    page = browser.page()
    page.click("text=OK")

def fill_the_form(orders_row):
    """complete the order filling"""
    page = browser.page()
    page.select_option("#head", str(orders_row["Head"]))
    id_body="#"+"id-body-"+orders_row["Body"]
    page.click(id_body)
    page.fill("//*[@placeholder='Enter the part number for the legs']", orders_row["Legs"])
    page.fill("//*[@id='address']",orders_row["Address"])

    page.click("text=preview")
    page.wait_for_selector("#robot-preview-image",timeout=5000)

    page.click("#order")
    click_order = True
    counter = 0
    while click_order == True and counter < 5: 
        try:
            page.wait_for_selector("#receipt",timeout=5000)
            click_order = False
        except:
            page.click("#order")
            counter+=1

    pdf_file = store_receipt_as_pdf(orders_row["Order number"])
    screenshot = screenshot_robot(orders_row["Order number"])
    embed_screenshot_to_receipt(screenshot, pdf_file)
    page.click("#order-another")
    close_annoying_modal()

def store_receipt_as_pdf(order_number):
    """Store order completion receipt PDF File"""
    page = browser.page()
    receipt = page.locator("#receipt").inner_html()
    pdf = PDF()
    pdf.html_to_pdf(receipt, "output/receipts/"+order_number+".pdf") 
    return "output/receipts/"+order_number+".pdf"

def screenshot_robot(order_number):
    """Take a screenshot of the page"""
    page = browser.page()
    robot = page.locator("#robot-preview-image")
    robot.screenshot(path="output/receipts/"+order_number+".png")
    return "output/receipts/"+order_number+".png"

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """merge screenshot into pdf file"""
    pdf=PDF()
    pdf.add_files_to_pdf(files=[pdf_file,screenshot], target_document=pdf_file)

def archive_receipts():
    """Archive the receip folder and create ZIP"""
    if os.path.exists("output/receipts"+ '.zip'):
        os.remove("output/receipts"+ '.zip')
    shutil.make_archive("output/receipts", 'zip', "output/receipts")