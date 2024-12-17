from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
import zipfile
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
    download_order_csv()
    orders = get_orders()
    open_robot_order_website()
    order_robots(orders)
    archive_receipts()


def open_robot_order_website():
    browser.configure(slowmo=300)
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    
def download_order_csv():
    """Downloads the csv with the orders"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv",overwrite= True)
    
def get_orders():
    """Extracts the orders from the downloaded csv"""
    csv = Tables()
    orders = csv.read_table_from_csv("orders.csv", True, ["Order number", "Head", "Body", "Legs", "Address"])
    return orders

def close_annoying_modal():
    """Closes the modal that opens everytime on reload"""
    page = browser.page()
    page.click("button:text('OK')")


def order_robots(orders):
    """Creates the order"""
    for row in orders:
        fill_the_form(row) 

def fill_the_form(order):
    """Fills in the form for the creation of an order, retries up to 3 times if an error occurs."""
    
    max_attempts = 3
    attempts = 0
    success = False
    
    close_annoying_modal()
    while attempts < max_attempts and not success:
        try:
            attempts += 1
            page = browser.page()
            
           
            page.select_option("#head", order['Head'])
            page.set_checked("#id-body-" + order['Body'], True)
            page.fill("input[placeholder='Enter the part number for the legs']", order['Legs'])
            page.fill("#address", order['Address'])
            page.click("#preview")
            page.click("#order")
            
        
            if check_for_error():
                print(f"Attempt {attempts}: Error detected, retrying...")
            else:
                success = True 
        except Exception as exception:
            print(f"Attempt {attempts}: Exception occurred - {exception}")
    
    if not success:
        print(f"Failed to process order after {max_attempts} attempts and will be skipped: {order}")
        page.reload()
        return
    page = browser.page()
    order_number = order["Order number"]
    pdf_path = store_receipt_as_pdf(order_number)
    screenshot_path = screenshot_robot(order_number)
    embed_screenshot_to_receipt(screenshot_path, pdf_path)
    page.click("#order-another")
    
    
def check_for_error():
    """Checks if an error element is displayed in the HTML."""
    page = browser.page()
    error_element = ".alert-danger"  
    return page.query_selector(error_element) is not None

def store_receipt_as_pdf(order_number):
    """Stores the receipt from an order to a pdf file. The name of the pdf file is based by the order number"""
    pdf = PDF()
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()
    pdf_path = f"output/{order_number}.pdf"
    pdf.html_to_pdf(receipt_html, pdf_path)
    return pdf_path
    
def screenshot_robot(order_number):
    """Takes a screenshot of the ordered robot from the preview. Saves the image based by the order number"""
    page = browser.page()
    robot_preview = page.locator("#robot-preview-image")
    screenshot_path = f"output/{order_number}.png"
    robot_preview.screenshot(path=screenshot_path)
    return screenshot_path
    
def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Appends the taken screenshot to a pdf"""
    pdf = PDF()
    pdf.add_files_to_pdf(files=[screenshot], target_document=pdf_file, append=True)
    
def archive_receipts():
    """Takes all receipt-pdfs and creates a zip"""
    with zipfile.ZipFile("output/receipts.zip", 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk("output"):
            for file in files:
                if file.endswith(".pdf"):
                    file_path = os.path.join(root, file)
                    zipf.write(file_path, arcname=file)
    


    
    
    
    
    
    
    