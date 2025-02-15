from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
import csv
from RPA.PDF import PDF
from PIL import Image
from RPA.FileSystem import FileSystem
import shutil

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(browser_engine="chrome", slowmo=100)
    open_robot_order_website()
    get_orders()
    archive_receipts()
    
def open_robot_order_website():
    browser.goto('https://robotsparebinindustries.com/#/robot-order')

def get_orders():
    """Downloads excel file from the given URL"""
    http=HTTP()
    http.download('https://robotsparebinindustries.com/orders.csv', overwrite=True)

    """Read data from excel and return in a variable"""
    with open('orders.csv', 'r') as file:
        orders=csv.DictReader(file)
        for row in orders:
            order=row['Order number']
            head=row['Head']
            body=row['Body']
            legs=row['Legs']
            address=row['Address'] 
            place_order(order, head, body, legs, address)            

def close_annoying_modal():
    try:
        """ close the pop-up """
        page=browser.page()
        page.click('//*[@id="root"]/div/div[2]/div/div/div/div/div/button[1]')
    except Exception as e:
        print(f"No pop-up: {e}")

def place_order(order, head, body, legs, address):
        """ palce order on robote page """
        close_annoying_modal()
        page=browser.page()
        page.select_option("#head", index=int(head))
        page.check('#id-body-' + str(body))
        page.fill("input[placeholder='Enter the part number for the legs']", str(legs))
        page.fill('#address', address)
        page.click('#preview')
        page.click('#order')
        
        if page.locator("div.alert.alert-danger").is_visible():
            page.click('#order')
        if page.locator("div.alert.alert-danger").is_visible():
            page.click('#order')

        if not page.locator("div.alert.alert-danger").is_visible():
            # save receipt & screenshot
            pdf_file=store_receipt_as_pdf(order)
            screenshot_path=screenshot_robot(order)
            # append the screenshot in pdf
            embed_screenshot_to_receipt(screenshot_path, pdf_file)

            page.click('#order-another')
    
def store_receipt_as_pdf(order_number):
    """Export the page to a pdf file"""
    page = browser.page()
    robot_results_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    pdf_path="output/receipt/robot_receipt_" + str(order_number) + ".pdf"
    pdf.html_to_pdf(robot_results_html, pdf_path)
    return pdf_path

def screenshot_robot(order_number):
    """Save the screenshot"""
    page = browser.page()
    screenshot_path="output/screenshot/robot_image_" + str(order_number) + ".png"
    page.locator('#robot-preview-image').screenshot(path= screenshot_path)
    return screenshot_path

def embed_screenshot_to_receipt(screenshot, pdf_file):
    # Convert PNG to PDF
    image_pdf_path = screenshot.replace('.png', '.pdf')
    image = Image.open(screenshot)
    image.save(image_pdf_path, "PDF", resolution=100.0)
    # Append the PDF to the existing robot_receipt.pdf
    pdf=PDF()
    pdf.add_files_to_pdf([image_pdf_path], target_document=pdf_file, append=True)

def archive_receipts():
    """convert all the robot receipts to zip"""
    fs=FileSystem()
    pdf_path='output/receipt'
    zip_path='output/robot_receipt'
    # Create ZIP file
    shutil.make_archive(zip_path, 'zip', pdf_path)
