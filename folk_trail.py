from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import openpyxl

# Initialize the Firefox WebDriver
driver = webdriver.Firefox()

# Load the webpage
driver.get("https://app.folk.app/shared/100-VC-firm-Investing-In-SaaS-eBQ61SEn13lP1A06ONpjdYyrI1dbgmT3")
time.sleep(6)

# Create a new Excel workbook and select the active sheet
wb = openpyxl.Workbook()
ws = wb.active

# Set the headers for the Excel sheet
ws.append(["Company Name", "URL","Description", "Investor Type","Stage Invested", "Thesis","Region Invested","Ticket Size Sweet spot","Value add"])

# Scroll function
def scroll():
    scroll_element = driver.find_element(By.CLASS_NAME, "c-cPjCsh")
    driver.execute_script("arguments[0].scrollBy(0, 40);", scroll_element)

# Extract data and store in Excel sheet
for i in range(1, 870):  # Adjust the range based on the number of rows
    company_name = driver.find_element(By.CSS_SELECTOR, f"[aria-rowindex='{i}'] .c-gZNEbh").text
    
    row = driver.find_element(By.XPATH, f'//div[@role="row" and @aria-rowindex="{i}" and contains(@class, "c-bxgLFE")]')
    
    url_element = row.find_element(By.XPATH, './/div[@role="gridcell"]//span[contains(@class, "c-gZNEbh")]')
    url = url_element.text
    
    desc = driver.find_element(By.XPATH, f"//div[@role='gridcell'][@aria-rowindex='{i}'][@aria-colindex='3']")
    desc_text = desc.text if desc.text else "NIL"
    
    thesis = driver.find_element(By.XPATH, f"//div[@role='gridcell'][@aria-rowindex='{i}'][@aria-colindex='6']")
    thesis_text = thesis.text if thesis.text else "NIL"
    
    value = driver.find_element(By.XPATH, f"//div[@role='gridcell'][@aria-rowindex='{i}'][@aria-colindex='9']")
    value_text = value.text if value.text else "NIL"
    
    
    investor_type = row.find_elements(By.XPATH, './/div[@aria-colindex="4"]//div[contains(@class, "PJLV")]//span[@class="c-hJobYV"]')
    investor = ', '.join([investor.text for investor in investor_type if investor.text.strip()])
    
    stage_invested = row.find_elements(By.XPATH, './/div[@aria-colindex="5"]//div[contains(@class, "PJLV")]//span[@class="c-hJobYV"]')
    stage = ', '.join([stage.text for stage in stage_invested if stage.text.strip()])
    
    region_invested = row.find_elements(By.XPATH, './/div[@aria-colindex="7"]//div[contains(@class, "PJLV")]//span[@class="c-hJobYV"]')
    region = ', '.join([region.text for region in region_invested if region.text.strip()])
    
    ticket_size = row.find_elements(By.XPATH, './/div[@aria-colindex="8"]//div[contains(@class, "PJLV")]//span[@class="c-hJobYV"]')
    ticket = ', '.join([ticket.text for ticket in ticket_size if ticket.text.strip()])
    
    ws.append([company_name,url,desc_text, investor,stage,thesis_text,region,ticket,value_text])
    wb.save("output3.xlsx")
    time.sleep(0.5)
    scroll()

# Close the WebDriver
driver.close()
