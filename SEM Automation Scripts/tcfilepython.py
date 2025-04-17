import time
import win32com.client  # For Excel automation
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pyautogui  # To bring the browser to the front

# Attach to Active Excel Workbook
excel = win32com.client.Dispatch("Excel.Application")
wb = excel.ActiveWorkbook
ws = wb.ActiveSheet  # Active sheet

# Extract Tactic Code column dynamically
header_row = ws.Rows(1)
tactic_code_col = None
for col in range(1, ws.UsedRange.Columns.Count + 1):
    if ws.Cells(1, col).Value and "Tactic code" in str(ws.Cells(1, col).Value).lower():
        tactic_code_col = col
        break

if not tactic_code_col:
    print("Error: 'Tactic code' column not found!")
    exit()

# Get all tactic codes (skip header)
tactic_codes = [ws.Cells(row, tactic_code_col).Value for row in range(2, ws.UsedRange.Rows.Count + 1) if ws.Cells(row, tactic_code_col).Value]

# ✅ Fix: Ensure Chrome opens normally (remove headless mode)
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")  # Open in full screen
options.add_experimental_option("detach", True)  # Prevents closing immediately

# Initialize WebDriver (Make sure ChromeDriver is installed)
driver = webdriver.Chrome(options=options)

# ✅ Fix: Bring Chrome to the front
time.sleep(2)  # Wait a bit for Chrome to open
pyautogui.hotkey('alt', 'tab')  # Switch focus to Chrome

# Open website
driver.get("https://gcampaign.gene.com/")
time.sleep(3)  # Wait for the page to load

# Login
driver.find_element(By.ID, "login-username").send_keys("shubham.puri@perficient.com")
driver.find_element(By.ID, "login-password").send_keys("ShubhamP111")
driver.find_element(By.ID, "login-login").click()
time.sleep(5)  # Wait for login to complete

# Process each tactic code
for tactic_code in tactic_codes:
    print(f"Processing Tactic Code: {tactic_code}")
    
    # Search for tactic code
    search_box = driver.find_element(By.NAME, "search")
    search_box.clear()
    search_box.send_keys(tactic_code)
    search_box.send_keys(Keys.ENTER)

    time.sleep(3)  # Wait for search results

print("Process Completed!")

# Close browser
driver.quit()
