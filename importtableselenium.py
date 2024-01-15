from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog
import time


# Initialize the web driver
driver = webdriver.Chrome()
driver.get("")

# Prompt user to input email with Tkinter
#email = simpledialog.askstring("Input", "Enter your email:")
email_input = driver.find_element(By.XPATH, '//*[@id="email"]')
email_input.send_keys("")
# Prompt user to input password with Tkinter
#password = simpledialog.askstring("Input", "Enter your password:")
password_input = driver.find_element(By.XPATH, '//*[@id="password"]')
password_input.send_keys("")

button_login = driver.find_element(By.XPATH, "//button[@type='button' and contains(@class, 'tw-bg-blue-500') and contains(@class, 'tw-text-white') and contains(@class, 'tw-w-full') and contains(@class, 'tw-py-2.5') and contains(@class, 'tw-mt-14') and contains(@class, 'tw-block')]")
button_login.click()

# Wait for up to 10 seconds until the element is present
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CLASS_NAME, "MuiAvatar-fallback"))
)

# Now try to find the element
avatar_element = driver.find_element(By.CLASS_NAME, "MuiAvatar-fallback")
avatar_element.click()

# Click on the span with the text "Painel de Controle"
panel_element = driver.find_element(By.XPATH, "//span[text()='Painel de Controle']")
panel_element.click()
time.sleep(2)

# Click on the link with the text "Usuários"
users_link = driver.find_element(By.XPATH, "//a[text()='Usuários']")
users_link.click()

# Type and search for users by name or email
emailsearch = simpledialog.askstring("Input", "Search user by email:")
search_input = driver.find_element(By.ID, "user")
search_input.send_keys(emailsearch)



search_button = driver.find_element(By.XPATH, "//button[text()='Buscar']")
search_button.click()
time.sleep(2)

# Click on the MoreVertIcon (Assuming there is only one MoreVertIcon on the page)
more_vert_icon = driver.find_element(By.CSS_SELECTOR, ".MuiSvgIcon-root.MuiSvgIcon-fontSizeMedium.tw-text-blue-500.css-vubbuv")
more_vert_icon.click()


# Wait for up to 10 seconds until the element is present
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, ".tw-absolute > .hover\\3Atw-no-underline > .tw-py-3"))
)

# Now try to find the element
edit_permissions = driver.find_element(By.CSS_SELECTOR, ".tw-absolute > .hover\\3Atw-no-underline > .tw-py-3")
edit_permissions.click()

time.sleep(2)


def select_excel():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.askopenfilename(
        title="Select the Excel file",
        filetypes=[("Excel", "*.xls;*.xlsx")]
    )
    return file_path

# Call the function to select the Excel file
file_path = select_excel()

if file_path:
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active  # Assuming you're working with the active sheet

    # Define the column you want to process Excel
    column_letter = 'A'    

    # Iterate through each row in the specified column
    for row in sheet[column_letter]:
        cell_data = row.value

        # Check if the cell is not empty
        if cell_data is not None:
            # Perform actions on the web form using Selenium
            # Search Box
            search_box = driver.find_element(By.ID, "unitName")
            search_box.send_keys(str(cell_data))
            
            ## Wait for up to 10 seconds until the element is present
            search_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.tw-bg-blue-500.tw-h-10.tw-px-6.tw-text-white:enabled"))
            )

            # Now try to click on the element
            search_button.click()
            
            
            time.sleep(2)
            # Click the check mark button
            # Select All (Flag)
            flag_select_all = driver.find_element(By.CSS_SELECTOR, "button[role='checkbox']")
            flag_select_all.click()

            # Click the add button
              
            # Add Button
            add_button = driver.find_element(By.CSS_SELECTOR, "svg.MuiSvgIcon-root.MuiSvgIcon-fontSizeInherit.css-1cw4hi4")
            add_button.click()

            # Add additional actions as needed

            # Pause to allow time for the web form to process
            time.sleep(2)  # You may need to adjust this duration

    # Close the browser window when done
    driver.quit()
