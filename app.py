from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pyperclip
from docx import *
import time
import read_excel
import os

def setup_driver():
    # Set up Chrome options to use a specific profile
    chrome_options = Options()
    chrome_options.add_argument("user-data-dir=" + os.getenv("CHROME_PROFILE_PATH"))  # Path to your Chrome profile
    chrome_options.add_argument("profile-directory=Profile 1")  # Profile name
    chrome_options.add_argument("start-maximized") # open Browser in maximized mode
    chrome_options.add_argument("disable-infobars") # disabling infobars
    chrome_options.add_argument("--disable-extensions") # disabling extensions
    chrome_options.add_argument("--disable-gpu") # applicable to windows os only
    chrome_options.add_argument("--disable-dev-shm-usage") # overcome limited resource problems
    chrome_options.add_argument("--no-sandbox") # Bypass OS security model

    # Initialize the Chrome WebDriver with the specified options and automatic driver management
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)
    return driver

def format_each_line(driver):
    spans = driver.find_elements(By.CSS_SELECTOR, "div[contenteditable='true'] span[data-text='true']")
    
    for index, span in enumerate(spans):
        # check if any(label in lesson_data for label in list_label):
        if index == 0:
            apply_font_size(driver, span, "20")
        if span.text.strip() and (
            any(label in span.text.strip() for label in read_excel.list_label) or 
            any(label in span.text.strip() for label in read_excel.extra_label.values())
        ):
            move_cursor_to_end_of_text(driver, span)
            time.sleep(1)  # Ensure the cursor is positioned
            click_bold_button(driver)
                
            time.sleep(1)  # Delay to ensure each format is applied

def apply_font_size(driver, element, size):
    # Apply the font size using JavaScript
    script = f"arguments[0].style.fontSize = '{size}px';"
    driver.execute_script(script, element)

def move_cursor_to_end_of_text(driver, element):
    script = """
    var range = document.createRange();
    range.selectNodeContents(arguments[0]);
    range.setEnd(arguments[0].lastChild, arguments[0].lastChild.length);
    var sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    """
    driver.execute_script(script, element)

def click_bold_button(driver):
    bold_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "div[data-id='btn_RTF_Bold']"))
    )
    driver.execute_script("arguments[0].click();", bold_button)
    
def send_text_to_chat(driver, search_input_text, lesson_datas):
    # Open the webpage
    driver.get("https://chat.zalo.me/")
    
    # Wait for the search input element to be present
    search_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "contact-search-input"))
    )
    
    # Interact with the input element
    search_input.send_keys(search_input_text)
    
    # Wait for the container to be present
    container = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "ReactVirtualized__Grid__innerScrollContainer"))
    )
    
    # Locate the second child element using XPath
    second_child = container.find_element(By.XPATH, "./div[2]")
    
    # Perform a click on the second child element
    second_child.click()
    WebDriverWait(driver, 10).until(
        lambda d: d.find_element(By.CSS_SELECTOR, "div[data-id='div_RTF_Menu']").is_displayed()
    )
    
    format_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "div[data-id='div_RTF_Menu']"))
    )
    format_button.click()
    WebDriverWait(driver, 10).until(
        lambda d: d.find_element(By.CSS_SELECTOR, "div[contenteditable='true']").is_displayed()
    )
    
    # Wait for the new chat input element to be present
    chat_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div[contenteditable='true']"))
    )
    
    # Join the lesson text into a single string
    lesson_text = "\n".join(lesson_datas)
    
    # Copy the formatted text to clipboard
    pyperclip.copy(lesson_text)
    
    # Focus on the chat input and paste the formatted text
    chat_input.click()
    chat_input.send_keys(Keys.CONTROL, 'v')

def extract_lesson_text(docx_file, lesson_number):
    # Load the document
    doc = Document(docx_file)
    
    # Initialize variables to store the lesson text
    lesson_text = []
    capture = False
    start_marker = f"Start lesson {lesson_number}"
    end_marker = f"End lesson {lesson_number}"
    
    # Iterate over each paragraph in the document
    for paragraph in doc.paragraphs:
        paragraph_text = paragraph.text.strip()
        
        # Check for the start and end markers
        if start_marker in paragraph_text:
            capture = True
            continue
        elif end_marker in paragraph_text:
            capture = False
            break
        
        # If within the lesson, add the plain text to the list
        if capture:
            lesson_text.append(paragraph_text)
    
    # Join the lesson text into a single string
    return "\n".join(lesson_text)

def apply_bold(driver):
    bold_button = driver.find_element(By.CSS_SELECTOR, "div[data-id='btn_RTF_Bold']")
    driver.execute_script("arguments[0].click();", bold_button)

def apply_text_size(driver, size):
    size_map = {
        "small": "STR_FORMAT_SMALL",
        "medium": "STR_FORMAT_MEDIUM",
        "large": "STR_FORMAT_LARGE",
        "very large": "STR_FORMAT_EXLARGE"
    }
    
    # Locate the size option element
    size_selector = f"div[data-translate-inner='{size_map[size]}']"
    size_option = driver.find_element(By.CSS_SELECTOR, size_selector)
    
    # Get the parent element of the size option
    parent_element = size_option.find_element(By.XPATH, "..")
    
    # Click the parent element
    driver.execute_script("arguments[0].click();", parent_element)

def click_with_js(driver, element):
    driver.execute_script("arguments[0].click();", element)

def clear_selection(driver):
    # JavaScript to clear the text selection
    script = "window.getSelection().removeAllRanges();"
    driver.execute_script(script)

def main(search_input_text, file_path, sheet_name, lesson_number, class_performance_text, homework_result_text, deadline_text, next_requirement_text):
    driver = setup_driver()
    
    # Parameters for the function
    # search_input_text = "Test tool nxbh"
    # file_path = "data.xlsx"
    # sheet_name = "PTA"
    # lesson_number = 9
    # class_performance_text = """"""
    # homework_result_text = """- Trí Cường và Trí Khi��m làm bài còn sơ sài
    # - Minh Tường, Thái An, Nam Khánh làm bài tốt, chỉnh chu"""
    # deadline_text = "- Push code lên github trước ngày 28/11/2024"
    # next_requirement_text = """- Hoàn thành giao diện đăng nhập và đăng ký, hoàn thiện ít nhất 1 màn hình chính"""
    
    # Extract all text with formatting from the Word document
    data = read_excel.read_lesson_and_format_from_excel(file_path, sheet_name, lesson_number)
    if class_performance_text:
        data.append(read_excel.extra_label["class_performance"])
        data.append(class_performance_text)  
    if homework_result_text:
        data.append(read_excel.extra_label["homework_result"])
        data.append(homework_result_text)
    if next_requirement_text:
        data.append(read_excel.extra_label["next_requirement"])
        data.append(next_requirement_text)
    if deadline_text:
        data.append(read_excel.extra_label["deadline"])
        data.append(deadline_text)
    
    # Send the formatted text to the chat
    send_text_to_chat(driver, search_input_text, data)
    time.sleep(1)
    
    format_each_line(driver)
    
    # Keep the browser open
    input("Press Enter to close the browser...")
    
    # Close the driver
    driver.quit()

# if __name__ == "__main__":
#     main()
