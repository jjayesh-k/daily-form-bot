import os
import time
import smtplib
from datetime import datetime
from email.message import EmailMessage
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- SECURE CONFIGURATION ---
# We now pull these from the environment variables
SENDER_EMAIL = os.environ.get("jayeshkandar7@gmail.com") 
SENDER_PASSWORD = os.environ.get("EMAIL_APP_PASSWORD")
FORM_URL = "https://forms.office.com/Pages/ResponsePage.aspx?id=DRj7o8QjyUSiQceBCSAr05X2XiTtPspNgAlIHjt75qdUNDJWTVFCSDZaNTM1UFUzRk1MVDZDNkVTRy4u"

# User Data List
users = [
    { "name": "Jayesh Kandar", "p_no": "99989514", "email": "jayeshkandar001@gmail.com" },
    { "name": "Mehul Patil", "p_no": "99999862", "email": "mkpatil5556@gmail.com" },
    { "name": "Anirudh Ajagekar", "p_no": "99989489", "email": "anirudhaajagekar14@gmail.com" }
]

def send_confirmation_email(user_data, screenshot_path):
    msg = EmailMessage()
    msg['Subject'] = f"✅ Attendance Marked: {user_data['name']}"
    msg['From'] = SENDER_EMAIL
    msg['To'] = user_data['email']
    msg.set_content(f"Hello {user_data['name']},\n\nYour daily attendance form was successfully submitted at {datetime.now().strftime('%H:%M')}.\n\nPlease find the confirmation screenshot attached.")

    # Attach the screenshot
    with open(screenshot_path, 'rb') as f:
        file_data = f.read()
        file_name = f.name
    msg.add_attachment(file_data, maintype='image', subtype='png', filename=file_name)

    # Send the email
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(SENDER_EMAIL, SENDER_PASSWORD)
            smtp.send_message(msg)
        print(f"📧 Email sent to {user_data['name']}")
    except Exception as e:
        print(f"❌ Failed to send email: {e}")

def fill_and_notify(user):
    print(f"\n--- Starting process for {user['name']} at {datetime.now().strftime('%H:%M:%S')} ---")
    
    # Setup Headless Chrome
    options = webdriver.ChromeOptions()
    options.add_argument("--headless") 
    options.add_argument("--window-size=1920,1080") # Important for clear screenshots
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 20)
    screenshot_name = f"confirmation_{user['name'].split()[0]}.png"

    try:
        driver.get(FORM_URL)
        time.sleep(3)

        # --- FILL FORM LOGIC ---
        def fill_text(label, value):
            try:
                xpath = f"//span[contains(text(), '{label}')]/ancestor::div[contains(@class, '-question-root')]//input"
                field = driver.find_element(By.XPATH, xpath)
            except:
                field = driver.find_element(By.XPATH, f"//span[contains(text(), '{label}')]/following::input[1]")
            field.clear()
            field.send_keys(value)

        fill_text("Full Name", user['name'])
        fill_text("P. No", user['p_no'])
        fill_text("Department", "Pv ops")
        fill_text("Current Date", datetime.now().strftime("%d/%m/%Y"))
        
        # Click Present
        try:
            driver.find_element(By.XPATH, "//span[contains(text(), 'Present')]").click()
        except:
            pass # Sometimes it's already selected or structured differently

        fill_text("In Time", "830")
        fill_text("Out Time", "1700")

        # Submit
        submit_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Submit')]")))
        submit_btn.click()
        
        # Wait for "Thanks" page
        time.sleep(3) 
        
        # Take Screenshot
        driver.save_screenshot(screenshot_name)
        print(f"📸 Screenshot saved: {screenshot_name}")
        
        # Send Email
        send_confirmation_email(user, screenshot_name)

    except Exception as e:
        print(f"❌ Error for {user['name']}: {e}")
    
    finally:
        driver.quit()

# --- MAIN EXECUTION LOOP ---
# This runs immediately when called. 
# To schedule it, use a Cron Job (GitHub Actions) or Windows Task Scheduler to run this file at 5:10 PM.

if __name__ == "__main__":
    # 1. Jayesh (17:10)
    fill_and_notify(users[0])
    
    # Wait 2 minutes for Mehul (17:12)
    print("⏳ Waiting 2 minutes for next user...")
    time.sleep(120) 
    
    # 2. Mehul
    fill_and_notify(users[1])

    # Wait 2 minutes for Anirudh (17:14)
    print("⏳ Waiting 2 minutes for next user...")
    time.sleep(120)

    # 3. Anirudh
    fill_and_notify(users[2])
    
    print("✅ All tasks completed.")