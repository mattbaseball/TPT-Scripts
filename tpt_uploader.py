import streamlit as st
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os

# --- Main Automation Function ---
def fill_tpt_listing(product_info):
    """
    This function launches a browser, navigates to the TPT new product page,
    and fills it out with the provided info.
    """
    st.info("Launching Chrome browser...")

    try:
        # Initialize the Chrome driver automatically
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

        # Go to the TPT login page first.
        # The user MUST log in manually. The script will wait.
        driver.get("https://www.teacherspayteachers.com/login")
        st.warning("Please log in to your TPT account in the new Chrome window. The script will wait for you to continue.")

        # This will hold the browser open until the user is ready.
        # We check if they have navigated away from the login page.
        WebDriverWait(driver, timeout=300).until(
            lambda d: d.current_url != "https://www.teacherspayteachers.com/login"
        )
        
        st.success("Login detected! Navigating to the 'Add New Product' page...")
        
        # Navigate to the page for adding a new digital product
        driver.get("https://www.teacherspayteachers.com/My-Products/New/Digital-Next")

        # --- Fill out the form ---

        # Wait for the main form to be present
        wait = WebDriverWait(driver, 20)
        wait.until(EC.presence_of_element_located((By.ID, "ItemAddForm")))
        
        st.write("Filling in Product Title...")
        # The title input is inside a React component, we'll find it by its 'name' attribute
        title_input = driver.find_element(By.NAME, "data[Item][name]")
        title_input.send_keys(product_info["title"])
        
        # --- File Uploads ---
        # NOTE: File paths must be absolute for Selenium to work reliably.
        
        st.write("Uploading downloadable file...")
        product_file_input = driver.find_element(By.ID, "ItemDigitalProduct")
        product_file_input.send_keys(product_info["product_file"])

        # Wait a moment for the upload to start processing
        time.sleep(5)

        st.write("Uploading main cover thumbnail...")
        # The user must select "Upload thumbnails now" for this to work
        upload_thumbs_radio = driver.find_element(By.ID, "ItemGenerateThumbnail2")
        driver.execute_script("arguments[0].click();", upload_thumbs_radio) # Use JS click for reliability
        time.sleep(1) # wait for the upload boxes to appear

        thumb1_input = driver.find_element(By.ID, "ItemDigitalThumb1")
        thumb1_input.send_keys(product_info["thumbnail_file"])

        # --- Product Description (This is a rich text editor and requires a special approach) ---
        st.write("Filling in product description...")
        # The description editor is inside an iframe. We must switch to it first.
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ItemDescription_ifr")))
        
        # Now inside the iframe, find the body element and type the description
        description_body = driver.find_element(By.ID, "tinymce")
        description_body.send_keys(product_info["description"])

        # IMPORTANT: Switch back to the main page content
        driver.switch_to.default_content()

        # --- Price ---
        st.write("Setting the price...")
        # TPT's price field is complex. We'll find the single license price input.
        # It's better to find it by a more specific selector. Let's assume its name.
        # This selector may need updating if TPT changes their site.
        price_input = wait.until(EC.presence_of_element_located((By.NAME, "data[Item][price]")))
        price_input.clear()
        price_input.send_keys(str(product_info["price"]))


        # --- Final Step ---
        st.success("All fields filled! Please review the listing in the browser and click the final submit button manually.")
        st.balloons()
        
        # The script will keep the browser open for the user to finish
        st.info("The browser window will remain open. You can close it when you are done.")
        while True:
            time.sleep(1)

    except Exception as e:
        st.error(f"An error occurred: {e}")
        st.error("This could be because the page structure on TPT has changed, or an element was not found in time. Please check the browser window for details.")
        if 'driver' in locals():
            driver.quit()


# --- Streamlit App UI ---

st.set_page_config(page_title="TPT Listing Automator", layout="centered")

st.title("Teachers Pay Teachers Listing Filler")
st.markdown("Fill out the form below, and the script will open a browser and enter the data for you.")

with st.form("tpt_product_form"):
    st.header("Product Information")
    product_title = st.text_input("Product Title", placeholder="e.g., Ultimate Grammar Worksheets for Grade 5")
    product_desc = st.text_area("Product Description", height=200, placeholder="Describe your product here. What's included? How can teachers use it?")
    product_price = st.number_input("Price ($)", min_value=0.0, step=0.25, format="%.2f")

    st.header("File Uploads")
    # Using st.file_uploader is for the UI only. We save the file to a temp location to get a path for Selenium.
    uploaded_product_file = st.file_uploader("1. Select your main product file (ZIP, PDF, etc.)")
    uploaded_thumbnail_file = st.file_uploader("2. Select your main cover image (JPG, PNG)")
    
    submitted = st.form_submit_button("Start Automation")

if submitted:
    # --- Validation ---
    if not all([product_title, product_desc, uploaded_product_file, uploaded_thumbnail_file]):
        st.error("Please fill out all fields and upload both files before starting.")
    else:
        # Save uploaded files to a temporary directory so Selenium can access them by path
        temp_dir = "temp_tpt_uploads"
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)

        product_file_path = os.path.join(temp_dir, uploaded_product_file.name)
        with open(product_file_path, "wb") as f:
            f.write(uploaded_product_file.getbuffer())

        thumbnail_file_path = os.path.join(temp_dir, uploaded_thumbnail_file.name)
        with open(thumbnail_file_path, "wb") as f:
            f.write(uploaded_thumbnail_file.getbuffer())
            
        # Create a dictionary to hold all the product info
        product_data = {
            "title": product_title,
            "description": product_desc,
            "price": product_price,
            # Use absolute paths for selenium
            "product_file": os.path.abspath(product_file_path),
            "thumbnail_file": os.path.abspath(thumbnail_file_path)
        }
        
        fill_tpt_listing(product_data)