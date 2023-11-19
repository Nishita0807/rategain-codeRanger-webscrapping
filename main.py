import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import time
import subprocess  # Import the subprocess module

url = "https://rategain.com/blog/"

# Set up the web driver in headless mode
options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--blink-settings=imagesEnabled=false')  # Disable image loading
driver = webdriver.Chrome(options=options)

try:
    # Initialize an empty list to store data
    blog_data = []

    # Record the start time
    start_time = time.time()

    # Iterate through pagination pages
    page_num = 1
    while True:
        driver.get(url + f"page/{page_num}/")

        # Wait for the content to load (adjust the timeout as needed)
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CLASS_NAME, 'wrap')))

        # Get the page source after the JavaScript has executed
        page_source = driver.page_source

        # Use BeautifulSoup to parse the HTML
        soup = BeautifulSoup(page_source, 'html.parser')

        # Find all div elements with class 'wrap'
        wrap_divs = soup.select('div.wrap')

        # Extract background-image URLs and blog titles, and add them to the list
        for wrap_div in wrap_divs:
            img_div = wrap_div.select_one('div.img')
            content_div = wrap_div.select_one('div.content')
            date_div = wrap_div.select_one('div.bd-item')
            likes_div = wrap_div.select_one('a.zilla-likes span')

            background_image_url = img_div.find('a')['data-bg'] if img_div and img_div.find('a') else 'no-image'
            blog_title = content_div.find('h6').find('a').text if content_div and content_div.find('h6') else 'no-title'
            blog_date = date_div.find('span').text.strip() if date_div and date_div.find('span') else 'no-date'
            blog_likes = likes_div.text.strip() if likes_div and likes_div.text else 'no-likes'

            blog_data.append({'Blog Title': blog_title, 'Blog Date': blog_date, 'Blog Image URL': background_image_url, 'Blog Likes Count': blog_likes})

        # Find the next page link
        next_page_link = soup.select_one('a.next.page-numbers')

        # Break the loop if there is no next page
        if not next_page_link:
            break

        # Increment page number for the next iteration
        page_num += 1

    # Record the end time
    end_time = time.time()

    # Calculate the time taken
    elapsed_time = end_time - start_time
    print(f"Time taken to fetch data: {elapsed_time} seconds")

    # Convert the list of dictionaries to a DataFrame
    df = pd.DataFrame(blog_data)

    # Save the DataFrame to an Excel file with auto-sized columns
    with pd.ExcelWriter('blog_data.xlsx', engine='openpyxl') as writer:
        df.to_excel(writer, index=False, engine='openpyxl')
        writer.sheets['Sheet1'].column_dimensions['A'].width = 50  # Set the width for column A (adjust as needed)
        writer.sheets['Sheet1'].column_dimensions['B'].width = 20
        writer.sheets['Sheet1'].column_dimensions['C'].width = 100
        writer.sheets['Sheet1'].column_dimensions['D'].width = 30

    # Open the Excel file with the default program
    subprocess.run(['start', 'blog_data.xlsx'], shell=True)

except Exception as e:
    print(f"An error occurred: {e}")

finally:
    driver.quit()
