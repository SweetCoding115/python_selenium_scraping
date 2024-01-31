```py
from selenium import webdriver
from selenium.webdriver.common.by import By
# from selenium.webdriver.chrome.options import Options
from selenium.webdriver.firefox.options import Options
from selenium.common.exceptions import NoSuchElementException
import openpyxl
import re


# enable headless mode in Selenium
options = Options()

# headless option for chrome
options.add_argument('--headless=new')

# headless option for firefox
options.add_argument('--headless')

# initialize an instance of the chrome/firefox driver (browser)
# it doesn't matter you write webdriver.Firefox or webdriver.Chrome, 
# you just need to drill options with what you import from selenium.webdriver.firefox.options.
driver_initial = webdriver.Chrome(
    options=options,
)

# visit your www.vehicle-spares.com site
driver_initial.get('https://www.vehicle-spares.com/collections/crash-parts?page=1')

# searching products...
product_count = driver_initial.find_element(By.CSS_SELECTOR, f"html body div.VehicleContainer div.ProdCount span")
product_num = int(re.search(r"\d+", ''.join(product_count.text.split(","))).group())
print(int(product_num),'items found.')

# # searching pages...
```

you need to install its dependencies by running this command on your terminal before you run this python script.

```bash
pip install openpyxl selenium regex
```
then run this command on your terminal to start scraping.

```bash
python scrape_new.py
```
