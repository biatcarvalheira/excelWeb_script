import os

# Selenium ChromeDriver
TIMEOUT = 10

# --- chrome ---- #
current_path = os.path.dirname(os.path.abspath(__file__))
chromeDriverFile = "chromedriver"

# --- firefox ---- #
folder_path = os.path.join(current_path, chromeDriverFile)
firefoxDriverFile = 'geckodriver'
firefox_folder_path = os.path.join(current_path, firefoxDriverFile)

