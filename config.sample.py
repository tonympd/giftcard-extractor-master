import os
directory = os.path.dirname(os.path.abspath(__file__))

# Windows or Mac, uncomment correct line
# If drivers fail, update drivers here: https://chromedriver.storage.googleapis.com/index.html
# Place driver in same folder as this file
file = 'chromedriver.exe' #Windows
#file = 'chromedriver.dmg' #Mac

CHROMEDRIVER_PATH = os.path.join(directory, file)

# Email Settings (These should not need to be changed if using gmail)
IMAP_HOST = 'imap.gmail.com'
IMAP_PORT = 993
IMAP_SSL = True

# Change these to your gmail username and password and the Folder or Label your gift cards are in
IMAP_USERNAME = 'xxxxxxxxxxxxxxxxx'
IMAP_PASSWORD = 'xxxxxxxxxxxxxx'
FOLDER = 'xxxxxxxxxxx'

# List email addresses we should look for within FOLDER defined above
FROM_EMAILS = ['Costco.com@memberedelivery.com', 'gifts@paypal.com', 'no-reply@samsungpay.com', 'info@newegg.com', 'customerservice@giftcardmall.com', 'DoNotReply.Staples@blackhawk-net.com']

# Flag to save screenshots in screenshot folder if available (not all cards support this)
SAVE_SCREENSHOTS = False

# CSV Output Formats:
# TCB: card_number, card_pin, card_amount
# GCW: card_amount, card_number, card_pin
CSV_OUTPUT_FORMAT = "TCB"

# Flag to enable debug messages
DEBUG = False

# PPDG Gift Card Settings (Sometimes these need to be changed for specific brands)
PPDG_CARD_AMOUNT = '//*[@id="app"]/div/div/div/div/section/div/div[1]/div[2]/div/dl[1]/dd'
PPDG_CARD_NUMBER = '//*[@id="app"]/div/div/div/div/section/div/div[1]/div[2]/div/dl[2]/dd'
PPDG_CARD_PIN = '//*[@id="app"]/div/div/div/div/section/div/div[1]/div[2]/div[2]/dl[3]/dd'

