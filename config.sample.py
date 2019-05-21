import os
directory = os.path.dirname(os.path.abspath(__file__))

file = 'chromedriver.exe' #Windows
#file = 'chromedriver.dmg' #Mac

CHROMEDRIVER_PATH = os.path.join(directory, file)

# Email Settings
IMAP_HOST = 'imap.gmail.com'
IMAP_PORT = 993
IMAP_SSL = True
IMAP_USERNAME = 'xxxxxxxxxxxxxxxxx'
IMAP_PASSWORD = 'xxxxxxxxxxxxxx'

FOLDER = 'xxxxxxxxxxx'

FROM_EMAILS = ['Costco.com@memberedelivery.com','gifts@paypal.com', 'no-reply@samsungpay.com', 'info@newegg.com', 'customerservice@giftcardmall.com', 'DoNotReply.Staples@blackhawk-net.com']

# Gift Card Settings
PPDG_CARD_AMOUNT = '//*[@id="app"]/div/div/div/div/section/div/div[1]/div[2]/div/dl[1]/dd'
PPDG_CARD_NUMBER = '//*[@id="app"]/div/div/div/div/section/div/div[1]/div[2]/div/dl[2]/dd'
PPDG_CARD_PIN = '//*[@id="app"]/div/div/div/div/section/div/div[1]/div[2]/div[2]/dl[3]/dd'

SAVE_SCREENSHOTS = False

DEBUG = False

# CSV Output Formats:
# TCB: card_number, card_pin, card_amount
# GCW: card_amount, card_number, card_pin
CSV_OUTPUT_FORMAT = "TCB"

