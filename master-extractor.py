import os
import sys
import email.utils
import re
import csv
from PIL import Image
from datetime import datetime
from imaplib import IMAP4, IMAP4_SSL
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from time import sleep
import config
import time
import json


def parse_activationspot(egc_link):

    link_type = 'activationspot'

    # Open the link in the browser
    browser.get(egc_link['href'])

    card_parsed = BeautifulSoup(browser.page_source, 'html.parser')

    gcm_format = False

    # First find if its typical format
    if card_parsed.find("input", id="retailerName") is not None:
        card_brand = card_parsed.find("input", id="retailerName")['value']
        gcm_format = True

    # Uber, Xbox
    elif card_parsed.find("strong", {"class": "ribbon-content"}) is not None:
        card_brand = card_parsed.find("strong", {"class": "ribbon-content"})

        if "Uber" in card_brand.text:
            card_brand = "Uber"

        elif "Xbox" in card_brand.text:
            card_brand = "Xbox"

        elif "Chipotle" in card_brand.text:
            card_brand = "Chipotle"

    # Staples
    elif card_parsed.find("input", id="Hidden2") is not None:
        card_brand = card_parsed.find("input", id="Hidden2")['value'].replace("®","")

    # AppleBee
    elif card_parsed.find("h1", {"class": "ribbon"}) is not None:
        print("HELLO")
        exit()
        card_brand = card_parsed.find("h1", {"class": "ribbon"}).text.replace(" eGift Card", "").replace("Your ","")

    # Childrens Place
    elif card_parsed.find("div", {"class": "showCard"}) is not None:
        card_brand = card_parsed.find("div", {"class": "showCard"}).find("img")['alt'].replace("'s eGift Card", "")

    # xbox
    elif card_parsed.find("div", {"class": "showCard"}) is not None:
        card_brand = card_parsed.find("h1", {"class": "ribbon"}).text.replace("Your ","")

    else:
        print("Unknown card brand for {}".format(link_type))

    if gcm_format:
        if card_parsed.find("input", id="cardNumber") is not None:
            card_number = card_parsed.find("input", id="cardNumber")['value']
        else:
            card_number = "N/A"

        if card_parsed.find("input", id="pinNumber") is not None:
            card_pin = card_parsed.find("input", id="pinNumber")['value']
        else:
            card_pin = "N/A"

        if card_brand == 'Best Buy':
            header = card_parsed.find("div", {"class": "headingText"}).find("h1").text
            match = re.search('\$(\d*)', header)
            if match:
                card_amount = match.group(1).strip() + '.00'

        elif card_brand == 'GameStop':
            header = card_parsed.find("h1").find("span", {"class": "red"}).text
            match = re.search('\$(\d*)', header)
            if match:
                card_amount = match.group(1).strip() + '.00'

        elif card_brand =='Kohl\'s':
            card_amount = card_parsed.find("span", {"id": "amount"}).text

        else:
            card_amount = card_parsed.find("div", {"class": "showCardInfo"}).find("h2").text.replace('$','').strip()+'.00'

    elif card_brand == 'Regal Cinemas e-GIFT Card':

        header = card_parsed.find("span", {"id": "egc-amount"}).text
        match = re.search('\$(\d*)', header)
        if match:
            card_amount = match.group(1).strip() + '.00'

        card_number = card_parsed.find("span", id="cardNumber2").text.replace(" ", "")
        card_pin = card_parsed.find("p", id="pin-num").find("span").text

    elif card_brand == 'Applebee':

        card_number = card_parsed.find("span", id="cardNumber2").text
        card_pin = card_parsed.find("span", id="securityCode").text
        card_amount = card_parsed.find("div", id="amount").text.replace("$", "")

    elif card_brand == 'Chipotle':

        card_number = card_parsed.find("span", id="cardNumber2").text.strip().replace(" ", "")
        card_pin = card_parsed.find("div", {"class": "cardNum"}).find_all("span")[1].text.strip()
        card_amount = card_parsed.find("div", id="amount").text.strip().replace("$", "")

    elif card_brand == 'Uber':

        card_number = card_parsed.find("div", {"class": "cardNum"}).find("span").text.strip()
        card_pin = "N/A"
        card_amount = card_parsed.find("div", id="amount").text.strip().replace("$", "")

    elif card_brand == 'Xbox':

        card_number = card_parsed.find("span", id="cardNumber2").text.replace(" ", "").strip()
        card_pin = "N/A"
        card_amount = card_parsed.find("div", id="amount").text.strip().replace("$", "")

    elif card_brand == 'Staples':

        card_number = card_parsed.find("input", id="cardNumber")['value']
        card_pin = card_parsed.find("input", id="Hidden5")['value']
        header = card_parsed.find("span", id="egc-amount").text
        match = re.search('\$(.*)', header)
        if match:
            card_amount = match.group(1).strip()

    elif card_brand == 'The Children\'s Place' or card_brand == 'StubHub':
        card_number = card_parsed.find("span", id="cardNumber2").text.replace(" ", "").strip()
        card_pin = card_parsed.find("div", {"class": "cardNum"}).find_all("span")[1].text
        card_amount = card_parsed.find("div", {"id": "amount"}).text.replace('$','').strip()+'.00'

    else:
        print("Unknown card brand for {}".format(link_type))

    # set redeem_flag to zero to stay compatible with ppdg (effects screen capture)
    redeem_flag = 0

    # Create Gift Card Dictionary
    gift_card = {'type': link_type,
                 'brand': card_brand,
                 'amount': card_amount,
                 'number': card_number,
                 'pin': card_pin,
                 'redeem_flag': redeem_flag}

    return gift_card


def parse_costco(egc_link):

    link_type = 'costco'

    card_amount = re.search('\$(.\d+)', str(egc_link.find_all("p", style="font-size:20px;")[0].contents[0])).group(1)
    card_number = egc_link.find_all("p", style="font-size:20px;")[0].find('a').contents[0]

    if card_number[0] == 'X':
        card_brand = 'iTunes'
        card_pin = 'N/A'
    else:
        card_brand = ''
        card_pin = ''
        print("ERROR: Unsupported gift card type for Costco")

    redeem_flag = 0

    gift_card = {'type': link_type,
                 'brand': card_brand,
                 'amount': card_amount,
                 'number': card_number,
                 'pin': card_pin,
                 'redeem_flag': redeem_flag}

    return gift_card    


def parse_gyft(egc_link):

    link_type = 'gyft'

    # Gyft does some weird stuff with javascript, cant open new gift card in same browser window
    browser2 = webdriver.Chrome(config.CHROMEDRIVER_PATH)
    browser2.get(egc_link['href'])
    time.sleep(2)

    if len(browser2.find_elements_by_xpath('/html/body/main/aside/table/tbody/tr/td[2]/h6[2]')) > 0:
        card_brand = browser2.find_elements_by_xpath('/html/body/main/aside/table/tbody/tr/td[2]/h6[2]')[0].text
        card_amount = browser2.find_elements_by_xpath('/html/body/main/aside/table/tbody/tr/td[2]/h6[1]')[0].text.replace('$','')
        card_number = browser2.find_elements_by_xpath('/html/body/main/aside/div[5]/div/div[2]/div[2]')[0].text
        card_pin = browser2.find_elements_by_xpath('/html/body/main/aside/div[5]/div/div[4]/div[2]')[0].text

    browser2.close()

    if len(card_pin) < 1:
        card_pin = 'N/A'

    # set redeem_flag to zero to stay compatible with ppdg (effects screen capture)
    redeem_flag = 0

    # Create Gift Card Dictionary
    gift_card = {'type': link_type,
                 'brand': card_brand,
                 'amount': card_amount,
                 'number': card_number,
                 'pin': card_pin,
                 'redeem_flag': redeem_flag}

    return gift_card


def parse_kroger(egc_link):

    link_type = 'kroger'

    # Open the link in the browser
    browser.get(egc_link['href'])
    card_parsed = BeautifulSoup(browser.page_source, 'html.parser')

    try:
        card_brand = card_parsed.find("input", id="retailerName")['value'].replace('®', '')
        card_number = card_parsed.find("input", id="cardNumber")['value']
    except TypeError:
        card_number = card_parsed.find("a", id="redeem").text
        if card_number[0] == 'X':
            card_brand = 'iTunes'

    if card_parsed.find("input", id="pinNumber") is not None:
        card_pin = card_parsed.find("input", id="pinNumber")['value']
    else:
        card_pin = "N/A"

    if card_brand == 'Best Buy':
        header = card_parsed.find("div", {"class": "headingText"}).find("h1").text
        match = re.search('\$(\d*)', header)
        if match:
            card_amount = match.group(1).strip() + '.00'
    elif card_brand == 'iTunes':
        description = card_parsed.find("div", {"class": "cardNum"}).find("p", {"class": "large"}).text
        card_amount = re.search('\$(\d*)', description).group(1).strip() + '.00'
    else:
        card_amount = card_parsed.find("div", {"class": "showCardInfo"}).find("h2").text.replace('$', '').strip()+'.00'

    # set redeem_flag to zero to stay compatible with ppdg (effects screen capture)
    redeem_flag = 0

    # Create Gift Card Dictionary
    gift_card = {'type': link_type,
                 'brand': card_brand,
                 'amount': card_amount,
                 'number': card_number,
                 'pin': card_pin,
                 'redeem_flag': redeem_flag}

    return gift_card


def parse_newegg(egc_link):

    link_type = 'newegg'

    # Open the link in the browser
    browser.get(egc_link['href'])

    if (len(browser.find_elements_by_id('lblHumanBarcodeReadable')) > 0 and
            len(browser.find_elements_by_id('imgCertBarCode')) > 0):
        card_brand = 'Regal'
        card_amount = '{0:.2f}'.format(float(browser.find_element_by_id('lblCertAmount').text.replace('$', '').strip()))
        card_number = browser.find_element_by_id('lblHumanBarcodeReadable').text.replace(' ', '').strip()
        card_pin = browser.find_element_by_id('lblPin').text.strip()
        #order_recipient = browser.find_element_by_id('lblRecipient').text.strip()
        redeem_flag = 1

    elif len(browser.find_elements_by_id('imgCertBarCode')) > 0:
        card_brand = 'Nike'
        card_amount = '{0:.2f}'.format(float(browser.find_element_by_id('lblCertAmount').text.replace('$', '').strip()))
        nike_barcode_src = browser.find_element_by_id('imgCertBarCode').get_attribute('src')
        card_number_pos = nike_barcode_src.find('&CBID=')
        card_number = nike_barcode_src[card_number_pos + 6:card_number_pos + 25]
        card_pin = browser.find_element_by_id('lblPin').text.strip()
        #order_recipient = browser.find_element_by_id('lblRecipient').text.strip()
        redeem_flag = 1

    elif len(browser.find_elements_by_id('ids-configuration')) > 0:
        card_element = browser.find_element_by_id('ids-configuration')
        card_data = json.loads(card_element.get_attribute('data-certificate'))
        card_configuration = json.loads(card_element.get_attribute('data-configuration'))
        card_brand = card_configuration[0]['settings']['brandName'].replace(u"\u00AE", '').replace("'", '')
        card_amount = '{0:.2f}'.format(card_data['CurrentBalance'])
        card_number = card_data['CardNumber']
        card_pin = str(card_data['Pin'])
        redeem_flag = 0

    if len(card_pin) < 1:
        card_pin = 'N/A'

    # Create Gift Card Dictionary
    gift_card = {'type': link_type,
                 'brand': card_brand,
                 'amount': card_amount,
                 'number': card_number,
                 'pin': card_pin,
                 'redeem_flag': redeem_flag}

    return gift_card


def parse_ppdg(egc_link):

    link_type = 'PPDG'

    # Open the link in the browser
    browser.get(egc_link['href'])

    # Captcha Testing
    card_type_exists = browser.find_elements_by_xpath(
        '//*[@id="app"]/div/div/div/div/section/div/div[3]/div[2]/div/h2[1]')

    if card_type_exists:
        pass
    else:
        input("Press Enter to continue...")

    # Get the card type
    card_brand = browser.find_element_by_xpath(
        '//*[@id="app"]/div/div/div/div/section/div/div[3]/div[2]/div/h2[1]').text.strip()
    card_brand = re.compile(r'(.*) Terms and Conditions').match(card_brand).group(1)

    # Get the card amount
    card_amount = browser.find_element_by_xpath(config.PPDG_CARD_AMOUNT).text.replace('$', '').strip() + '.00'

    # Get the card number
    card_number = browser.find_element_by_xpath(config.PPDG_CARD_NUMBER).text

    # Get the card PIN
    try:
        card_pin = browser.find_elements_by_xpath("//*[text()='PIN']/following-sibling::dd")
    except NoSuchElementException:
        card_pin = browser.find_elements_by_xpath(config.PPDG_CARD_PIN)

    if len(card_pin) > 0:
        card_pin = card_pin[0].text
    else:
        card_pin = "N/A"

    # Look for Redeem button as this effects crop size
    redeem = browser.find_elements_by_id("redeem_button")
    if len(redeem) > 0:
        redeem_flag = 1
    else:
        redeem_flag = 0

    # Create Gift Card Dictionary
    gift_card = {'type': link_type,
                 'brand': card_brand,
                 'amount': card_amount,
                 'number': card_number,
                 'pin': card_pin,
                 'redeem_flag': redeem_flag}

    return gift_card



# def parse_staples(egc_link):
#
#     link_type = 'staples'
#
#     # Open the link in the browser
#     browser.get(egc_link['href'])
#     card_parsed = BeautifulSoup(browser.page_source, 'html.parser')
#
#     card_brand = card_parsed.find("input", id="retailerName")['value'].replace('®', '')
#     card_number = card_parsed.find("input", id="cardNumber")['value']
#
#     if card_parsed.find("input", id="pinNumber") is not None:
#         card_pin = card_parsed.find("input", id="pinNumber")['value']
#     else:
#         card_pin = "N/A"
#
#
#     card_amount = card_parsed.find("div", {"class": "showCardInfo"}).find("h2").text.replace('$', '').strip()+'.00'
#
#     # set redeem_flag to zero to stay compatible with ppdg (effects screen capture)
#     redeem_flag = 0
#
#     # Create Gift Card Dictionary
#     gift_card = {'type': link_type,
#                  'brand': card_brand,
#                  'amount': card_amount,
#                  'number': card_number,
#                  'pin': card_pin,
#                  'redeem_flag': redeem_flag}
#
#     return gift_card



def take_screenshot(gift_card):

    # Create a directory for screenshots if it doesn't already exist
    screenshots_dir = os.path.join(os.getcwd(), 'screenshots')
    if not os.path.exists(screenshots_dir):
        os.makedirs(screenshots_dir)

    supported = True

    # Save a screenshot
    if gift_card['type'] == 'PPDG':
        element = browser.find_element_by_xpath('//*[@id="app"]/div/div/div/div/section/div/div[1]/div[2]')
    elif gift_card['type'] == 'activationspot':
        element = browser.find_element_by_xpath('//*[@id="main"]')
    elif gift_card['type'] == 'kroger' and gift_card['brand'] == 'Best Buy':
        element = browser.find_element_by_xpath('//*[@id="main"]/div[1]')
    elif gift_card['type'] == 'kroger':
        element = browser.find_element_by_xpath('//*[@id="main"]/div[2]')
    else:
        print("Screen Shots are not currently supported for {}".format(gift_card['type']))
        supported = False

    if supported:
        location = element.location

        size = element.size
        screenshot_name = os.path.join(screenshots_dir, gift_card['number'] + '.png')
        screenshot_name_new = os.path.join(screenshots_dir, gift_card['number'] + '.jpg')
        browser.save_screenshot(screenshot_name)

        im = Image.open(screenshot_name)
        left = location['x']
        top = location['y']
        right = location['x'] + size['width']

        if gift_card['redeem_flag'] == 1:
            bottom = location['y'] + size['height'] - 80
        else:
            bottom = location['y'] + size['height']

        im = im.crop((left, top, right, bottom))
        im.convert('RGB').save(screenshot_name_new)
        sleep(0.1)
        os.remove(screenshot_name)


def write_card(gift_card):

    # Write the details to the CSV
    if config.CSV_OUTPUT_FORMAT == "TCB":
        csv_writer.writerow([gift_card['number'], gift_card['pin'], gift_card['amount']])
    elif config.CSV_OUTPUT_FORMAT == "GCW":
        csv_writer.writerow([gift_card['amount'], gift_card['number'], gift_card['pin']])
    else:
        print("ERROR: Invalid output format, please specify TCB or GCW in config.py")

    # Print out the details to the console
    print("{}: {},{},{}".format(gift_card['brand'], gift_card['number'], gift_card['pin'], gift_card['amount']))


##################
# Main Code Logic
##################
if os.name == 'nt':
    sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf8', buffering=1)

# Connect to the server
if config.IMAP_SSL:
    mailbox = IMAP4_SSL(host=config.IMAP_HOST, port=config.IMAP_PORT)
else:
    mailbox = IMAP4(host=config.IMAP_HOST, port=config.IMAP_PORT)

# Log in and select the configured folder
mailbox.login(config.IMAP_USERNAME, config.IMAP_PASSWORD)
mailbox.select(config.FOLDER)

# Grab epoch timestamp for SINGLE_CSV_FILE
epoch = str(time.time()).split('.')[0]

for from_email in config.FROM_EMAILS:

    if config.DEBUG:
        print("---> Processing Gift Cards from {}".format(from_email))

    # Search for matching emails
    status, messages = mailbox.search(None, '(FROM {})'.format(from_email))

    if status == "OK":
        # Convert the result list to an array of message IDs
        messages = messages[0].split()

        if len(messages) < 1:
            # No matching messages, stop
            if config.DEBUG:
                print("No matching messages found for {}, nothing to do.".format(from_email))

        else:
            # Open the CSV for writing
            if config.SINGLE_CSV_FILE:
                csv_filename = 'cards_' + epoch + '.csv'
            else:
                csv_filename = from_email + '_cards_' + datetime.now().strftime('%m-%d-%Y_%H%M%S') + '.csv'

            with open(csv_filename, 'a', newline='') as csv_file:

                # Start the browser and the CSV writer
                browser = webdriver.Chrome(config.CHROMEDRIVER_PATH)
                browser.set_page_load_timeout(10)
                csv_writer = csv.writer(csv_file)

                # For each matching email...
                for msg_id in messages:
                    if config.DEBUG:
                        print("--> Processing message id {}...".format(msg_id.decode('UTF-8')))

                    # Fetch it from the server
                    status, data = mailbox.fetch(msg_id, '(RFC822)')

                    if status == "OK":
                        # Convert it to an Email object
                        msg = email.message_from_bytes(data[0][1])

                        # Get the HTML body payload.
                        if not msg.is_multipart():
                            msg_html = msg.get_payload(decode=True)
                        else:
                            try:
                                msg_html = msg.get_payload(1).get_payload(decode=True)
                            except IndexError:
                                msg_html = msg.get_payload(0).get_payload(decode=True)

                        # Parse the message
                        try:
                            msg_parsed = BeautifulSoup(msg_html, 'html.parser')
                        except TypeError: # Kluge to fix issue with encoding on Staples
                            msg_html = str(msg.get_payload(0).get_payload()[0])
                            msg_parsed = BeautifulSoup(msg_html, 'html.parser')


                        # Determine Message type to parse accordingly
                        # PPDG
                        if (msg_parsed.find("a", text="View My Code") or
                            msg_parsed.find("a", text="Unwrap Your Gift")) is not None:

                            if config.DEBUG:
                                print('PPDG')

                            egc_link = msg_parsed.find("a", text="View My Code") or \
                                       msg_parsed.find("a", text="Unwrap Your Gift")
                            gift_card = parse_ppdg(egc_link)
                            time.sleep(60)

                        # Samsung Pay
                        elif msg_parsed.select_one("a[href*=activationspot]") is not None:
                            if config.DEBUG:
                                print('SPAY')

                            egc_link = msg_parsed.select_one("a[href*=activationspot]")
                            gift_card = parse_activationspot(egc_link)

                        # Gyft
                        elif msg_parsed.select_one("a[href*=gyft]") is not None:
                            if config.DEBUG:
                                print('GYFT')

                            egc_link = msg_parsed.select_one("a[href*=gyft]")
                            gift_card = parse_gyft(egc_link)

                        # Newegg
                        elif len(msg_parsed.find_all("a", text="View and Print the card")) > 0:
                            if config.DEBUG:
                                print('Newegg')

                            egc_links = msg_parsed.find_all("a", text=" View and Print the card ")

                            gift_cards = []
                            if egc_links is not None:
                                for egc_link in egc_links:
                                    gift_card = parse_newegg(egc_link)
                                    gift_cards.append(gift_card)
                                    time.sleep(3)

                        # Costco
                        elif (len(msg_parsed.find_all("div", style="cardStuff")) > 0):
                            if config.DEBUG:
                                print('Costco')

                            egc_links = msg_parsed.find_all("p", id='primaryCode')

                            gift_cards = []
                            if egc_links is not None:
                                for egc_link in egc_links:
                                    gift_card = parse_costco(egc_link)
                                    gift_cards.append(gift_card)

                        # Kroger
                        elif (len(msg_parsed.find_all("a", text=re.compile('.*Redeem your eGift'))) > 0) or \
                             (len(msg_parsed.find_all("a", text="Click to Access eGift")) > 0) or \
                             (len(msg_parsed.find_all("a", text=re.compile('.*Click to View'))) > 0):
                            if config.DEBUG:
                                print('Kroger')

                            egc_links = msg_parsed.find_all("a", text=re.compile('.*Redeem your eGift'))

                            if len(egc_links) == 0:
                                egc_links = msg_parsed.find_all("a", text="Click to Access eGift")
                            if len(egc_links) == 0:
                                egc_links = msg_parsed.find_all("a", text=re.compile('.*Click to View'))

                            gift_cards = []
                            if egc_links is not None:
                                for egc_link in egc_links:
                                    gift_card = parse_kroger(egc_link)
                                    gift_cards.append(gift_card)
                                    time.sleep(3)

                        # Staples
                        elif len(msg_parsed.find_all("a", text=re.compile('.*View Gift')) > 0):

                            if config.DEBUG:
                                print('Staples')

                            egc_links = msg_parsed.find_all("a", text=re.compile('View Gift'))

                            gift_cards = []
                            if egc_links is not None:
                                for egc_link in egc_links:
                                    gift_card = parse_kroger(egc_link)
                                    gift_cards.append(gift_card)
                                    time.sleep(3)

                        # Amazon
                        elif msg_parsed.select_one("a[href*=amazon]") is not None:
                            if config.DEBUG:
                                print('Amazon')

                            egc_link = msg_parsed.select_one("a[href*=amazon]")
                            gift_card = parse_activationspot(egc_link)

                        else:
                            print(msg_parsed)
                            print("ERROR: Couldn't determine gift card type")
                            exit()

                        # Write Card to CSV
                        try:
                            if gift_cards is not None:
                                for gift_card in gift_cards:
                                    write_card(gift_card)

                                    # Grab Screenshots
                                    if config.SAVE_SCREENSHOTS:
                                        take_screenshot(gift_card)

                        except NameError:
                            if 'number' in gift_card:
                                write_card(gift_card)

                                # Grab Screenshots
                                if config.SAVE_SCREENSHOTS:
                                    take_screenshot(gift_card)

                    else:
                        print("ERROR: Unable to fetch message {}, skipping.".format(msg_id.decode('UTF-8')))

                    time.sleep(10)

                # Close the browser
                browser.close()

    else:
        print("FATAL ERROR: Unable to fetch list of messages from server.")

print("")
print("Thank you, come again!")
print("")