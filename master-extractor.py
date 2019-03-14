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
from time import sleep
import config
import time
import json


def parse_activationspot(egc_link):

    link_type = 'activationspot'

    # Open the link in the browser
    browser.get(egc_link['href'])

    # first set of xpath of most spay gc
    try:
        card_brand = browser.find_element_by_xpath('//*[@id="retailerName"]').get_attribute("value")
    except:
        card_brand = None

    # Get Card Information
    if card_brand == 'GameStop':
        try:
            card_amount = browser.find_element_by_xpath('//*[@id="main"]/h1/span').text.replace('$', '').replace(
                'USD', '').strip() + '.00'
            card_number = browser.find_element_by_xpath('//*[@id="cardNumber2"]').text.replace(" ", "")
            card_pin = browser.find_element_by_xpath('//*[@id="main"]/div[4]/p[2]/span').text
        except:
            input('Couldnt get card info for ' + card_brand + '. Please let h4xdaplanet or tony know')
            raise

    # Normal Card
    elif card_brand is not None:
        try:
            card_amount = browser.find_element_by_xpath('//*[@id="main"]/div[1]/div[2]/h2').text.replace(
                '$','').strip() + '.00'
            card_number = browser.find_element_by_xpath('//*[@id="cardNumber2"]').text.replace(" ", "")
            card_pin = browser.find_element_by_xpath('//*[@id="main"]/div[2]/div[2]/p[2]/span').text
        except:
            input('Couldnt get card info for ' + card_brand + '. Please let h4xdaplanet or tony know')
            raise

    # Wayfair or Columbia GC
    elif card_brand is None:

        try:
            card_brand = browser.find_element_by_xpath('//*[@id="main"]/h1/strong').text.strip().split()[1]
        except:
            pass

        if card_brand is not None:
            try:
                card_amount = browser.find_element_by_xpath('//*[@id="amount"]').text.replace('$','').strip() + '.00'
            except:
                pass

            try:
                card_amount = browser.find_element_by_xpath('//*[@id="main"]/div[1]/div[2]/h2').text.replace(
                    '$','').strip() + '.00'
            except:
                input('Couldnt get card info for ' + card_brand + '. Please let h4xdaplanet or tony know')
                raise

            card_number = browser.find_element_by_xpath('//*[@id="main"]/div[2]/div[2]/p/span').text
            card_pin = browser.find_element_by_xpath('//*[@id="main"]/div[2]/div[2]/p[2]/span').text

        # Apple Bees
        elif card_brand is None:
            try:
                card_amount = browser.find_element_by_xpath('//*[@id="amount"]/span').text.replace('$','').strip()
                card_number = browser.find_element_by_xpath('//*[@id="cardNumber2"]').text
                card_pin = browser.find_element_by_xpath('//*[@id="securityCode"]').text
            except:
                input('Couldnt get card info. Please let h4xdaplanet or tony know')
                raise

    # Ensure there is a pin number
    if len(card_pin) == 0:
        card_pin = "N/A"

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

    # Get card details
    try:
        card_brand = browser.find_element_by_xpath('//*[@id="main"]/div[1]/div[1]/h1').text
        if 'Best Buy' in card_brand:
            card_brand = 'Best Buy'
    except:
        card_brand = None

    try:
        card_amount = browser.find_element_by_xpath('//*[@id="main"]/div[1]/div[1]/h1').text
        card_amount = re.search('\$(?P<value>\d*)', card_amount).group(0).replace('$','')
    except:
        card_amount = None

    if card_amount is None:
        try:
            card_amount = browser.find_element_by_xpath('//*[@id="amount"]').text
        except:
            card_amount = None

    card_number = browser.find_element_by_xpath('//*[@id="cardNumber2"]').text.replace(" ", "")

    try:
        card_pin = browser.find_element_by_xpath('//*[@id="Span2"]').text.replace(" ", "")
    except:
        card_pin = None

    if card_pin is None:
        try:
            card_pin = browser.find_element_by_xpath('//*[@id="cardInfo"]/p[3]/span').text.replace(" ", "")
        except:
            card_pin = None

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
    card_pin = browser.find_elements_by_xpath(config.PPDG_CARD_PIN)
    if len(card_pin) > 0:
        card_pin = browser.find_element_by_xpath(config.PPDG_CARD_PIN).text
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

for from_email in config.FROM_EMAILS:

    print("---> Processing Gift Cards from {}".format(from_email))

    # Search for matching emails
    status, messages = mailbox.search(None, '(FROM {})'.format(from_email))

    if status == "OK":
        # Convert the result list to an array of message IDs
        messages = messages[0].split()

        if len(messages) < 1:
            # No matching messages, stop
            print("No matching messages found for {}, nothing to do.".format(from_email))

        else:
            # Open the CSV for writing
            with open('cards_' + datetime.now().strftime('%m-%d-%Y_%H%M%S') + '.csv', 'w', newline='') as csv_file:
                # Start the browser and the CSV writer
                browser = webdriver.Chrome(config.CHROMEDRIVER_PATH)
                csv_writer = csv.writer(csv_file)

                # For each matching email...
                for msg_id in messages:
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
                            msg_html = msg.get_payload(1).get_payload(decode=True)

                        # Parse the message
                        msg_parsed = BeautifulSoup(msg_html, 'html.parser')

                        # PPDG
                        if (msg_parsed.find("a", text="View My Code") or
                            msg_parsed.find("a", text="Unwrap Your Gift")) is not None:

                            egc_link = msg_parsed.find("a", text="View My Code") or \
                                       msg_parsed.find("a", text="Unwrap Your Gift")
                            gift_card = parse_ppdg(egc_link)

                        # Samsung Pay
                        elif msg_parsed.select_one("a[href*=activationspot]") is not None:
                            egc_link = msg_parsed.select_one("a[href*=activationspot]")
                            gift_card = parse_activationspot(egc_link)

                        # Newegg
                        elif msg_parsed.find_all("a", text="View and Print the card") is not None:
                            egc_links = msg_parsed.find_all("a", text=" View and Print the card ")

                            gift_cards = []
                            if egc_links is not None:
                                for egc_link in egc_links:
                                    gift_card = parse_newegg(egc_link)
                                    gift_cards.append(gift_card)
                                    time.sleep(3)

                        # Kroger
                        elif(msg_parsed.find_all("a", text="Redeem your eGift") or
                             msg_parsed.find_all("a", text="Click to Access eGift")) is not None:

                            egc_links = msg_parsed.find_all("a", text="Redeem your eGift") or \
                                        msg_parsed.find_all("a", text="Click to Access eGift")

                            gift_cards = []
                            if egc_links is not None:
                                for egc_link in egc_links:
                                    gift_card = parse_kroger(egc_link)
                                    gift_cards.append(gift_card)
                                    time.sleep(3)
                        else:
                            print("ERROR: Couldn't determine gift card type")

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

                    time.sleep(8)

                # Close the browser
                browser.close()
                print("")
                print("Thank you, come again!")
                print("")
    else:
        print("FATAL ERROR: Unable to fetch list of messages from server.")
