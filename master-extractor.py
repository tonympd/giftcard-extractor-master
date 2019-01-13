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

    # Search for matching emails
    status, messages = mailbox.search(None, '(FROM {})'.format(from_email))

    if status == "OK":
        # Convert the result list to an array of message IDs
        messages = messages[0].split()
        skip = False

        if len(messages) < 1:
            # No matching messages, stop
            print("No matching messages found for {}, nothing to do.".format(from_email))
            skip = True

        if not skip:
            # Open the CSV for writing
            with open(from_email + '_cards_' + datetime.now().strftime('%m-%d-%Y_%H%M%S') + '.csv', 'w', newline='') as csv_file:
                # Start the browser and the CSV writer
                browser = webdriver.Chrome(config.CHROMEDRIVER_PATH)
                csv_writer = csv.writer(csv_file)

                # Create a directory for screenshots if it doesn't already exist
                screenshots_dir = os.path.join(os.getcwd(), 'screenshots')
                if not os.path.exists(screenshots_dir):
                    os.makedirs(screenshots_dir)

                # For each matching email...
                for msg_id in messages:
                    print("---> Processing message id {}...".format(msg_id.decode('UTF-8')))

                    # Fetch it from the server
                    status, data = mailbox.fetch(msg_id, '(RFC822)')

                    if status == "OK":
                        # Convert it to an Email object
                        msg = email.message_from_bytes(data[0][1])
                        msg_html = msg.get_payload(1).get_payload(decode=True)

                        # Save the email timestamp
                        datetime_received = datetime.fromtimestamp(
                            email.utils.mktime_tz(email.utils.parsedate_tz(msg.get('date'))))

                        # Parse the message
                        msg_parsed = BeautifulSoup(msg_html, 'html.parser')

                        # Find the "View Gift" link
                        egc_link = msg_parsed.find("a", text="View My Code") or msg_parsed.find("a", text="Unwrap Your Gift")
                        if egc_link is not None:
                            # Open the link in the browser
                            browser.get(egc_link['href'])

                            # Get the type of card
                            card_type_exists = browser.find_elements_by_xpath('//*[@id="app"]/div/div/div/div/section/div/div[3]/div[2]/div/h2[1]')

                            if card_type_exists:
                                card_type = browser.find_element_by_xpath('//*[@id="app"]/div/div/div/div/section/div/div[3]/div[2]/div/h2[1]').text.strip()
                                card_type = re.compile(r'(.*) Terms and Conditions').match(card_type).group(1)

                            else:
                                input("Press Enter to continue...")
                                card_type = browser.find_element_by_xpath('//*[@id="app"]/div/div/div/div/section/div/div[3]/div[2]/div/h2[1]').text.strip()
                                card_type = re.compile(r'(.*) Terms and Conditions').match(card_type).group(1)

                            # Get the card amount
                            card_amount = browser.find_element_by_xpath(config.card_amount).text.replace('$', '').strip() + '.00'

                            # Get the card number
                            card_number = browser.find_element_by_xpath(config.card_number).text

                            # Get the card PIN
                            card_pin = browser.find_elements_by_xpath(config.card_pin)
                            if len(card_pin) > 0:
                                card_pin = browser.find_element_by_xpath(config.card_pin).text
                            else:
                                card_pin = "N/A"

                            redeem = browser.find_elements_by_id("redeem_button")
                            if len(redeem) > 0:
                                redeem_flag = 1
                            else:
                                redeem_flag = 0

                            # Save a screenshot
                            element = browser.find_element_by_xpath('//*[@id="app"]/div/div/div/div/section/div/div[1]/div[2]')
                            location = element.location

                            size = element.size
                            screenshot_name = os.path.join(screenshots_dir, card_number + '.png')
                            screenshot_name_new = os.path.join(screenshots_dir, card_number + '.jpg')
                            browser.save_screenshot(screenshot_name)

                            im = Image.open(screenshot_name)
                            left = location ['x']
                            top = location['y']
                            right = location['x'] + size['width']

                            if redeem_flag == 1:
                                bottom = location['y'] + size['height'] - 80
                            else:
                                bottom = location['y'] + size['height']

                            im = im.crop((left, top, right, bottom))
                            im.convert('RGB').save(screenshot_name_new)
                            sleep(0.1)
                            os.remove(screenshot_name)

                            # Write the details to the CSV
                            csv_writer.writerow([card_number, card_pin, card_amount])

                            # Print out the details to the console
                            print("{}: {},{},{}".format(card_type, card_number, card_pin, card_amount))
                        else:
                            print("ERROR: Unable to find eGC link in message {}, skipping.".format(msg_id.decode('UTF-8')))
                    else:
                        print("ERROR: Unable to fetch message {}, skipping.".format(msg_id.decode('UTF-8')))

                    time.sleep(8)

                # Close the browser
                browser.close()
                print("")
                print("Thank you, come again!")
    else:
        print("FATAL ERROR: Unable to fetch list of messages from server.")
