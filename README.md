## Credit
@Churn2burn for ppdg-extractor

# giftcard-extractor
This is a modified version of ppdg-extractor.  The format of the card output is different but adds functionality for multiple from email addresses (useful for forwarded emails).

## Setup ##
Extract (CSV) codes from and screenshot (JPG) PayPal Digital Gifts gift cards. Adapted from https://github.com/stevenmirabito/ code in order to clean up screenshots, save screenshots as .JPG, and to process PIN and non-PIN cards using the same program.

Program specifically configured on Windows - includes ChromeDriver.exe - for gmail.

1) Install the newest version of Python: https://www.python.org/downloads/

2) Open command prompt (cmd) and navigate to this folder. Install the dependencies by running the following commands:
	
	```bash
	 pip3 install -r requirements.txt
    ```
    
3) Rename config.sample.py to config.py and edit the following variables:
	
	a) Change IMAP_USERNAME to your gmail account
	
	b) Change IMAP_PASSWORD to your gmail password
	
	c) Change XPATH locations for card_amount, card_number, and card_pin if necessary (will be required if these fields come back as N/A in the console. XPATH can be found by inspecting element in Chrome, right clicking on the item in the elements console, Copy, "Copy XPATH".
	
	d) Create a gmail label for the cards you would like to extract and change FOLDER to this label. This label will serve as the processing folder for your cards. When you would like to extract a card, label it with this label, run the extractor, and then move the card to another label. The program can only see cards that reside in this label.
	
	e) Change FROM_EMAILS to any email you want to check the FROM address of.  This is nice if you have forwarded emails.  
	
	f) Note: When logging in for the first time, gmail may block access. You will need to follow the steps in the email to follow to enable less secure applications.
	
4) Double click on MasterExtractor.bat to run the program. You will get .jpg screenshots in the "screenshots" folder and will have a .csv file for each FROM email address ~~with near GCW (GiftCardWiki) submission standards.~~ In order to access the cards, open the .csv file with your notepad program of choice. I prefer Notepad++ (https://notepad-plus-plus.org/download/v7.5.3.html).

5) Enjoy!
