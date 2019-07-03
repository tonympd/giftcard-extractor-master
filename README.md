## Contributors
- @Churn2burn 
- @Tony 
- @Michael (Kroger)
- @h4xdaplanet (Spay)
- @VoyagingDoc (Newegg)

# giftcard-extractor
This is a modified version of the ppdg-extractor to support the following:
 - Paypal Digital Gifts
 - Samsung Pay
 - Kroger
 - Newegg
 - Staples
 - Costco (Itunes)
 - Amazon

## Setup ##
Extract (CSV) codes from and screenshot (JPG) PayPal Digital Gifts gift cards. Adapted from https://github.com/stevenmirabito/ code in order to clean up screenshots, save screenshots as .JPG, and to process PIN and non-PIN cards using the same program.

Program specifically configured on Windows - includes ChromeDriver.exe - for gmail.

Mac users will need additional configuration steps.

1) Install the newest version of Python: https://www.python.org/downloads/

2) Open command prompt (cmd) and navigate to this folder. Install the dependencies by running the following commands:
	
	```bash
	 pip3 install -r requirements.txt
    ```
    
3) Rename config.sample.py to config.py and edit the following variables:
	
	a) Change IMAP_USERNAME to your gmail account
	
	b) Change IMAP_PASSWORD to your gmail password
	
	c) Create a gmail label for the cards you would like to extract and change FOLDER to this label. This label will serve as the processing folder for your cards. When you would like to extract a card, label it with this label, run the extractor, and then move the card to another label. The program can only see cards that reside in this label.
	
	d) Change FROM_EMAILS to any list of email you want to check the FROM address of.  This is nice if you have forwarded emails or for multiple types of cards.
	
	e) Set SAVE_SCREENSHOTS = False if you do not want screenshots (capitalization matters).  Screenshots are not supported with all card types.
	
	    SAVE_SCREENSHOTS = True
	    SAVE_SCREENSHOTS = False
	
	f) Set CSV_OUTPUT_FORMAT to "TCB" or "GCW" depending on desired output
	
	    TCB: card_number, card_pin, card_amount
	    GCW: card_amount, card_number, card_pin
	    
	g) Set SINGLE_CSV_FILE
	
	    True - single CSV file will be created with all cards from all emails within the CSV file
	    False - CSV file will be created for EACH email address with all cards from that email address
	
	h) **Note:** When logging in for the first time, gmail may block access. You will need to follow the steps in the email to follow to enable less secure applications.  (https://myaccount.google.com/lesssecureapps)
	
4) Windows Users: Double click on MasterExtractor.bat to run the program. You will get .jpg screenshots in the "screenshots" folder and will have a .csv file for each FROM email address.

    Alternatively you can run master-extractor.py for a python IDE.  PyCharm Recommended (https://www.jetbrains.com/pycharm/)
 
    In order to access the cards, open the .csv file with your notepad program of choice. Notepad++ Recommended (https://notepad-plus-plus.org/download/v7.5.3.html).

5) Sometimes chrome updates drivers automatically and you will need to download updated drivers here: https://chromedriver.storage.googleapis.com/index.html

    Place driver in the same folder as this README.md file

6) Useful Excel Formulas to turn CSV output into Cells.  Paste output into Cell A

        Cell A Example: 1234567891011,1234,50.00 
        Cell B: =LEFT(A1, SEARCH(",",A1,1)-1)        
        Cell C: =MID(A1,SEARCH(",",A1,1)+1,SEARCH(",",A1,SEARCH(",",A1,1)+1)-SEARCH(",",A1,1)-1)        
        Cell D: =RIGHT(A1,LEN(A1)-SEARCH(",",A1,SEARCH(",",A1,1)+1))

7) Enjoy!
