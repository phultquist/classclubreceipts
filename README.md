# Class Club Receipts
These scripts have each been made to make the Class Clubs more efficient and provide better experiences for those who interact with the Class Clubs. For any questions about a specific script or a request for a new one please email phultquist@imsa.edu. Any edits to the code can be done through Tools > Script Editor. These use Google Apps Script (nearly identical to JavaScript) and HTML (for emails). Each of these should be tested before put into general use. Patrick Hultquist is not responsible for any errors in each script or documentation or if Admin gets mad.

## General & Form Setup
Originally made for the Class of 2019 Portillo’s Food Cart, this script automatically returns a receipt to anyone who pre-orders using a Google Form. It also provides basic calculations for the amount of items bought. Each Google Form must be linked to the spreadsheet in the folder. If a new spreadsheet is made, it is crucial that the code found in the Script Editor remains. The columns for each of the values (date, name, email, items bought, etc.) should stay the same. If there are changes to it, open the Script Editor and at the top change each of the column numbers. There should be no columns after the items that can be bought. The active form used for the analytics and receipt must be the first page in the Spreadsheet. To change the organization name from SCC to JCC or SoCC, open the script editor and change the value for the variable named “organizationName.”

## Settings Page
The settings page (which must be named “Settings”) contains several crucial attributes in order to make sure the script runs smoothly. The “Food Cart Name” is used in the email. The number of items in the form is based off the “Number Of Columns Before Item Quantities” in the settings sheet. Without this being accurate the script will disclude an item or recognize an attribute such as “ID Number” as an item. The image URL is the URL for the image used in the email. The food cart date must be a date type (when clicked a calendar shows up). Finally, all items are listed. To update this list in the menu bar click “Update Settings Sheet” then “Update.” The price and display name should be entered manually.

## Analytics Page
The analytics page totals the amount of each item bought. This runs automatically as long as the trigger is set up. Note if any values are added in manually they will be overridden. 

## Permissions & Triggers
Setting up triggers is another important event for this script. Triggers must be set up to allow each script to run automatically. The most important trigger is to run the function sendWhenSubmitted() onFormSubmit. To set up triggers go to the G Suite Developer Hub and click on the project name. Then, under project details, go to the Triggers page. updateAnalyticsSheet() runs the analytics sheet and updates the totals for each item. This trigger should be run every minute. Emails will be sent as whoever set up the trigger. Finally, sendReminder() is the last function that a Trigger must run. This will send an email to anyone who purchased something whenever it is run on the day of the food cart described in the Setting page. Thus, it should only be run the once a day. To ensure it does not run accidentally, it is best to set this trigger to run on a specific day at a specific time. 
