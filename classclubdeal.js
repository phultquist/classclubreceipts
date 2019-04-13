//©Patrick Hultquist 2019
//for help contact phultquist@imsa.edu

var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings")
var formSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]

var lastRow = formSheet.getLastRow() //formSheet, obviously. when a form is submitted, it goes into last row. 


//The following are row/column numbers for settings sheet
var foodCartNameRow = 2,
    numberOfIColumnsBeforeItemsRow = 3,
    imageSRCRow = 5, //there is a reason this is 4
    foodCartDateRow = 6,
    settingsSheetAttributeColumn = 2,
    settingsSheetDisplayNameColumn = 3,
    settingsSheetHeaderColumn = 1; //meaning the titles for each attribute


var lastRowBeforeItems = foodCartDateRow; //On settings Sheet. I know this is confusing. On settings sheet, the last row of 'settings' listed is, right now, row 5. After that, item prices are different. I just need to distinguish between the 2.

//The following are row/column numbers referring to formSheet
var emailAddressColumn = 2,
    nameColumn = 3,
    numberOfColumnsBeforeItems = parseInt(settingsSheet.getRange(numberOfIColumnsBeforeItemsRow, settingsSheetAttributeColumn).getValue()), //reffering to formSheet. The number of columns (questions) before the item quantities are specified in the 
    titleRow = 1 //i.e. the row with the headers like "Email Address"

var numberOfItems = formSheet.getLastColumn() - numberOfColumnsBeforeItems


function sendReminder() {
    var foodCartDateValue = settingsSheet.getRange(foodCartDateRow, settingsSheetAttributeColumn).getValue()
    var foodCartDate = new Date(foodCartDateValue)
    foodCartDate.setHours(0, 0, 0, 0)
    var currentDate = new Date()
    currentDate.setHours(0, 0, 0, 0)
    /*if ((currentDate.getYear() == foodCartDate.getYear()) && (currentDate.getDay() == foodCartDate.getDay()) && (currentDate.getMonth() == foodCartDate.getMonth())) {
        for (p = titleRow + 1; p < lastRow + 1; p++) { //this var name cannot be i.
            sendEmail(true, p)
            SpreadsheetApp.getActiveSpreadsheet().toast("Reminder was sent", "Success!")
        }
    }else{
      SpreadsheetApp.getActiveSpreadsheet().toast("Reminder can only be sent during the food cart day", "Sorry!")
    }*/
    for (p = titleRow + 1; p < lastRow + 1; p++) { //this var name cannot be i.
      sendEmail(true, p)
      SpreadsheetApp.getActiveSpreadsheet().toast("Reminder was sent", "Success!")
    }
}

//needs to go through all the email

function sendWhenSubmitted() {
    sendEmail(false, lastRow)
}

function sendEmail(sendReminder, rowToProcess) {

    var organizationName = "SoCC"

    //var file = DriveApp.getFilesByName(name)
    //var imsaLogoURL = "https://drive.google.com/uc?id=1_kkHO2btamsKljYmvvavSidBo2inrbLz"

    var headerURL = settingsSheet.getRange(imageSRCRow, settingsSheetAttributeColumn).getValue() //
    var bannerURL = "https://drive.google.com/uc?id=1Ckr9okz-EJBYofE0_sQGST8F2k-lwaEU"

    //Gets Image from google drive, formats it
    //Creates a blob of each logo to be referenced in HTML

    var headerBlob = UrlFetchApp
        .fetch(headerURL)
        .getBlob()
        .setName("headerBlob");
    var bannerBlob = UrlFetchApp
        .fetch(bannerURL)
        .getBlob()
        .setName("bannerBlob");


    var headerKey = "headerLogo",
        bannerKey = "bannerLogo";

    var inline = {
        headerKey: headerBlob,
        bannerKey: bannerBlob
    }

    var headerSRC = imgSRC('headerKey'),
        bannerSRC = imgSRC('bannerKey')

    var itemNames = getNames(true, rowToProcess, true)
    var itemQuantities = getQuantities(rowToProcess)
    var itemPrices = getPrices(rowToProcess),
        itemsBoughtPrices = getItemsBoughtPrices(rowToProcess, true)
    var numberOfItems = formSheet.getLastColumn() - numberOfColumnsBeforeItems
    var itemsTotal = getTotal(rowToProcess)
    var tacoCount = countTacos(rowToProcess),
        tacoDealCount = Math.floor(tacoCount / 3)

    if (tacoCount > 2) {
        itemNames.unshift("Taco Deal (3 for $5) "+tacoDealCount+"x");

        itemsBoughtPrices.unshift("$" + tacoDealPrice(rowToProcess))
        itemsTotal = parseFloat(itemsTotal) + tacoDealPrice(rowToProcess)
    }

    itemsBoughtPrices = itemsBoughtPrices.join('<br>')

    itemNames = itemNames.join('<br>')


    var foodCartName = settingsSheet.getRange(foodCartNameRow, settingsSheetAttributeColumn).getValue()
    var foodCartDateValue = settingsSheet.getRange(foodCartDateRow, settingsSheetAttributeColumn).getValue()
    var foodCartDate = Utilities.formatDate(new Date(foodCartDateValue), "z", "MMMM d")

    var email = formSheet.getRange(rowToProcess, emailAddressColumn).getValues(),
        name = formSheet.getRange(rowToProcess, nameColumn).getValue(),
        subject = organizationName + " Food Cart Purchase"

    name = toTitleCase(name)

    var purchasedByText = name;




    if (sendReminder == true) {
        subject = organizationName + " Food Cart Tonight @10 Check"
    }

    var html = '<!DOCTYPE html><html xmlns="http://www.w3.org/1999/xhtml"><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8" /><meta name="viewport" content="width=320, initial-scale=1" /><title>Thank You</title><style type="text/css" media="screen">/* ----- Client Fixes ----- *//* Force Outlook to provide a "view in browser" message */#outlook a {padding: 0;}/* Force Hotmail to display emails at full width */.ReadMsgBody {width: 100%;}.ExternalClass {width: 100%;}/* Force Hotmail to display normal line spacing */.ExternalClass,.ExternalClass p,.ExternalClass span,.ExternalClass font,.ExternalClass td,.ExternalClass div {line-height: 100%;}/* Prevent WebKit and Windows mobile changing default text sizes */body, table, td, p, a, li, blockquote {-webkit-text-size-adjust: 100%;-ms-text-size-adjust: 100%;}/* Remove spacing between tables in Outlook 2007 and up */table, td {mso-table-lspace: 0pt;mso-table-rspace: 0pt;}/* Allow smoother rendering of resized image in Internet Explorer */img {-ms-interpolation-mode: bicubic;}/* ----- Reset ----- */html, body, .body-wrap, .body-wrap-cell {margin: 0;padding: 0;background: #ffffff;font-family: "Roboto", Arial, Helvetica, sans-serif;font-size: 16px;color: black;text-align: left;}img {border: 0;line-height: 100%;outline: none;text-decoration: none;}table {border-collapse: collapse !important;}td, th {text-align: left;font-family: "Roboto", Arial, Helvetica, sans-serif;font-size: 16px;color: black;line-height: 1.5em;}/* ----- General ----- */h1, h2 {line-height: 1.1;text-align: right;}h1 {margin-top: 0;margin-bottom: 10px;font-size: 24px;}h2 {margin-top: 0;margin-bottom: 60px;font-weight: normal;font-size: 17px;}.outer-padding {padding: 50px 0;}.col-1 {border-right: 1px solid #D9DADA;width: 180px;}td.hide-for-desktop-text {font-size: 0;height: 0;display: none;color: #ffffff;}img.hide-for-desktop-image {font-size: 0 !important;line-height: 0 !important;width: 0 !important;height: 0 !important;display: none !important;}.body-cell {background-color: #ffffff;padding-top: 60px;vertical-align: top;}.body-cell-left-pad {padding-left: 30px;padding-right: 14px;}/* ----- Modules ----- */.brand td {padding-top: 25px;}.brand a {font-size: 16px;line-height: 59px;font-weight: bold;}.data-table th,.data-table td{width: 350px;height: 45px;line-height: 45px !important;padding-top: 2px !important;padding-bottom: 2px !important;padding-left: 10px !important;padding-right: 10px !important;background: #f3f3f3;border-top: 2px white solid;border-bottom: 2px white solid;}.data-table th {background-color: #f9f9f9;color: black;}.data-table td {padding-bottom: 30px;}.data-table .data-table-amount {font-weight: bold;font-size: 20px;}/*JGAO*/#banner{background-image: url("' + headerSRC + '");background-size: cover;color: white;height: 150px;padding-top: 50px;}#total{border-top: black 2px solid;}#thanks {padding: 5px;display: table;vertical-align: middle;;}.footer{}/** {border: red 2px solid !important;}*/</style><style type="text/css" media="only screen and (max-width: 650px)">@media only screen and (max-width: 650px) {table[class*="w320"] {width: 320px !important;}td[class*="col-1"] {border: none;}td[class*="hide-for-mobile"] {font-size: 0 !important;line-height: 0 !important;width: 0 !important;height: 0 !important;display: none !important;}img[class*="hide-for-desktop-image"] {width: 176px !important;height: 135px !important;display: block !important;padding-left: 60px;}td[class*="hide-for-desktop-image"] {width: 100% !important;display: block !important;text-align: right !important;}td[class*="hide-for-desktop-text"] {display: block !important;text-align: center !important;font-size: 16px !important;height: 61px !important;padding-top: 30px !important;padding-bottom: 20px !important;color: black !important;}td[class*="mobile-padding"] {padding-top: 15px;}td[class*="outer-padding"] {padding: 0 !important;}td[class*="body-cell-left-pad"] {padding-left: 20px;padding-right: 20px;}}</style></head><body class="body" style="padding:0; margin:0; display:block; background:#ffffff; -webkit-text-size-adjust:none" bgcolor="#ffffff"><table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#ffffff"><tr><td class="outer-padding" valign="top" align="left"><center><table class="w320" cellspacing="0" cellpadding="0" width="600" height="723"><!--<tr>--><td valign="top" class="col-2"><table cellspacing="0" cellpadding="0" width="100%"><tr><td class="body-cell body-cell-left-pad" width="600" height="661" valign="top"><table cellpadding="0" cellspacing="0"><tr><td width="600" height="50"><img src="' + bannerSRC + '" width="600" height="50"></td></tr><tr><td width="600" id="banner"><h1 style="text-align: center; font-size: 64px;"> <span> Enjoy Your Meal </span> </h1><h2 style="text-align: center;"> <span>Your Food Cart order is in</span> </h2></td></tr><tr><td width="600"><table class="data-table" cellpadding="0" cellspacing="0"><!--Order details--><tr><td><span> ' + foodCartName + ' </span></td><td style="text-align: right;"><span> ' + foodCartDate + ' </span></td></tr><tr><td><span> ' + itemNames + ' </span></td><td style="text-align: right;"><span> ' + itemsBoughtPrices + ' </span></td></tr><tr id="total"><td><span> Total </span></td><td style="text-align: right;"><span> ' + "$" + itemsTotal + ' </span></td></tr></table></td></tr><tr class="footer" width="600"><td style="text-align: center" height="150">For ' + purchasedByText + ' <br>Thank you for supporting ' + organizationName + '</td></tr><tr><td><center><img width="600" height="50" src="' + bannerSRC + '"><p style="font-size: 12px; color: gray">©Patrick Hultquist 2019</p><p style="font-size: 12px; color: gray">©Jon Gao 2019</p></center></td></tr></table></td></tr></table></td><!--</tr>--></table></center></td></tr></table></body></html>'

    //sheet.insertImage(image, 1, 1)
    Logger.log("processing row " + rowToProcess)

    if (itemsTotal == 0 && sendReminder == false) {
        //if someone checks out with a total of $0, sends this email (easter egg)
        MailApp.sendEmail(email, "Your 'Purchase'", "Hey, good try. You can't purchase $0. If you actually did buy something and you were sent this email contact "+organizationName)
        return;
    }
    MailApp.sendEmail("phultquist@imsa.edu", subject, "", {
        htmlBody: html,
        inlineImages: inline,
        attachments: null
    })

}

function testEmail(){
 sendEmail(false,41) 
}

function imgSRC(imageKey) {
    return 'cid:' + imageKey
}

function getNames(wantsItemsBought, rowToProcess, wantsDisplayName) {
    //Note: if wantsItemsBought, will not include options (i.e. do you want cheese?)
    //If watntsItemsBought must provide rowToProcess. Also will attach the quantity of each thing they bought
    var itemNames = [],
        itemPrices = getPrices(rowToProcess),
        itemQuantities = getQuantities(rowToProcess),
        dealItems = getDealItems(),
        tacoCount = countTacos(rowToProcess)

    for (i = 1; i < numberOfItems + 1; i++) {
        var itemName = formSheet.getRange(titleRow, i + numberOfColumnsBeforeItems).getValue()
        //var itemName = settingsSheet.getRange(i+lastRowBeforeItems, settingsSheetHeaderColumn).getValue()
        if (wantsDisplayName == true) {
            itemName = settingsSheet.getRange(i + lastRowBeforeItems, settingsSheetDisplayNameColumn).getValue()
        }
        if (wantsItemsBought == true) {
            if (dealItems[i - 1] == 1 && tacoCount > 2) {
                itemName = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + itemName
            }
            if (itemPrices[i - 1] == 0 || itemQuantities[i - 1] == 0) { //Let's say someone bought ketchup? Do you want it to say that on the receipt even if it cost $0? if so, delete itemPrices[i-1]==0 ||
            } else if (isNaN(itemQuantities[i - 1]) != true) {
                itemName = itemName + " " + itemQuantities[i - 1] + "x"
                itemNames.push(itemName)
            }
        } else if (wantsItemsBought == false) {
            itemNames.push(itemName)
        }
    }
    return itemNames
}

function getPrices(rowToProcess) {
    var itemPrices = [],
        dealItems = getDealItems()
    for (i = 1; i < numberOfItems + 1; i++) {
        //number of items-1 because i starts at 0
        var itemPrice = formSheet.getRange(rowToProcess, i + numberOfColumnsBeforeItems).getValue()
        try {
              var dollarIndex = parseFloat(itemPrice.indexOf("$"));
          itemPrice=itemPrice.substring(dollarIndex+1,itemPrice.length-1)
        } catch (e) {

        }

        if (itemPrice.length == 0) {
            itemPrice = 0
        }
        itemPrices.push(itemPrice)
    }
    Logger.log("Item Prices: " + itemPrices)
    return itemPrices
}

function getTotal(rowToProcess) {
    var itemsTotal = 0
    var itemQuantities = getQuantities(rowToProcess),
        itemPrices = getPrices(rowToProcess),
        dealItems = getDealItems(),
        tacoCount = countTacos(rowToProcess)
    for (i = 1; i < numberOfItems + 1; i++) {
        //For however many items there are, adds the price of each individual
        var itemTotal = parseFloat(itemPrices[i - 1])
        Logger.log("item total " + itemTotal)
        if (itemQuantities[i - 1] < 1) {
            itemQuantities[i - 1] = 0
        }
        if (dealItems[i - 1] == 1 && tacoCount > 2) {
            itemTotal = 0
        }
        itemsTotal = parseFloat(itemsTotal) + itemTotal
    }
    Logger.log("Total: " + itemsTotal)
    return itemsTotal
}

function testQ() {
    for (j = 2; j < formSheet.getLastRow() + 1; j++) {
        sendEmail(false, j)
    }
}

function getQuantities(rowToProcess) {
    var itemQuantities = []
    for (i = 1; i < numberOfItems + 1; i++) {
        //number of items-1 because i starts at 0
        var itemQuantity = getQuantityOfCell(rowToProcess, i + numberOfColumnsBeforeItems, formSheet)
        itemQuantities.push(itemQuantity)
    }
    Logger.log("item quantities: " + itemQuantities)
    return itemQuantities
}

function getItemsBoughtPrices(rowToProcess, withDollarSign) {
    var itemQuantities = getQuantities(rowToProcess),
        itemPrices = getPrices(rowToProcess),
        itemsBoughtPrices = [],
        dealItems = getDealItems(),
        tacoCount = countTacos(rowToProcess)

    Logger.log(dealItems + " is deal items")
    for (i in itemQuantities) {
        var itemBoughtPrice = itemPrices[i]

        if (itemQuantities[i] == 0) {
            continue;
        }
        if (dealItems[i] == 1 && tacoCount > 2) {
            itemBoughtPrice = "&nbsp;";
        } else {
            if (withDollarSign == true) {
                itemBoughtPrice = "$" + itemBoughtPrice
            }
        }
        itemsBoughtPrices.push(itemBoughtPrice)



    }
    return itemsBoughtPrices;

}

function updateSettingsSheet() {
    var names = getNames(false, 1, false) //does not matter what the rowToProcess is
    Logger.log(names)
    for (i in names) {
        settingsSheet.getRange(parseInt(lastRowBeforeItems) + parseInt(i) + 1, settingsSheetHeaderColumn).setValue(names[i])
    }
    //lastRowBeforeItems   
}

function getQuantityOfCell(row, column, sheet) {
    //is b lt for yes/no adoption
    var value = sheet.getRange(row, column).getValue()
    try {
        value = getFirstCharacter(value)
    } catch (e) {

    }
    if (isNaN(value) == true || value.length == 0) {
        if (value.toUpperCase() == "YES" || value.toUpperCase() == "YES") {
            value = 1
        } else {
            value = 0
        }
    }

    //makes quantity forced to be integer
    return parseInt(value)
}


function getTotalQuantities() {
    //gets and returns the sum of each individual part
    var totalQuantities = []
    for (j = 1; j <= numberOfItems; j++) {
        var itemTotalQuantity = 0
        for (o = 1; o <= lastRow - 1; o++) {
            var row = o + titleRow
            var column = parseFloat(j) + parseFloat(numberOfColumnsBeforeItems)
            var cellQuantity = getQuantityOfCell(row, column, formSheet)
            itemTotalQuantity = itemTotalQuantity + cellQuantity
        }
        totalQuantities.push(itemTotalQuantity)
    }
    return totalQuantities;
}

function toTitleCase(str) {
    return str.replace(
        /\w\S*/g,
        function(txt) {
            return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
        }
    );
}

function getFirstCharacter(string) {
    return string.charAt(0)
}

function getDealItems() {
    var start = parseInt(lastRowBeforeItems) + parseInt(1)
    var dealValues = settingsSheet.getRange(start, settingsSheetAttributeColumn, settingsSheet.getLastRow() - start + 1, 1).getValues(),
        deals = []
    dealValues = convert2DArray(dealValues)
    for (d in dealValues) {
        if (dealValues[d] == true) {
            deals.push(1)
        } else {
            if (dealValues[d] == false) {
                deals.push(0)
            }
        }
    }

    return deals
}

function convert2DArray(array) {
    var arrToConvert = array;
    var newArr = [];
    for (var i = 0; i < arrToConvert.length; i++) {

        newArr = newArr.concat(arrToConvert[i]);
    }
    return newArr;
}


function countTacos(rowToProcess) {
    var quantities = getQuantities(rowToProcess),
        dealItems = getDealItems(),
        tacoArray = [],
        tacoCount = 0

    for (q in quantities) {
        tacoArray.push(quantities[q] * dealItems[q])
    }
    for (t in tacoArray) {
        tacoCount += tacoArray[t]
    }
    Logger.log("tacoCount " + tacoCount)
    return tacoCount
}

function tacoDealPrice(rowToProcess) {
    var tacoCount = parseInt(countTacos(rowToProcess))
    var tacoPrice = tacoCount * 2 - Math.floor(tacoCount / 3)
    return tacoPrice
}

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    // Or DocumentApp or FormApp.
    ui.createMenu('Send Emails')
        .addItem('Send To Everyone Who Has Submitted', 'testQ')
        .addSeparator()
        .addItem('Send Reminder [To everyone]', 'sendReminder')
        .addToUi();
}