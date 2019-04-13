//©Patrick Hultquist 2019
//for help contact phultquist@imsa.edu
//to make this work best, change your userName to whatever your organizationName is. https://support.google.com/mail/answer/8158?hl=en


var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings")
var formSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]

var lastRow = formSheet.getLastRow()//formSheet, obviously. when a form is submitted, it goes into last row. 


//The following are row/column numbers for settings sheet
var foodCartNameRow = 2,
    numberOfIColumnsBeforeItemsRow = 3,
    imageSRCRow = 5,//there is a reason this is 4
    foodCartDateRow = 6,
    settingsSheetAttributeColumn = 2,
    settingsSheetDisplayNameColumn = 3,
    settingsSheetHeaderColumn = 1; //meaning the titles for each attribute


var lastRowBeforeItems = foodCartDateRow; //On settings Sheet. I know this is confusing. On settings sheet, the last row of 'settings' listed is, right now, row 5. After that, item prices are different. I just need to distinguish between the 2.

//The following are row/column numbers referring to formSheet
var emailAddressColumn = 3,
    nameColumn = 2,
    numberOfColumnsBeforeItems = parseInt(settingsSheet.getRange(numberOfIColumnsBeforeItemsRow, settingsSheetAttributeColumn).getValue()), //reffering to formSheet. The number of columns (questions) before the item quantities are specified in the 
    titleRow = 1 //i.e. the row with the headers like "Email Address"

var numberOfItems = formSheet.getLastColumn()-numberOfColumnsBeforeItems


function sendReminder(){
  var foodCartDateValue = settingsSheet.getRange(foodCartDateRow, settingsSheetAttributeColumn).getValue()  
  var foodCartDate = new Date(foodCartDateValue)
  foodCartDate.setHours(0,0,0,0)
  var currentDate = new Date()
  currentDate.setHours(0,0,0,0)
  if ((currentDate.getYear() == foodCartDate.getYear())&&(currentDate.getDay() == foodCartDate.getDay())&&(currentDate.getMonth()==foodCartDate.getMonth())){
    for (p=titleRow+1;p<lastRow+1;p++){ //this var name cannot be i.
      //sendEmail(true, p)
    }
  } 
}

//needs to go through all the email

function sendWhenSubmitted(){
  sendEmail(false, lastRow)
}

function sendEmail(sendReminder, rowToProcess) {

  var organizationName = "SCC"

  //var file = DriveApp.getFilesByName(name)
  var imsaLogoURL = "https://drive.google.com/uc?id=1rkkzUbGZyYNBmRnA52DJPa-JDdirIDLO"
  //var imsaLogoURL = "https://drive.google.com/uc?id=1_kkHO2btamsKljYmvvavSidBo2inrbLz"

  var foodCartURL = settingsSheet.getRange(imageSRCRow, settingsSheetAttributeColumn).getValue() //
  var thankYouURL = "https://www.filepicker.io/api/file/2KMVSEJSOaxy1uHWWt1A"
  
  //Gets Image from google drive, formats it
  //Creates a blob of each logo to be referenced in HTML
  var imsaLogoBlob = UrlFetchApp
                       .fetch(imsaLogoURL)
                       .getBlob()
                       .setName("imsaLogoBlob");
  var foodCartBlob = UrlFetchApp
                       .fetch(foodCartURL)
                       .getBlob()
                       .setName("foodCartBlob");
  var thankYouBlob = UrlFetchApp
                       .fetch(thankYouURL)
                       .getBlob()
                       .setName("thankYouBlob");
  
    
  var imsaKey = "imsaLogo",
      foodCartKey = "foodCartLogo",
      thankYouKey = "thankYouLogo";
  
  var inline = {
    imsaKey: imsaLogoBlob,
    foodCartKey: foodCartBlob,
    thankYouKey: thankYouBlob
      }
  
  var imsaSRC = imgSRC('imsaKey'),
      foodCartSRC = imgSRC('foodCartKey'),
      thankYouSRC = imgSRC('thankYouKey')
  
  var itemNames = getNames(true, rowToProcess, true)
  var itemQuantities = getQuantities(rowToProcess)
  var itemPrices = getPrices()
  var numberOfItems = formSheet.getLastColumn()-numberOfColumnsBeforeItems
  var itemsTotal = getTotal(rowToProcess)

  itemNames = itemNames.join('<br>')

  
  var foodCartName = settingsSheet.getRange(foodCartNameRow, settingsSheetAttributeColumn).getValue()  
  var foodCartDateValue = settingsSheet.getRange(foodCartDateRow, settingsSheetAttributeColumn).getValue()  
  var foodCartDate = Utilities.formatDate(new Date(foodCartDateValue), "z", "MMMM d")
  
  var email = formSheet.getRange(rowToProcess, emailAddressColumn).getValues(),
      name = formSheet.getRange(rowToProcess, nameColumn).getValue(),
      subject = organizationName+" Food Cart Purchase",
      purchasedByText = name + '<br>' + email;
  
  
  
  if (sendReminder == true){
    subject = organizationName+" Food Cart Tonight @10 Check"
  }
  
  var html = '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml"><head> <meta http-equiv="Content-Type" content="text/html; charset=utf-8" /> <meta name="viewport" content="width=320, initial-scale=1" /> <title>Thank You</title> <style type="text/css" media="screen"> /* ----- Client Fixes ----- */ /* Force Outlook to provide a "view in browser" message */ #outlook a { padding: 0; } /* Force Hotmail to display emails at full width */ .ReadMsgBody { width: 100%; } .ExternalClass { width: 100%; } /* Force Hotmail to display normal line spacing */ .ExternalClass, .ExternalClass p, .ExternalClass span, .ExternalClass font, .ExternalClass td, .ExternalClass div { line-height: 100%; } /* Prevent WebKit and Windows mobile changing default text sizes */ body, table, td, p, a, li, blockquote { -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; } /* Remove spacing between tables in Outlook 2007 and up */ table, td { mso-table-lspace: 0pt; mso-table-rspace: 0pt; } /* Allow smoother rendering of resized image in Internet Explorer */ img { -ms-interpolation-mode: bicubic; } /* ----- Reset ----- */ html, body, .body-wrap, .body-wrap-cell { margin: 0; padding: 0; background: #ffffff; font-family: Arial, Helvetica, sans-serif; font-size: 16px; color: #89898D; text-align: left; } img { border: 0; line-height: 100%; outline: none; text-decoration: none; } table { border-collapse: collapse !important; } td, th { text-align: left; font-family: Arial, Helvetica, sans-serif; font-size: 16px; color: #89898D; line-height: 1.5em; } /* ----- General ----- */ h1, h2 { line-height: 1.1; text-align: right; } h1 { margin-top: 0; margin-bottom: 10px; font-size: 24px; } h2 { margin-top: 0; margin-bottom: 60px; font-weight: normal; font-size: 17px; } .outer-padding { padding: 50px 0; } .col-1 { border-right: 1px solid #D9DADA; width: 180px; } td.hide-for-desktop-text { font-size: 0; height: 0; display: none; color: #ffffff; } img.hide-for-desktop-image { font-size: 0 !important; line-height: 0 !important; width: 0 !important; height: 0 !important; display: none !important; } .body-cell { background-color: #ffffff; padding-top: 60px; vertical-align: top; } .body-cell-left-pad { padding-left: 30px; padding-right: 14px; } /* ----- Modules ----- */ .brand td { padding-top: 25px; } .brand a { font-size: 16px; line-height: 59px; font-weight: bold; } .data-table th, .data-table td { width: 350px; padding-top: 5px; padding-bottom: 5px; padding-left: 5px; } .data-table th { background-color: #f9f9f9; color: #1A3768; } .data-table td { padding-bottom: 30px; } .data-table .data-table-amount { font-weight: bold; font-size: 20px; } </style> <style type="text/css" media="only screen and (max-width: 650px)"> @media only screen and (max-width: 650px) { table[class*="w320"] { width: 320px !important; } td[class*="col-1"] { border: none; } td[class*="hide-for-mobile"] { font-size: 0 !important; line-height: 0 !important; width: 0 !important; height: 0 !important; display: none !important; } img[class*="hide-for-desktop-image"] { width: 176px !important; height: 135px !important; display: block !important; padding-left: 60px; } td[class*="hide-for-desktop-image"] { width: 100% !important; display: block !important; text-align: right !important; } td[class*="hide-for-desktop-text"] { display: block !important; text-align: center !important; font-size: 16px !important; height: 61px !important; padding-top: 30px !important; padding-bottom: 20px !important; color: #89898D !important; } td[class*="mobile-padding"] { padding-top: 15px; } td[class*="outer-padding"] { padding: 0 !important; } td[class*="body-cell-left-pad"] { padding-left: 20px; padding-right: 20px; } } </style></head><body class="body" style="padding:0; margin:0; display:block; background:#ffffff; -webkit-text-size-adjust:none" bgcolor="#ffffff"> <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#ffffff"> <tr> <td class="outer-padding" valign="top" align="left"> <center> <table class="w320" cellspacing="0" cellpadding="0" width="600" height="723"> <tr> <td class="col-1 hide-for-mobile"> <table cellspacing="0" cellpadding="0" width="100%"> <tr> <td class="hide-for-mobile" style="padding:30px 0 10px 0;"> <img width="156" height="41" src="'+imsaSRC+'" alt="logo" /> </td> </tr> <tr> <td class="hide-for-mobile" height="150" valign="top"> <b> <span>'+organizationName+'</span> </b> <br> <span></span> </td> </tr> <tr> <td class="hide-for-mobile" style="height:180px; width:180px;"> <img width="180" height="180" src="'+foodCartSRC+'" alt="large logo" /> </td> </tr> </table> </td> <td valign="top" class="col-2"> <table cellspacing="0" cellpadding="0" width="100%"> <tr> <td class="body-cell body-cell-left-pad" width="355" height="661" valign="top"> <table cellpadding="0" cellspacing="0"> <tr> <td width="355"> <h1> <span> Your Food Cart Purchase </span> </h1> <h2> <span></span> </h2> </td> </tr> <tr> <td width="355"> <table class="data-table" cellpadding="0" cellspacing="0"> <tr> <th> Purchased By </th> </tr> <tr> <td> <span> '+purchasedByText+' </span> </td> </tr> <tr> <th> Food Cart </th> </tr> <tr> <td> <span> '+foodCartName+' </span> </td> </tr> <tr> <th> Food Cart Date </th> </tr> <tr> <td> <span> '+foodCartDate+' </span> </td> </tr> <tr> <th> Items </th> </tr> <tr> <td> <span> '+itemNames+' </span> </td> </tr> <tr> <th> Amount Due </th> </tr> <tr> <td> <table cellpadding="0" cellspacing="0"> <tr> <td class="data-table-amount" style="width: 85px;"> <span> '+"$"+itemsTotal+' </span> </td> </tr> </table> </td> </tr> </table> </td> </tr> <tr> <td width="355" class="footer"> <center> <img width="213" height="48" src="'+thankYouSRC+'"> </center> </td> </tr> </table> <table cellspacing="0" cellpadding="0" width="100%"> <tr> <td class="hide-for-desktop-text"> <b> <span>Thanks for buying from SCC</span> </b> <br> <span><br></span> </td> </tr> </table> </td> </tr> </table> </td> </tr> </table> </center> </td> </tr> </table></body></html>'

  //sheet.insertImage(image, 1, 1)
  
  if (itemsTotal == 0&&sendReminder == false){
    //if someone checks out with a total of $0, sends this email (easter egg)
    MailApp.sendEmail(email, "Your 'Purchase'", "Hey, good try. You can't purchase $0. If you actually did buy something and you were sent this email contact "+organizationName)
    return;
  }
  
  //MailApp.sendEmail(email,subject,"",{htmlBody: html, inlineImages: inline, attachments: null})
}

function imgSRC(imageKey){
  return 'cid:'+imageKey 
}

function getNames(wantsItemsBought, rowToProcess, wantsDisplayName){
  //Note: if wantsItemsBought, will not include options (i.e. do you want cheese?)
  //If watntsItemsBought must provide rowToProcess. Also will attach the quantity of each thing they bought
  var itemNames = [],
      itemPrices = getPrices(),
      itemQuantities = getQuantities(rowToProcess)
      
      for (i = 1; i < numberOfItems+1;i++){
        var itemName = formSheet.getRange(titleRow, i+numberOfColumnsBeforeItems).getValue()
        //var itemName = settingsSheet.getRange(i+lastRowBeforeItems, settingsSheetHeaderColumn).getValue()
        if (wantsDisplayName==true){
          itemName = settingsSheet.getRange(i+lastRowBeforeItems, settingsSheetDisplayNameColumn).getValue()
        }
        if (wantsItemsBought==true){
          if (itemPrices[i-1]==0 || itemQuantities[i-1]==0){ //Let's say someone bought ketcup? Do you want it to say that on the receipt even if it cost $0? if so, delete itemPrices[i-1]==0 ||
          }else if (isNaN(itemQuantities[i-1])!=true){
            itemName = itemName+" ("+itemQuantities[i-1]+")"
            itemNames.push(itemName)
          }
        } else if (wantsItemsBought==false){
          itemNames.push(itemName)
        }
      }
  return itemNames
}

function getPrices(){
  var itemPrices = []

  for (i = 1; i < numberOfItems+1;i++){
    var itemPrice = parseFloat(settingsSheet.getRange(lastRowBeforeItems+1+(i-1), settingsSheetAttributeColumn).getValue())
    if (isNaN(itemPrice)==true){itemPrice = 0}
    itemPrices.push(itemPrice)
  }
  Logger.log("Item Prices: "+itemPrices)
    return itemPrices
}

function getTotal(rowToProcess){
  var itemsTotal = 0
  var itemQuantities = getQuantities(rowToProcess),
      itemPrices = getPrices()
    for (i =1; i < numberOfItems+1;i++){
    //For however many items there are, adds the price of each individual
    if (itemQuantities[i-1]<1){itemQuantities[i-1]=0}
    itemsTotal = parseFloat(itemsTotal) + parseFloat(itemQuantities[i-1])*parseFloat(itemPrices[i-1])
  }
  Logger.log("Total: "+ itemsTotal)
  return itemsTotal
}

function getQuantities(rowToProcess){
  var itemQuantities = []
  for (i = 1; i < numberOfItems+1;i++){
    //number of items-1 because i starts at 0
    var itemQuantity = getQuantityOfCell(rowToProcess, i+numberOfColumnsBeforeItems, formSheet)
    itemQuantities.push(itemQuantity)
  }
  Logger.log(itemQuantities)
  return itemQuantities
}

function updateSettingsSheet(){
  var names = getNames(false,1,false)//does not matter what the rowToProcess is
  Logger.log(names)
   for (i in names){
     settingsSheet.getRange(parseInt(lastRowBeforeItems)+parseInt(i)+1, settingsSheetHeaderColumn).setValue(names[i])
   }
}

function getQuantityOfCell(row, column, sheet) {
  //is built for yes/no adoption
  var value = sheet.getRange(row, column).getValue()
  if (isNaN(value)==true || value.length == 0){
      if (value.toUpperCase() == "YES" || value.toUpperCase() == "YES! CHEESEBURGERS!") {
        value = 1
      }else{
        value = 0
      }
    }
  
  //makes quantity forced to be integer
  return parseInt(value)
}


function getTotalQuantities(){
  //gets and returns the sum of each individual part
  var totalQuantities = []
  for (j=1;j<=numberOfItems;j++){
    var itemTotalQuantity = 0
    for (o=1;o<=lastRow-1;o++){
      var row = o+titleRow
      var column = parseFloat(j)+parseFloat(numberOfColumnsBeforeItems)
      var cellQuantity = getQuantityOfCell(row,column,formSheet)
      itemTotalQuantity = itemTotalQuantity + cellQuantity 
    }
    totalQuantities.push(itemTotalQuantity)
  }
  return totalQuantities;
}

