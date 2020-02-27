//©Patrick Hultquist 2019
//for help contact phultquist@imsa.edu
//to make this work best, change your userName to whatever your organizationName is. https://support.google.com/mail/answer/8158?hl=en

//reminder working
//does not have to be only for food carts
//yes/no option working
//do you want it on the receipt even if price = 0?
//updateSettingsSheet to make it easier to manage
//add support for optionals by treating them as objects
//show in receipt as column? So the receipt is cleaner than the questions
//the yes! cheeseburgers! crisis
//new forms how to set up
//where this can go
//UI changes


//set up triggers
//Submitted
//sendReminder
//analytics
//updateSettingsSheet?
var organizationName = "JCC";

var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings")
var formSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0],
    peopleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("People");

var lastRow = formSheet.getLastRow()//formSheet, obviously. when a form is submitted, it goes into last row. 


//The following are row/column numbers for settings sheet
var foodCartNameRow = 2,
    numberOfIColumnsBeforeItemsRow = 3,
    foodCartDateRow = 4,
    settingsSheetAttributeColumn = 2,
    settingsSheetDisplayNameColumn = 3,
    settingsSheetHeaderColumn = 1; //meaning the titles for each attribute


var lastRowBeforeItems = foodCartDateRow; //On settings Sheet. I know this is confusing. On settings sheet, the last row of 'settings' listed is, right now, row 5. After that, item prices are different. I just need to distinguish between the 2.

//The following are row/column numbers referring to formSheet
var emailAddressColumn = 3,
    nameColumn = 2,
    hallColumn = 4,
    numberOfColumnsBeforeItems = parseInt(settingsSheet.getRange(numberOfIColumnsBeforeItemsRow, settingsSheetAttributeColumn).getValue()), //reffering to formSheet. The number of columns (questions) before the item quantities are specified in the 
    titleRow = 1,
    sentCol = 5;//i.e. the row with the headers like "Email Address"

var numberOfItems = formSheet.getLastColumn()-numberOfColumnsBeforeItems

function onOpen(){
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Food Cart')
      .addItem('Summary By Person', 'summaryByPerson')
      .addItem('Update Settings Sheet', 'updateSettingsSheet')
      .addToUi();
}

function sendReminder(){
  var foodCartDateValue = settingsSheet.getRange(foodCartDateRow, settingsSheetAttributeColumn).getValue()  
  var foodCartDate = new Date(foodCartDateValue)
  foodCartDate.setHours(0,0,0,0)
  var currentDate = new Date()
  currentDate.setHours(0,0,0,0)
  for (p=titleRow+1;p<lastRow+1;p++){ //this var name cannot be i.
    send(createEmail(true, p));
  }
}

function sendToEveryone(){
  for (q=titleRow+1;q<lastRow+1;q++){ //this var name cannot be i.
    //if (q==lastRow) continue;
    send(createEmail(false, q));
  }
}

function sendToAllPast(){
  send(createEmail(false, 6));
  for (q=66;q<lastRow+1;q++){ //this var name cannot be i.
    //send(createEmail(false, q));
  }
}
//needs to go through all the email

function sendWhenSubmitted(){
  send(createEmail(false, lastRow));
}

function sendToSpecific(){
  send(createEmail(false, 7));
}

function sendTestEmail(){
  var row = 4
  var orderObj = createEmail(false, row);
  orderObj.address="phultquist@imsa.edu";
  send(orderObj);
}

function summaryByPerson(){
  if (!peopleSheet){
    peopleSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("People", 2);
  }
  peopleSheet.getRange(1, 1, 1, 4).setValues([["Name", "Hall", "Text", "Total"]]);
  
  var fullOrder;
  var fullOrdersCells = [];
  for (u=titleRow+1;u<lastRow+1;u++){ //this var name cannot be i.
    fullOrder = getOrder(u);
    fullOrdersCells.push([fullOrder.name, "0"+fullOrder.hall.toString(), fullOrder.paperReceiptText, "$"+fullOrder.itemsTotal])
    //peopleSheet.getRange(u, 1, 1, 4).setValues([[fullOrder.name, "0"+fullOrder.hall.toString(), fullOrder.paperReceiptText, "$"+fullOrder.itemsTotal]])
  }
  peopleSheet.getRange(2, 1, fullOrdersCells.length, 4).setValues(fullOrdersCells);
}

function getOrder(rowToProcess){
  var itemNames = getNames(true, rowToProcess, true),
      allNames = getNames(false, rowToProcess, true)
  var itemQuantities = getQuantities(rowToProcess);
  var itemPrices = getPrices();
  //var tacoItems = [0,1,2]
  
  var boughtItemPrices = itemPrices.map(function(x, i){
    if (itemQuantities[i] > 0){
      return itemQuantities[i] * x; 
    } else {
      return null; 
    }
  })
  boughtItemPrices = boughtItemPrices.filter(function(c){
   return c > 0; 
  })
  var itemPricesText = "";
  boughtItemPrices.map(function(x){
    itemPricesText += "&nbsp;$"+x+"<br>";
  })
  var numberOfItems = formSheet.getLastColumn()-numberOfColumnsBeforeItems
  var itemsTotal = getTotal(rowToProcess)
  itemNames = itemNames.join('<br>');

  var paperReceiptText = "";
  itemQuantities.map(function(x, i){
    if (x>0) {
      paperReceiptText += allNames[i] + " x" + x + " • ";
    }
  })
  
  var foodCartName = settingsSheet.getRange(foodCartNameRow, settingsSheetAttributeColumn).getValue()  
  var foodCartDateValue = settingsSheet.getRange(foodCartDateRow, settingsSheetAttributeColumn).getValue()  
  var foodCartDate = Utilities.formatDate(new Date(foodCartDateValue), "z", "MMMM d")
  
  var emailAddress = formSheet.getRange(rowToProcess, emailAddressColumn).getValue(),
      name = formSheet.getRange(rowToProcess, nameColumn).getValue(),
      hall = formSheet.getRange(rowToProcess, hallColumn).getValue(),
      subject = organizationName+" Food Cart Purchase";
  
  var purchasedByText = name;
  
  var order = {
    name: name,
    emailAddress: emailAddress,
    subject: subject,
    hall: hall,
    foodCartName: foodCartName,
    foodCartDate: foodCartDate,
    itemNames: itemNames,
    itemPrices: itemPrices,
    itemPricesText: itemPricesText,
    itemsTotal: itemsTotal,
    purchasedByText: purchasedByText,
    paperReceiptText: paperReceiptText
  }
  
  return order;
}

function createEmail(sendReminder, rowToProcess) {
  
  var order = getOrder(rowToProcess);
  order.emailAddress = order.emailAddress.toLowerCase();
  if (sendReminder == true){
    order.subject = organizationName+" Food Cart Tonight @10 Check"
  }
  //var pSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ppl")
  //pSheet.getRange(pSheet.getLastRow()+1, 1).setValue(itemsTotal);
  //var html = '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml"><head> <meta http-equiv="Content-Type" content="text/html; charset=utf-8" /> <meta name="viewport" content="width=320, initial-scale=1" /> <title>Thank You</title> <style type="text/css" media="screen"> /* ----- Client Fixes ----- */ /* Force Outlook to provide a "view in browser" message */ #outlook a { padding: 0; } /* Force Hotmail to display emails at full width */ .ReadMsgBody { width: 100%; } .ExternalClass { width: 100%; } /* Force Hotmail to display normal line spacing */ .ExternalClass, .ExternalClass p, .ExternalClass span, .ExternalClass font, .ExternalClass td, .ExternalClass div { line-height: 100%; } /* Prevent WebKit and Windows mobile changing default text sizes */ body, table, td, p, a, li, blockquote { -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; } /* Remove spacing between tables in Outlook 2007 and up */ table, td { mso-table-lspace: 0pt; mso-table-rspace: 0pt; } /* Allow smoother rendering of resized image in Internet Explorer */ img { -ms-interpolation-mode: bicubic; } /* ----- Reset ----- */ html, body, .body-wrap, .body-wrap-cell { margin: 0; padding: 0; background: #ffffff; font-family: Arial, Helvetica, sans-serif; font-size: 16px; color: #89898D; text-align: left; } img { border: 0; line-height: 100%; outline: none; text-decoration: none; } table { border-collapse: collapse !important; } td, th { text-align: left; font-family: Arial, Helvetica, sans-serif; font-size: 16px; color: #89898D; line-height: 1.5em; } /* ----- General ----- */ h1, h2 { line-height: 1.1; text-align: right; } h1 { margin-top: 0; margin-bottom: 10px; font-size: 24px; } h2 { margin-top: 0; margin-bottom: 60px; font-weight: normal; font-size: 17px; } .outer-padding { padding: 50px 0; } .col-1 { border-right: 1px solid #D9DADA; width: 180px; } td.hide-for-desktop-text { font-size: 0; height: 0; display: none; color: #ffffff; } img.hide-for-desktop-image { font-size: 0 !important; line-height: 0 !important; width: 0 !important; height: 0 !important; display: none !important; } .body-cell { background-color: #ffffff; padding-top: 60px; vertical-align: top; } .body-cell-left-pad { padding-left: 30px; padding-right: 14px; } /* ----- Modules ----- */ .brand td { padding-top: 25px; } .brand a { font-size: 16px; line-height: 59px; font-weight: bold; } .data-table th, .data-table td { width: 350px; padding-top: 5px; padding-bottom: 5px; padding-left: 5px; } .data-table th { background-color: #f9f9f9; color: #1A3768; } .data-table td { padding-bottom: 30px; } .data-table .data-table-amount { font-weight: bold; font-size: 20px; } </style> <style type="text/css" media="only screen and (max-width: 650px)"> @media only screen and (max-width: 650px) { table[class*="w320"] { width: 320px !important; } td[class*="col-1"] { border: none; } td[class*="hide-for-mobile"] { font-size: 0 !important; line-height: 0 !important; width: 0 !important; height: 0 !important; display: none !important; } img[class*="hide-for-desktop-image"] { width: 176px !important; height: 135px !important; display: block !important; padding-left: 60px; } td[class*="hide-for-desktop-image"] { width: 100% !important; display: block !important; text-align: right !important; } td[class*="hide-for-desktop-text"] { display: block !important; text-align: center !important; font-size: 16px !important; height: 61px !important; padding-top: 30px !important; padding-bottom: 20px !important; color: #89898D !important; } td[class*="mobile-padding"] { padding-top: 15px; } td[class*="outer-padding"] { padding: 0 !important; } td[class*="body-cell-left-pad"] { padding-left: 20px; padding-right: 20px; } } </style></head><body class="body" style="padding:0; margin:0; display:block; background:#ffffff; -webkit-text-size-adjust:none" bgcolor="#ffffff"> <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#ffffff"> <tr> <td class="outer-padding" valign="top" align="left"> <center> <table class="w320" cellspacing="0" cellpadding="0" width="600" height="723"> <tr> <td class="col-1 hide-for-mobile"> <table cellspacing="0" cellpadding="0" width="100%"> <tr> <td class="hide-for-mobile" style="padding:30px 0 10px 0;"> </td> </tr> <tr> <td class="hide-for-mobile" height="150" valign="top"> <b> <span></span> </b> <br> <span></span> </td> </tr> <tr> <td class="hide-for-mobile" style="height:180px; width:180px;"> </td> </tr> </table> </td> <td valign="top" class="col-2"> <table cellspacing="0" cellpadding="0" width="100%"> <tr> <td class="body-cell body-cell-left-pad" width="355" height="661" valign="top"> <table cellpadding="0" cellspacing="0"> <tr> <td width="355"> <h1> <span> Your Food Cart Purchase </span> </h1> <h2> <span></span> </h2> </td> </tr> <tr> <td width="355"> <table class="data-table" cellpadding="0" cellspacing="0"> <tr> <th> Purchased By </th> </tr> <tr> <td> <span> '+purchasedByText+' </span> </td> </tr> <tr> <th> Food Cart </th> </tr> <tr> <td> <span> '+foodCartName+' </span> </td> </tr> <tr> <th> Food Cart Date </th> </tr> <tr> <td> <span> '+foodCartDate+' </span> </td> </tr> <tr> <th> Items </th> </tr> <tr> <td> <span> '+itemNames+' </span> </td> </tr> <tr> <th> Amount Due </th> </tr> <tr> <td> <table cellpadding="0" cellspacing="0"> <tr> <td class="data-table-amount" style="width: 85px;"> <span> '+"$"+itemsTotal+' </span> </td> </tr> </table> </td> </tr> </table> </td> </tr> <tr> <td width="355" class="footer"> <center> <img width="213" height="48" src="'+thankYouSRC+'"> </center> </td> </tr> </table> <table cellspacing="0" cellpadding="0" width="100%"> <tr> <td class="hide-for-desktop-text"> <b> <span>Thanks for buying from JCC</span> </b> <br> <span><br></span> </td> </tr> </table> </td> </tr> </table> </td> </tr> </table> </center> </td> </tr> </table></body></html>'
  //var html = '<!DOCTYPE html><html><head><title>issa me mario</title><link href="https://fonts.googleapis.com/css?family=Roboto:400,700" rel="stylesheet"><style>.email-background {background: #eee !important;padding: 10px;}.email-container, .pre-header{max-width: 500px;margin: 0 auto;overflow: hidden;background: white;font-family: "Roboto", "Helvetica", sans-serif;text-align: center;}.pre-header{background: #eee;color: #eee;font-size: 2px;}/* overall styling --------------------- styling to elements that appear often (like <p>, <img>, etc) *//*IMG*/img { max-width: 100%; font-family: "Roboto", "Helvetica", sans-serif; color: #0043ff; margin: 0; }p {margin: 20px;font-size: 14px;font-width: 300;color: black;line-height: 1.5;}/*LIST makes lists look prettier*/li {line-height: 1.5;}/* styling by sections --------------------- specific "sections" of the receipt have different appearances for example, the footer has very specific styling that differs heavily from say the banner pic */ /*BANNER PIC I ultimately chose to just use an img instead of a div so this css is now defunct*//*.banner { background: url("https://drive.google.com/uc?id=1Q7drAndwZVwv4SGYvRDX9O979Ql7ewlz");background-size: cover;overflow: hidden;width: 100%;padding-top: 70%;text-align: center;font-size: 64px;color: white;line-height: 250px;}*/ /*SPECIAL grey background instead of white*/.special {padding: 20px;text-align: left;background-color: #f7f7f7;}.special h2 {margin-bottom: 0;}/*THANK YOU "For Jon Gao; Thank you for supporting this club!*/.thankYou {margin: 50px 0 50px 0;}/*FOOTER*/.footer {max-width: 500px;margin: 0 auto;overflow: hidden;padding: 20px 0 20px 0;background: #eee;font-family: "Roboto", "Helvetica", sans-serif;font-size: 10px;text-align: center;}/*BUTTON*/.butt {margin: 20px;}.butt p {text-decoration: none;display: inline-block;background: #cc6600;color: white;padding: 10px 20px;}.butt a { text-decoration: none; color: white; }.data-table {margin-top: 0;border-top: 2px #f7f7f7 solid;border-bottom: 2px #f7f7f7 solid;border-left: 4px #f7f7f7 solid;border-right: 4px #f7f7f7 solid;}.data-table th,.data-table td{width: 350px;height: 45px !important;line-height: 45px !important;padding-top: 2px !important;padding-bottom: 2px !important;padding-left: 10px !important;padding-right: 10px !important;background: #eee;border-top: 2px #f7f7f7 solid;border-bottom: 2px #f7f7f7 solid;text-align: left;}.data-table th {background-color: #eee;color: black;}.data-table td {}</style></head><body><!--structural formatting--><div class="email-background"><!--PRE HEADER in notifications, the first bit of text is shown so this makes sure that the text shown in the notification is ideal--><div class="pre-header"> Your JCC Food Cart order is in! </div> <div class="email-container"> <!--TITLE this is where it says "J C C" at the top--> <h1 style = "margin-top: 14px; margin-bottom: 14px;">J C C</h1> <table class="data-table" cellpadding="0" cellspacing="0"><tr> <td><span> '+foodCartName+' </span></td> <td style="text-align: right;"><span> '+foodCartDate+' </span></td></tr><tr> <td><span> '+itemNames+' </span></td> <td style="text-align: right;"><span> '+itemPricesText+' </span></td></tr><tr id="total"> <td><span> <b>Total</b> </span></td> <td style="text-align: right;"><span><b> $'+itemsTotal+' </b></span></td></tr></table> <!--THANK YOU--> <div class = "thankYou"> For '+purchasedByText+' <br><br></div> </div> <!--FOOTER If we have other events we are advertising, they can also go here :)--> <div class="footer"> Thank you for supporting JCC! </div></div></body></html>'
  var html = HtmlService.createHtmlOutputFromFile("html").getContent();
  //Logger.log(html)
  html = html.replace(new RegExp("foodCartName", 'g'), order.foodCartName);
  html = html.replace(new RegExp("itemNames", 'g'), order.itemNames);
  html = html.replace(new RegExp("foodCartDate", 'g'), order.foodCartDate);
  html = html.replace(new RegExp("itemPrices", 'g'), order.itemPricesText);
  html = html.replace(new RegExp("itemsTotal", 'g'), order.itemsTotal);
  html = html.replace(new RegExp("purchasedByText", 'g'), order.purchasedByText);
  
  var email = {
    address: order.emailAddress,
    subject: order.subject,
    htmlBody: html
  };
  
  if (order.itemsTotal == 0){
    //if someone checks out with a total of $0
    return;
  }
  var wasSent = formSheet.getRange(rowToProcess, sentCol, 1, 1).getValue();
  Logger.log(wasSent)
  if (wasSent){
    //catch to make sure email does not send twice.
    Logger.log("already sent");
    return;
  }
  
  if (order.itemsTotal >= 50){
    MailApp.sendEmail("phultquist@imsa.edu", organizationName+" Over $50 Alert: "+order.name+" | "+order.emailAddress+". Row "+rowToProcess+" of spreadsheet. Please confirm receipt.", "",{htmlBody: email.htmlBody});
    return;
  }
  
  formSheet.getRange(rowToProcess, sentCol, 1, 1).setValue("TRUE");
  
  return email;
}

function send(email){
  try{
    MailApp.sendEmail(email.address,email.subject,"",{htmlBody: email.htmlBody, inlineImages: null, attachments: null})
  } catch(e) {
    try{email.address}catch(er){return;} //if no email address return empty
    MailApp.sendEmail("phultquist@imsa.edu", "Delivery didnt work for "+email.address, "", {htmlBody: email.htmlBody}); 
  }
}

function imgSRC(imageKey){
  return 'cid:'+imageKey;
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
            itemName = itemName+" x"+itemQuantities[i-1];
            itemNames.push(itemName)
          }
        } else if (wantsItemsBought==false){
          itemNames.push(itemName)
        }
      }
  return itemNames;
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
    itemQuantities.push(itemQuantity);
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
  value = value.toString();
  if (value.indexOf("(")!=-1){
    value = value.substr(0, value.indexOf("(") - 1);
  }
  if (isNaN(value)==true || value.length == 0){
      if (value.toUpperCase() == "YES" || value.toUpperCase().indexOf("UNLIMITED") != -1) {
        value = 1
      }else{
        value = 0
      }
    }
  
  //makes quantity forced to be integer
  return parseInt(value);
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
      itemTotalQuantity = itemTotalQuantity + cellQuantity;
    }
    totalQuantities.push(itemTotalQuantity);
  }
  return totalQuantities;
}
 
