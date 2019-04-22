var analyticsTitleColumn = 1,
    quantityColumn = 2,
    priceEachColumn = 5,
    extraColumn = 3,
    totalQuantityColumn = 4
    totalColumn = 6,
    headerColumn = 1;

var headerRow = 1

var prices = getPrices()
var totalQuantities = getTotalQuantities()

var columnHeaders = {
  quantityColumn : "Quantity",
  priceEachColumn : "Price Each",
  extraColumn : "Recomended Extras",
  totalQuantityColumn : "Total Quantities",
  totalColumn : "Pre-Order Revenue",
  copyrightMessage : "©Patrick Hultquist"
}
    
var analyticsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Analytics")

function updateAnalyticsSheet() {
 updateSettingsSheet()
 updateColumnHeaders()
 updateHeaders()
 updateQuantities()
 updateExtras()
 updateTotalQuantities()
 updatePrices()
 updateTotal() 
 SpreadsheetApp.getActiveSpreadsheet().toast("Analytics sheet was just updated. It'll update again in a minute.", "Woohoo!")
}

function updateHeaders() {
 var itemNames = getNames(false, 1, true)
  var data = []
  for (k in itemNames){
   data.push([itemNames[k]]) 
  }
  var titleRowRange = analyticsSheet.getRange(headerColumn+1,analyticsTitleColumn, itemNames.length, 1)
 titleRowRange.setValues(data)
 titleRowRange.setFontWeight("bold")
}

function updateExtras(){
  var extras = getExtras()
  var extraRowRange = analyticsSheet.getRange(headerColumn+1,extraColumn, extras.length, 1)
  extraRowRange.setValues(extras)
}

function getExtras(){
  //var extraRowRange = analyticsSheet.getRange(headerColumn+1, extraColumn, totalQuantities.length, 1)
  var data = []
  for (k in totalQuantities){
   data.push([extra(totalQuantities[k])]) 
  }
  return data
}

function updateQuantities() {
  var quantityRowRange = analyticsSheet.getRange(headerColumn+1, quantityColumn, totalQuantities.length, 1)
  var data = []
  for (k in totalQuantities){
   data.push([totalQuantities[k]]) 
  }
  quantityRowRange.setValues(data)
}

function updateTotalQuantities() {

  var totalQuantitiesRange = analyticsSheet.getRange(headerColumn+1, totalQuantityColumn, totalQuantities.length, 1)
  var extras = []
  var quantitySums = []
  for (q in getExtras()){
    extras.push(getExtras()[q])
  }
  for (i in totalQuantities){
   quantitySums.push([parseInt(totalQuantities[i])+parseInt(extras[i])]) 
  }
  totalQuantitiesRange.setValues(quantitySums)
}

function updatePrices() {
 var pricesRowRange = analyticsSheet.getRange(headerColumn+1, priceEachColumn, prices.length, 1)
 var data = []
 for (k in prices){
   data.push([prices[k]]) 
 }
 pricesRowRange.setValues(data)
}

function updateTotal() {
  var totals = []
  for(u=0;u<numberOfItems;u++){
    var price = prices[u]
    var totalQuantity = totalQuantities[u]
    var total = price*totalQuantity
    totals.push([total])
  }
  var totalsRange = analyticsSheet.getRange(headerColumn+1,totalColumn, totals.length, 1)
  var totalHeader = analyticsSheet.getRange(headerColumn,totalColumn, totals.length, 1)
  totalsRange.setValues(totals)
  totalsRange.setBorder(false, true, false, false, false, false)
  totalHeader.setBorder(false, true, false, false, false, false)
}



function extra(x){
 return Math.floor(0.28674*x) 
}

function modifyPieChart(){
  //not in use yet; waiting to be updated
  var range = analyticsSheet.getRange(1,1,analyticsSheet.getLastRow()-1,analyticsSheet.getLastColumn()-1)
  var chart = analyticsSheet.getCharts()[0];
  var ranges = chart.getRanges()


  chart = chart.modify()
  .asPieChart()
  .addRange(range)
  .setOption('title', 'Purchase Share')
 // .setPosition(10,2,0,0)
  .setLegendPosition(Charts.Position.RIGHT)
  .setOption('legend', {position: 'right', textStyle: {color: 'black', fontSize: 9}})
  .build();

  //analyticsSheet.updateChart(chart);
}


function updateColumnHeaders(){
  analyticsSheet.getRange(headerRow, quantityColumn).setValue(columnHeaders.quantityColumn)
  analyticsSheet.getRange(headerRow, priceEachColumn).setValue(columnHeaders.priceEachColumn)
  analyticsSheet.getRange(headerRow, totalColumn).setValue(columnHeaders.totalColumn)
  analyticsSheet.getRange(headerRow, extraColumn).setValue(columnHeaders.extraColumn)
  analyticsSheet.getRange(headerRow, totalQuantityColumn).setValue(columnHeaders.totalQuantityColumn)
  analyticsSheet.getRange(headerRow, analyticsTitleColumn).setValue(columnHeaders.copyrightMessage)
}