// ************ SETTINGS ************
const SHEET_NAME = "Orders"; // Change if your sheet name is different

// **********************************

function doGet(e) {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Customer Feedback")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Get customer details
function getCustomer(orderNo){

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  const data = sheet.getDataRange().getValues();

  for(var i=1;i<data.length;i++){

    if(data[i][0].toString().trim()==orderNo){

      return{
        status:true,
        row:i+1,
        order:data[i][0],
        customer:data[i][1]
      }

    }

  }

  return{
    status:false
  };

}


// Save Feedback

function saveFeedback(obj){

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  const data = sheet.getDataRange().getValues();

  for(var i=1;i<data.length;i++){

    if(data[i][0]==obj.orderNo){

      sheet.getRange(i+1,4).setValue(obj.rating);

      sheet.getRange(i+1,5).setValue(obj.experience);

      sheet.getRange(i+1,6).setValue(obj.suggestion);

      sheet.getRange(i+1,7).setValue(new Date());

      return true;

    }

  }

  return false;

}
