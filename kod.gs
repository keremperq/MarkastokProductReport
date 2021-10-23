
function mainFunction(){

  var main_url = "https://www.markastok.com"

  //take spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //get sheet by sheet name
  var first = ss.getSheetByName("MarkastokTable");
  first.clear(); //clear all data
  //get paths from "urls" sheet
  var url_paths = ss.getSheetByName("urls").getDataRange().getValues();
  //set metrics to first row
  first.appendRow(["URL", "SKU", "Product Name", "Availability", "Product Price", "Offer", "Sale Price", "Product Code"]);
  
  //set first row frozen and bold
  first.setFrozenRows(1);
  first.getRange(first.getFrozenRows(), 1, 1, first.getMaxColumns()).activate();
  first.getActiveRangeList().setFontWeight('bold');
  
  //apply operations to all paths
  for(var i = 0; i < 200;i++){
    //we should url_paths.length instead of 200 but occurs execution time error so i used 200 path
    var new_url = main_url + url_paths[i];
    new_url = new_url.replace(",,,,,,,",""); //remove redundant commas
    
    //Cheerio library imported
    const content = getContent_(new_url); //fetching url
    const $ = Cheerio.load(content);

    //check whether url is product page or not
    var og_type = $('meta[property="og:type"]').attr('content');

    if(og_type == "product"){

      var product_price = $('.currencyPrice.discountedPrice').text();
      product_price = product_price.replace('  TL',''); //product price
      var sale_price = $('.product-price').text(); //sale price
      $('.fbold').remove();//remove brand
      var product_name = $(".product-name").text(); //product name
      var new_product_name = product_name.replace('\n','');
      new_product_name = new_product_name.trim();
      var data_id = $("#product-name").attr("data-id"); //data id (SKU)
      var discount = $(".detay-indirim").text(); // offer
      var product_code = $("#urun-sezon").attr("value"); //product code
      //removed unnecessary childs from main class and just parse the text
      var product_code2 = $(".product-feature-content").children().remove().end().text().trim(); 
      var sizes_passive = $('a.col.box-border.passive').length;
      var sizes_all = $('a.col.box-border').length;
      var availability = ((sizes_all - sizes_passive) / sizes_all) * 100;
      //to sort availability ascending order, i changed to availability float to string
      //replaced dot with comma and added percentage to end of the string
      availability = availability.toFixed(2).toString().replace(".",",") + "%";
      
      if(product_code == ""){
        product_code = product_code2;
      }
      var values = [new_url,data_id,new_product_name,availability,product_price,discount,sale_price,product_code];
      
      
      ss.appendRow(values);
      
    }
    
    
  }
  //sorting as availability
  ss.sort(4);
  sendEmail();
}

function sendEmail() {

  var ssID = SpreadsheetApp.getActiveSpreadsheet().getId();

  var sheetName = SpreadsheetApp.getActiveSpreadsheet().getName();

  var email_ID = "okan@analyticahouse.com";
  var subject = "Kerem Berk Güçlü";
  var body = "";

  var requestData = {"method": "GET", "headers":{"Authorization":"Bearer "+ ScriptApp.getOAuthToken()}};
  var shID = getSheetID("MarkastokTable") //Get Sheet ID of sheet name
  var url = "https://docs.google.com/spreadsheets/d/"+ ssID + "/export?format=xlsx&id="+ssID+"&gid="+shID;

  var result = UrlFetchApp.fetch(url , requestData);  
  var contents = result.getContent();

  MailApp.sendEmail(email_ID,subject ,body, {attachments:[{fileName:sheetName+".xls", content:contents, mimeType:"application//xls"}]});

};

function getSheetID(name){
 var ss = SpreadsheetApp.getActive().getSheetByName(name)
 var sheetID = ss.getSheetId().toString() 
 return sheetID
}


function getContent_(url) {
  options = {muteHttpExceptions:true};
  return UrlFetchApp.fetch(url,options).getContentText()
}