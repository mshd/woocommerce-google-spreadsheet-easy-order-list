// Author mshd94@gmail.com

function update(sheet_name) {

    var ck = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange("B4").getValue();

    var cs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange("B5").getValue();

    var website = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange("B3").getValue();


    var after = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange("E3").getValue();//.toISOString();
    var before = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange("E4").getValue();//.toISOString();

    var surl = website + "/wp-json/wc/v2/orders?consumer_key=" + ck + "&consumer_secret=" + cs  +"&per_page=100&after="+ after +"T00:00:00&before="+ before +"T00:00:00"; //+ "&filter[created_at_min]=" + n //"&after=2016-10-27T10:10:10Z"7
    var url = surl
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange("O1").setValue(url);
    var options =

        {
            "method": "GET",
            "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
            "muteHttpExceptions": true,

        };

    var result = UrlFetchApp.fetch(url, options);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange("A7:Z1000").clearContent();

    Logger.log(result.getResponseCode())
    
    if (result.getResponseCode() == 200) {

        var orders = JSON.parse(result.getContentText());

    }
    var doc = SpreadsheetApp.getActiveSpreadsheet();

    var temp = doc.getSheetByName(sheet_name);
    var consumption = {}

    arrayLength = orders.length
    var order = [];
    var line = 0;
    for (var i = 0; i < arrayLength; i++) {

  if(orders[i]["shipping_total"] != "0.00"){
    orders[i]["line_items"].push({
      product_id : 1,
      name: "Shipping",
      quantity: 1,
      tax_class : "",
      total_tax : orders[i]["shipping_tax"],
      total : orders[i]["shipping_total"],
      meta_data : []
    });
    }
    
        for (var k = 0; k < orders[i]["line_items"].length; k++) {
        
        
        var container = [];
        a = container.push(orders[i].id);
        a = container.push(orders[i]["status"]);

        a = container.push(orders[i]["date_created"]);

        a = container.push(orders[i]["billing"]["company"]);

        a = container.push(orders[i]["billing"]["first_name"]);

        a = container.push(orders[i]["billing"]["last_name"]);

        a = container.push(orders[i]["billing"]["address_1"] );

        a = container.push(orders[i]["billing"]["address_2"]);

        a = container.push(orders[i]["billing"]["postcode"]);

        a = container.push(orders[i]["billing"]["city"]);

        a = container.push(orders[i]["billing"]["country"]);

        a = container.push(orders[i]["billing"]["phone"]);

        a = container.push(orders[i]["billing"]["email"]);

       // a = container.push(orders[i]["total"]); //price

        a = container.push(orders[i]["payment_method"]);

        a = container.push(orders[i]["customer_note"]);

          container.push(orders[i]["line_items"][k]["product_id"])
          item = orders[i]["line_items"][k]["name"];
          container.push(item)
          qty = orders[i]["line_items"][k]["quantity"];
          container.push(qty)
          container.push(orders[i]["line_items"][k]["tax_class"])
          container.push(orders[i]["line_items"][k]["total_tax"])
          container.push(orders[i]["line_items"][k]["total"])
          container.push(parseFloat(orders[i]["line_items"][k]["total_tax"])+parseFloat(orders[i]["line_items"][k]["total"]))
          
          container.push(Math.round(orders[i]["line_items"][k]["total_tax"]/orders[i]["line_items"][k]["total"]*100))

                    //=ROUND(S8/T8*100)

          var metastuff = [];
          for(var m = 0;m < orders[i]["line_items"][k]["meta_data"].length; m++){
              metastuff.push(orders[i]["line_items"][k]["meta_data"][m]["value"]);
          }
          container.push(metastuff.join(" "));
          temp.appendRow(container);
}
    }
    
}
