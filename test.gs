function MakeSheets(){
  var ss = SpreadsheetApp.create('テスト');
  
  var this_sheet = SpreadsheetApp.getActiveSheet();
  
  var data_ss = SpreadsheetApp.openById('');
  var data_sheet = data_ss.getSheetByName('情報色々');
  var active_sheet = data_ss.getSheetByName('シート1');
  var last_row = data_sheet.getLastRow();
  var last_column = data_sheet.getLastColumn();
  
  var date_range_low = new Date('2019/8/1');
  var date_range_high = new Date('2019/8/31');
  var raw_data = data_sheet.getRange(2,1,last_row-1,last_column).getValues();
  var data = raw_data.filter(function(record){
    return (record[0] >= date_range_low) && (record[0] <= date_range_high);
  })
  
  var data_length = Object.keys(data).length;
  Logger.log(data);
  Logger.log('data length : '+data_length);
  Logger.log('row  length : '+last_row);
  Logger.log('colm length : '+last_column);
  Logger.log(data[0][1]);
  
  var company_dict = {};
  
  for(var i=0;i<data_length;i++){
    if(!company_dict[data[i][1]]){
      company_dict[data[i][1]] = {};
    }
    if(!company_dict[data[i][1]][data[i][0]]){
      company_dict[data[i][1]][data[i][0]] = { product : '' };
    }
    company_dict[data[i][1]][data[i][0]]['product'] += 　'\n' + data[i][3];      
  }
  Logger.log(company_dict);
  
  var i=0;
  var j=1;
  var products_name;
  for(　company_name in company_dict){
    Logger.log('company_name : ' + company_name);
    copied_sheet = this_sheet.copyTo(ss);
    copied_sheet.setName('test_'+i);
    j = 1;
    for( shipping in company_dict[company_name] ){
      products_name = company_dict[company_name][shipping]['product'];
      Logger.log('  shipping : '+ shipping);
      Logger.log('  product  : '+company_dict[company_name][shipping]['product']);
      copied_sheet.getRange(j,1).setValue(shipping);
      copied_sheet.getRange(j,2).setValue(products_name);
      j++;
    }
    i++;
  }
}

function myFunction() {
  MakeSheets();
}
