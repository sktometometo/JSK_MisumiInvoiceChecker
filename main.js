var id_sheet_of_orderlist = "";
var checkaddress = {"a@hoge.fuga","b@foo.bar"};

function main( target_date ) {
  if ( target_date == undefined ) {
    target_date = new Date();
  }
  
  var hoge = checkMisumiInvoice( target_date );
  var mailbody_1 = hoge[0];
  var mailbody_2 = hoge[1];
  var address = checkaddress; // チェック結果の送信先アドレス
  var title   = "[部品係] " + target_date.getFullYear() + "年 " 
                           + (target_date.getMonth()+1) + "月 "
                           + "MISUMI注文チェック";
  MailApp.sendEmail(address, title + "_1", mailbody_1);
  MailApp.sendEmail(address, title + "_2", mailbody_2); 
}

function dumpArray( array ) {
  var ans_str = "";
  var indent_str = "    ";
  for ( var i=0,l=array.length; i<l; i++ ) {
    ans_str += indent_str + array[i][0] + ",  " + array[i][1] + "\n";
  }
  return ans_str;
}

function checkMisumiInvoice( target_date ) {
  if ( target_date == undefined ) {
    target_date = new Date();
  }
  
  var array_Mail = getNameAndDayFromInvoiceMessage( target_date );
  var array_SS   = getEntryMisumi( target_date );
  array_Mail.sort(function(a,b){
    if( a[0] < b[0] ) return -1;
    if( a[0] > b[0] ) return  1;
    if( a[0] ==b[0] ) return a[1]<b[1];
  });
  array_SS.sort(function(a,b){
    if( a[0] < b[0] ) return -1;
    if( a[0] > b[0] ) return  1;
    if( a[0] ==b[0] ) return a[1]<b[1];
  });
  
  Logger.log("Mail:");
  Logger.log("\n"+dumpArray(array_Mail));
  Logger.log("Spreadsheet:");
  Logger.log("\n"+dumpArray(array_SS));
  
  Logger.log("Mail での MISUMI請求書の枚数 :" + array_Mail.length);
  Logger.log("Spreadsheet でのMISUMIの件数:" + array_SS.length);
  
  
  // 
  var mailbody_1 = "";
  
  mailbody_1 += "Mail での MISUMI請求書の枚数 :" + array_Mail.length + "\n";
  mailbody_1 += "Spreadsheet でのMISUMIの件数:" + array_SS.length + "\n";
  mailbody_1 += "\n";
  mailbody_1 += "== メールで確認したMISUMIの注文リスト ==\n";
  mailbody_1 += dumpArray(array_Mail);
  mailbody_1 += "\n";
  mailbody_1 += "== スプレッドシートで確認したMISUMIの注文リスト ==\n";
  mailbody_1 += dumpArray(array_SS);
  mailbody_1 += "\n";
  
  // 
  var mailbody_2 = "";  
  var array_name = [];
  
  for ( var i=0,l=array_Mail.length; i<l; i++ ) {
    array_name.push( array_Mail[i][0] );
  }
  for ( var i=0,l=array_SS.length; i<l; i++ ) {
    array_name.push( array_SS[i][0] );
  }
  array_name.sort(function(a,b){
    if( a < b  ) return -1;
    if( a > b  ) return  1;
    if( a == b ) return  0;
  });
  array_name = array_name.filter(function (x, i, self) { return self.indexOf(x) === i; });
  
  for ( var i=0,l_name=array_name.length; i<l_name; i++ ) {
    mailbody_2 += "氏名:" + array_name[i] + "\n";
    mailbody_2 += "  " + "メールで届いた請求書" + "\n";
    for ( var j=0,l=array_Mail.length; j<l; j++ ) {
      if ( array_name[i] == array_Mail[j][0] ) {
        mailbody_2 += "    " + array_Mail[j][1].toLocaleString("ja-JP") + "\n";
      }
    }
    mailbody_2 += "  " + "Spreadsheetで確認できる注文書" + "\n";
    for ( var j=0,l=array_SS.length; j<l; j++ ) {
      if ( array_name[i] == array_SS[j][0] ) {
        mailbody_2 += "    " + array_SS[j][1].toLocaleString("ja-JP") + "\n";
      }
    }
    mailbody_2 += "\n";
  }
  
  return [mailbody_1, mailbody_2];
}

function getEntryMisumi( target_date ) {
  // 注文書のシートから, target_dateの月と, 納品日の月が同じMISUMIの注文を取得する
  var array_ans = [];
  
  if ( target_date == undefined ) {
    target_date = new Date();
  }
  var id_sheet = id_sheet_of_orderlist;
  var target_sheetname = "注文リスト";
  
  var companyname_candidates = [ "ミスミ", "MISUMI" ];
  
  var ss = SpreadsheetApp.openById( id_sheet );
  var sheet = ss.getSheetByName(target_sheetname);

  var rowindex_first = 3;
  var rowindex_last  = sheet.getLastRow();
  var colindex_first = 1;
  var colindex_last  = 10;
  
  var colindex_deliverydate = 2;
  var colindex_company = 3;
  var colindex_name = 8;
  
  var values = sheet.getRange( rowindex_first
                              ,colindex_first
                              ,rowindex_last - rowindex_first + 1
                              ,colindex_last - colindex_first + 1).getValues();
  
  for ( var i=0; i<rowindex_last-rowindex_first+1; i++ ) {
    
    // 納品日が空欄の行は読み飛ばす
    if ( values[i][colindex_deliverydate-1] == "" ) continue;    
    
    var date    = new Date(values[i][colindex_deliverydate-1]);
    var company = values[i][colindex_company-1];
    var name    = values[i][colindex_name-1];
    
    // 日付のチェック
    if ( date.getFullYear() != target_date.getFullYear() || date.getMonth() != target_date.getMonth() ) {
      continue;
    } else {
      // Logger.log("INFO:月の一致を確認");
    }
    
    // 注文先のチェック
    var flag = false;
    for ( var j=0,lj=companyname_candidates.length; j<lj; j++ ) {
      if ( company.indexOf(companyname_candidates[j]) != -1 ) {
        flag = true;
        break;
      }
    }
    if ( !flag ) {
      // Logger.log("INFO:注文先が一致しない:"+company);
      continue;
    } else {
      Logger.log("INFO:注文先の一致を確認");
      array_ans.push([name,date]);
    }     
  }
    
  // Logger.log("array_ans:");
  // Logger.log(array_ans);
  
  return array_ans;
}

function getNameAndDayFromInvoiceMessage( target_date ) {
  var array_ans = [];
  var messages = findMisumiInvoice( target_date );
  for ( var i=0,l=messages.length; i<l; i++ ) {
    var body = messages[i].getBody();
    var date = new Date(messages[i].getDate());
    var firstline = body.split("\n")[1];
    var names = firstline.split(" ");
    var name = names[names.length-2];
    //Logger.log(name);
    array_ans.push([name,date]);
  }
  // Logger.log(array_ans);
  return array_ans;
}

function findMisumiInvoice( target_date ) {
  // date と同じ月の Invoice の Message を取得する
  if ( target_date == undefined ) {
    target_date = new Date();
  }
  // Logger.log("target_date:" + target_date);
  
  var array_ans = [];
  var subjectMisumi = '（株）ミスミより請求書発行のご案内';
  var addressMisumi = "urikake2@misumi.co.jp";
  
  var threads_invoice = GmailApp.search(subjectMisumi);
  var messages_invoice = GmailApp.getMessagesForThreads(threads_invoice);
  
  for ( var i=0, li=messages_invoice.length; i<li; i++ ) {
    for ( var j=0, lj=messages_invoice[i].length; j<lj; j++ ) {
      var message = messages_invoice[i][j];
      var message_date = message.getDate();
      var message_from = message.getFrom();
      
      // Misumi からでないものはとばす
      if ( message_from.indexOf(addressMisumi) == -1 ) {
        continue;
      }
        
      // target_date と同じ年,月のmessage を array_ans に追加
      if ( message_date.getFullYear() == target_date.getFullYear()
        && message_date.getMonth() == target_date.getMonth() ) {
          array_ans.push( message );
        }
    }
  }
  // Logger.log("array ans length: " + array_ans.length);
  return array_ans;
}
