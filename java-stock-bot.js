function doPost(e) {
    var data = JSON.parse(e.postData.contents)
    var gg = data.originalDetectIntentRequest.payload.data.message.text;  
    var ss = SpreadsheetApp.openByUrl('')
    var sh = ss.getSheetByName('')
      var today = Utilities.formatDate(new Date(),"GMT+7","dd")
      var todayline = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/YYYY")
      var num = today; // "82144251"
      num = parseFloat(num);
      //console.log(num, typeof num); 
      //var numAsNumber = Number(num);
      num1=num+3;
      num2=num+2;
      //const gg = 1;
      if (gg==1){
        var g1 = gg + 7;
      }else if (gg==2){
        var g2 = gg + 15;
      }else{
        Logger.log('wow');
      }
      var values = sh.getRange(7,1, sh.getLastRow(), sh.getLastColumn()).getValues();
      for (i =0;i<values.length;i++){
        //Logger.log(i);
        if(i==g1 || i==g2){
          var cc = sh.getRange(i,2).getValue();
          Logger.log(cc);
           if(k=num1+1){
              //k=k+3;
              //Logger.log(k)
              //var qw1 = sh.getRange(i+1,k-1).getValue()+'\n'
              //var qw2 = sh.getRange(i+2,k-1).getValue()+'\n'
              //var qw3 = sh.getRange(i+3,k-1).getValue()+'\n'
              var qw4 = sh.getRange(i+1,k).getValue()+'\n'
              var qw5 = sh.getRange(i+2,k).getValue()+'\n'
              var qw6 = sh.getRange(i+3,k).getValue()+'\n'
    
              //_________กรณีที่ยังไม่มีการจัดส่ง________________________________________//
              if (qw5==0){
              var ok = 'No Delivery Yet \n'
                 //Logger.log(num)
              //var x = num2-3
              //var nday = 'Yesterday วันที่  ' +x
              //var perd = 'inventory : ' +qw1+ 'Delivery Plan : ' +qw2+ 'Production Plan : ' +qw3
              var y = num1-3
              var lastnday = 'Today วันที่ ' +y
              var lastd = 'Delivery Plan : ' +ok+ 'inventory : ' +qw4+ 'Production Plan : ' +qw6
              //Logger.log(nday)
              //Logger.log(perd)
              Logger.log(lastnday)
              Logger.log(lastd)
              
              Logger.log('ข้อมูล ณ วันที่ '+todayline)
    
              }else{
              //Logger.log(num)
              //var x = num2-3
              //var nday = 'Yesterday วันที่ ' +x
              //var perd = 'inventory : ' +qw1+ 'Delivery Plan: ' +qw2+ 'Production Plan: ' +qw3
              var y = num1-3
              var lastnday = 'Today วันที่ ' +y
              var lastd = 'Delivery Plan : ' +qw5+ 'inventory : ' +qw4 +'Production Plan : ' +qw6
              //Logger.log(nday)
              //Logger.log(perd)
              Logger.log(lastnday)
              Logger.log(lastd)
              Logger.log('ข้อมูล ณ วันที่ '+todayline)
              }
            }else{Logger.log('No Log')}
        }
      }
    }
    
    