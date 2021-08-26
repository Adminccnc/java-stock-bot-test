function doPost(e) {
  var data = JSON.parse(e.postData.contents)
  var gg = data.originalDetectIntentRequest.payload.data.message.text;  
  var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1JpKG_CKLKdz4YR2M9Qo2_lqP97AdLU8ZQWsej_o3yJQ/edit')
  var sh = ss.getSheetByName('July-21(New)')
    var today = Utilities.formatDate(new Date(),"GMT+7","dd")
    var todaynow = Utilities.formatDate(new Date(),"GMT+7","dd/MM")
    var todayline = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/YYYY")
    var num = today; // "82144251"
    num = parseFloat(num);
    //console.log(num, typeof num); 
    //var numAsNumber = Number(num);
    num1=num+3;
    num2=num+2;
    /* const gg = 1;
    if (gg==1){
      var g1 = gg + 7;
    }else if (gg==2){
      var g2 = gg + 15;
    }else{
      Logger.log('wow');
    } */
    //var gg = 2;
    var values = sh.getRange(7,1, sh.getLastRow(), sh.getLastColumn()).getValues();
    for (i =0;i<values.length;i++){
      //Logger.log(i);
      if(i==gg){
        if(i==1){
          i=i+7;
        }else if(i==2){
          i=i+15;
        }else if(i==3){
          i=i+21;
        }else if(i==4){
          i=i+27;
        }else if(i==5){
          i=i+33;
        }else if(i==6){
          i=i+39;
        }else{
          return;
        }
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
            var qw5 = 'No Delivery Yet'
               //Logger.log(num)
            //var x = num2-3
            //var nday = 'Yesterday วันที่  ' +x
            //var perd = 'inventory : ' +qw1+ 'Delivery Plan : ' +qw2+ 'Production Plan : ' +qw3
            var y = num1-3
            /*var lastnday = 'Today วันที่ ' +y
            var lastd = 'Delivery Plan : ' +ok+ 'inventory : ' +qw4+ 'Production Plan : ' +qw6
            //Logger.log(nday)
            //Logger.log(perd)
            Logger.log(lastnday)
            Logger.log(lastd)
            
            Logger.log('ข้อมูล ณ วันที่ '+todayline)*/
  
            }else{
            //Logger.log(num)
            //var x = num2-3
            //var nday = 'Yesterday วันที่ ' +x
            //var perd = 'inventory : ' +qw1+ 'Delivery Plan: ' +qw2+ 'Production Plan: ' +qw3
            var y = num1-3
            var lastnday = 'Today วันที่ ' +y
            var lastd = 'Delivery Plan : ' +qw5+ 'inventory : ' +qw4+'Production Plan : ' +qw6
            //Logger.log(nday)
            //Logger.log(perd)
            Logger.log(lastnday)
            Logger.log(lastd)
            Logger.log('ข้อมูล ณ วันที่ '+todayline)
            }var result = {
      "fulfillmentMessages": [
    {
      "platform": "line",
      "type": 4,
      "payload" : {
      "line":  {
      "type": "flex",
    "altText": "ข้อความตอบกลับจาก Stock",
    "contents": 
  {
    "type": "bubble",
    "size": "giga",
    "styles": {
      "footer": {
        "separator": true
      }
    },
    "body": {
      "type": "box",
      "layout": "vertical",
      "contents": [
        {
          "type": "text",
          "text": "คลังสินค้า",
          "weight": "bold",
          "color": "#de6825",
          "size": "xl",
          "align":"center"
        },
        {
          "type": "image",
          "url": "https://thecoffeenery.com/wp-content/uploads/2021/08/Logo_PANA_Coffee_new.jpg",
          "size": "lg",
          "aspectRatio": "2:1",
          "align":"center"
        },
        {
          "type": "text",
          "text": ""+cc,
          "weight": "bold",
          "size": "sm",
          "margin": "md",
          "align":"center"
        },
        {
          "type": "text",
          "text": "compunnyname",
          "size": "xs",
          "color": "#aaaaaa",
          "align":"center",
          "wrap": true
        },
        {
          "type": "separator",
          "margin": "xs",
        },
        {
          "type": "box",
          "layout": "vertical",
          "margin": "xxl",
          "spacing": "sm",
          "contents": [
            {
              "type": "box",
              "layout": "horizontal",
              "contents": [
                {
                  "type": "text",
                  "text": "วันที่ : "+todaynow,
                  "size": "sm",
                  "color": "#555555",
                },
                /*{
                  "type": "text",
                  "text": ""+y,
                  "size": "sm",
                  "color": "#111111",
                  "align": "end"
                }*/
              ]
            },
            {
            "type": "separator",
             "margin": "sm"
            },
            {
              "type": "box",
              "layout": "horizontal",
              "contents": [
                {
                  "type": "text",
                  "text": "Delivery Plan : ",
                  "size": "sm",
                  "color": "#555555",
    
                },
                {
                  "type": "text",
                  "text": ""+qw5,
                  "size": "sm",
                  "color": "#111111",
                  "align": "end"
                }
              ]
            },
            {
              "type": "box",
              "layout": "horizontal",
              "contents": [
                {
                  "type": "text",
                  "text": "inventory : ",
                  "size": "sm",
                  "color": "#555555",
     
                },
                {
                  "type": "text",
                  "text": ""+qw4,
                  "size": "sm",
                  "color": "#111111",
                  "align": "end"
                }
              ]
            },
               {
              "type": "box",
              "layout": "horizontal",
              "contents": [
                {
                  "type": "text",
                  "text": "Production Plan : ",
                  "size": "sm",
                  "color": "#555555",
     
                },
                {
                  "type": "text",
                  "text": ""+qw6,
                  "size": "sm",
                  "color": "#111111",
                  "align": "end"
                }
              ]
            },
            {
            "type": "separator",
             "margin": "sm"
            },
              {
              "type": "box",
              "layout": "horizontal",
              "contents": [
                {
                  "type": "text",
                  "text": "ข้อมูล ณ วันที่ ",
                  "size": "sm",
                  "color": "#555555",
     
                },
                {
                  "type": "text",
                  "text": ""+todayline,
                  "size": "sm",
                  "color": "#111111",
                  "align": "end"
                }
              ]
            },
            {
              "type": "separator",
              "margin": "xl"
            },
  
          ]
        },
   
        {
          "type": "box",
          "layout": "horizontal",
          "margin": "md",
          "contents": [
            {
              "type": "text",
              "text": "ผู้ดูแลสินค้าคงคลัง",
              "size": "xs",
              "color": "#aaaaaa",
    
            },
            {
              "type": "text",
              "text": "name stock keeper",
              "color": "#aaaaaa",
              "size": "xs",
              "align": "end"
            }
          ]
        },
      ]
    }, 
    "footer": {
          "type": "box",
          "layout": "horizontal",
          "contents": [
            {
              "type": "button",
              "style": "primary",
              "action": {
                "type": "uri",
                "label": "ดูรายละเอียดเพิ่มเติมทั้งหมด",
                "uri": "https://thecoffeenery.com"
              }
            }
          ]
        }
  }
  }
          
     }
    }
   ]
  }
      Logger.log(result);
      var replyJSON = ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
      //Logger.log(replyJSON);
      return replyJSON;
        }else{Logger.log('No Log')}
      }
    }
  }
  
  
