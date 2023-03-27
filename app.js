const jsoncsv = require('json-csv');
const papa = require('papaparse');
const request = require('request');
const fs = require('fs');
const { config } = require('process');
const { csv } = require('json-csv');
const itemTelegram = require('./component/ITEM00011.json');

const xlsx = require("xlsx");
const workbook = xlsx.readFile("./item.xls");
let worksheet = {};
for(const sheetName of workbook.SheetNames){

    worksheet[sheetName] = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
}
const table = worksheet.item;
  var i = 0 ;
  var telegram =[];


    table.forEach( element => {

         Template  = {...itemTelegram.ITEM00011};
         Template2 = {...itemTelegram.ITEMDESC00005};
         Template3 = {...itemTelegram.ITEMQTYUNIT00007};
         Template4 = {...itemTelegram.ITEMALIAS00006};


        let hostTelegram = element;   

        Template.itemNo = Template2.itemNo = Template3.itemNo  =Template4.itemNo  = hostTelegram.itemNo 
        Template.variant = Template2.variant = Template3.variant = Template4.variant  = hostTelegram.variant;
        Template3.qtyUnit = Template3.referenceQtyUnit =Template.baseQtyUnit = Template.whQtyUnit = hostTelegram.unit;
        Template.goodsCategory = hostTelegram.goodsCategory

        console.log(Template);

        if(hostTelegram.assortment){
            Template.assortment = hostTelegram.assortment;
        }

        if(hostTelegram.Gross_length_CM){
            if(isValid(hostTelegram.Gross_length_CM)){
                Template3.length = hostTelegram.Gross_length_CM
            }
        }
        if(hostTelegram.Gross_width_CM){
            if(isValid(hostTelegram.Gross_width_CM)){
                Template3.width = hostTelegram.Gross_width_CM
            }
        }
        if(hostTelegram.Gross_height_CM){
            
            if(isValid(hostTelegram.Gross_height_CM)){
                Template3.height = hostTelegram.Gross_height_CM
            }
        }
        if(hostTelegram.Gross_cubic_meter_CBM){
            console.log(isValid(hostTelegram.Gross_cubic_meter_CBM))
            if(isValid(hostTelegram.Gross_cubic_meter_CBM)){
                Template3.volume = hostTelegram.Gross_cubic_meter_CBM
            }
        }
        if(hostTelegram.Gross_weight_KG){
            if(isValid(hostTelegram.Gross_weight_KG)){
                Template3.grossWeight = hostTelegram.Gross_weight_KG
            }
        }

        if(hostTelegram.Barcode){
            if(isValid(hostTelegram.Barcode)){
                Template4.itemAliasNo = hostTelegram.Barcode; 
            }
        }

        if(hostTelegram.goodsCategory === 'FG'){
            Template.bbDateStockMode =  Template.igBbDateRegistration =Template.ogBbDateRegistration =  "OPTIONAL"
            Template.csiaStockMode01 = Template.csiaIgResMode01 = "MANDATORY"
            Template.csiaOgResMode01 = "OPTIONAL"

            Template.csiaStockMode02 = Template.csiaIgResMode02 = "MANDATORY"
            Template.csiaOgResMode02 = "OPTIONAL"
            
            Template.csiaStockMode03 = Template.csiaIgResMode03 = Template.csiaOgResMode03 = "OPTIONAL"
            Template.csiaStockMode04 = Template.csiaIgResMode04 = Template.csiaOgResMode04 = "OPTIONAL"
            Template.csiaStockMode05 = Template.csiaIgResMode05 = Template.csiaOgResMode05 = "OPTIONAL"
            Template.csiaStockMode06 = Template.csiaIgResMode06  = "OPTIONAL"
            Template.csiaOgResMode06 = "NOT_APPLICABLE"

        }
        if(hostTelegram.goodsCategory === 'Label'){
            Template.csiaStockMode01 = Template.csiaIgResMode01 = "MANDATORY"
            Template.csiaOgResMode01 = "OPTIONAL"
            
            Template.csiaStockMode04 = Template.csiaIgResMode04 = Template.csiaOgResMode04 = "OPTIONAL"
            Template.csiaStockMode05 = Template.csiaIgResMode05 = Template.csiaOgResMode05 = "OPTIONAL"

        }
        if(hostTelegram.goodsCategory === 'Packaging'){
            Template.csiaStockMode01 = Template.csiaIgResMode01 = "MANDATORY"
            Template.csiaOgResMode01 = "OPTIONAL"

            Template.csiaStockMode04 = Template.csiaIgResMode04 = Template.csiaOgResMode04 = "OPTIONAL"
            Template.csiaStockMode05 = Template.csiaIgResMode05 = Template.csiaOgResMode05 = "OPTIONAL"
        }
        if(hostTelegram.goodsCategory === 'Pre-FG'){
            Template.bbDateStockMode =  Template.igBbDateRegistration =Template.ogBbDateRegistration =  "OPTIONAL"
            Template.csiaStockMode01 = Template.csiaIgResMode01 = Template.csiaOgResMode01 ="MANDATORY"


            
            Template.csiaStockMode03 = Template.csiaIgResMode03 = Template.csiaOgResMode03 = "OPTIONAL"
            Template.csiaStockMode04 = Template.csiaIgResMode04 = Template.csiaOgResMode04 = "OPTIONAL"
            Template.csiaStockMode05 = Template.csiaIgResMode05 = Template.csiaOgResMode05 = "OPTIONAL"
            
        }
        if(hostTelegram.goodsCategory === 'TVA'){
            Template.bbDateStockMode =  Template.igBbDateRegistration =Template.ogBbDateRegistration =  "OPTIONAL"
            Template.csiaStockMode01 = Template.csiaIgResMode01 = "MANDATORY"
            Template.csiaOgResMode01 = "OPTIONAL"

            Template.csiaStockMode02 = Template.csiaIgResMode02 = "MANDATORY"
            Template.csiaOgResMode02 = "OPTIONAL"
            
            Template.csiaStockMode04 = Template.csiaIgResMode04 = Template.csiaOgResMode04 = "OPTIONAL"
            Template.csiaStockMode05 = Template.csiaIgResMode05 = Template.csiaOgResMode05 = "OPTIONAL"
            Template.csiaStockMode06 = Template.csiaIgResMode06  = "OPTIONAL"
            Template.csiaOgResMode06 = "NOT_APPLICABLE"
        }


        Template.Serialnumber = i++ ;
        Template2.Serialnumber = i++;
        Template3.Serialnumber = i++;
        Template4.Serialnumber = i++;

        let Des1 = hostTelegram.ItemDescription;
        Template2.Description1= Des1.substring(0,40);
        if( Des1.length > 40) {
            Template2.Description2 = Des1.substring(40, Des1.length);
        }
        
        var csvData1 = [Template];
        var csvData2 = [Template2];
        var csvData3 = [Template3];
        var csvData4 = [Template4];
        // console.log(csvData3);

        var csv1 = papa.unparse(csvData1);
        var csv2 = papa.unparse(csvData2);
        var csv3 = papa.unparse(csvData3);
        var csv4 = papa.unparse(csvData4);

        var data;
        let hasBarcode = hostTelegram.Barcode =! null && hostTelegram.Barcode != "-" && hostTelegram.Barcode

        if(hasBarcode){
            data = csv1+"\r"+csv2+"\r"+csv3+"\r"+csv4+"\r";
        }else data = csv1+"\r"+csv2+"\r"+csv3+"\r";

        telegram.push(data);
    });
    
    var totalMessage = getAllArray(telegram);
    var timeInMs = Date.now();
    fs.writeFile("./in/"+timeInMs+".csv", totalMessage, function(err) {
    if(err) {
        return console.log(err);
    }
    
    console.log("The file was conveted!");
}); 



  
  


function getAllArray(array) {
     data = "";
    for (var i=0; i < array.length ;i++){
        data += array[i];
    }
    return data;
};





function isValid(param){
    if(param != "-" && param != null ){
            return true
}else return false
}