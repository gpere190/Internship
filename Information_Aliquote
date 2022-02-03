// Author: Gricelle Perez
//Purpose: The purpose of this code is to transfer the information of cells from one spreadsheet to another spreadsheet depending on
//on the users input on the spreadsheet.

var mainsource1 = "1HXhUFfhgV74HkQezNr1hcI8kU2lLatLY-eduzoabOIs"; // Sheet name "PCRopsis Products"
var pcrOpsisInventory1 = "13Q-M1RuKHp-ByF0yXnE0Wjsih-JuH7lD3zBzQqQ8nmk"; // Sheet name "PCRopsis Inventory"
var plantOpsisInventory1 = ""; // Sheet name "Plantopsis Inventory"

var source = SpreadsheetApp.openById(mainsource1).getSheetByName("PR_1275.0.2 PCROpsis Aliquot Re");
var pcrOpsisSheet = SpreadsheetApp.openById(pcrOpsisInventory1).getSheetByName("Inventory");
//var plantOpsisSheet = SpreadsheetApp.openById(plantOpsisInventory1).getSheetByName("Inventory");


var numrow = source.getLastRow() - 1;
console.log("numrow" + numrow);

var passfailcolumn = source.getRange("M2:M").getValues();
console.log(passfailcolumn);
//console.log(passfailcolumn[0,0]);

var sourcelist1 = [8, 1, 2, 3, 4, 5, 6, 7, 9, 10, 11, 12, 5];
//console.log(sourcelist1);
var destinationlist1 = [1, 2, 3, 4, 5, 6, 9, 10, 11, 12, 13, 14, 7];
//console.log(destinationlist1);


var inventorychangeDate = new Date(Date.now()).toLocaleString().split(',')[0];

var productline = source.getRange("A2:A").getValues();
console.log("productline" +productline);


function transfer() {

  for (i = 0; i < numrow; i++) {

    // trouble shooting purpose 
    var x = source.getRange(i + 2, 1).getValues();
    var y = source.getRange(i + 2, 3).getValues();


    //var productline = source.getRange(i+2,1).getValue();
    //console.log("productline" + productline);

    var deslastRow = pcrOpsisSheet.getLastRow() + 1;
    console.log("deslastRow" +deslastRow);

    var deslastRow2 = plantOpsisSheet.getLastRow() + 1;
    console.log("destlastrow" +deslastRow2);



    if (passfailcolumn[0, i] == "Pending") {

      //console.log( x + " lot " + y + " is pending QC." );
      


 
      if (productline[0, i] == "DirectPCRA" || "DirectPCRB" || "PCRAVerne") {

        

        // "DirectPCRA" |
        for (iii = 0; iii < sourcelist1.length; iii++) {






          var myRange1 = source.getRange(i + 2, sourcelist1[iii]).getValues();
          console.log("myRange1" + myRange1);

          plantOpsisSheet.getRange(deslastRow2, destinationlist1[iii]).setValues(myRange1);

          plantOpsisSheet.getRange(deslastRow2, 8).setValue(inventorychangeDate);

          source.getRange(i + 2, 13).setValue("Pass");

        }

      }

      else if (productline[0, i] == "RVD" || "SRVD" || "SRVDEN" || "RVDRT" || "BUCC" || "RVDE" || "RVDU" || "RVDB" || "BCS Nano" || "BCS Concentrator" || "PCRACT" || "PCRSUP") {


        var workingbarcode = source.getRange(i + 2, 8).getValue(); 
        var workingALIcount = source.getRange(i + 2, 5).getValue();

        //create barcode list from inventory sheet
        var destBarcode = pcrOpsisSheet.getRange("A2:A").getValues();
        var destArray = [];
        for (d = 0; d < destBarcode.length; d++) {
          destArray.push(destBarcode[d][0]);
        }

        var idx = destArray.indexOf(workingbarcode);
        console.log("IDX: " +idx)
        




        if (idx > 0){ //Present in record
        var currentsourceRemainQT = pcrOpsisSheet.getRange(idx + 2, 7).getValue();
        var currentTotalRemainQT = pcrOpsisSheet.getRange(idx + 2, 6).getValue();

        var newdestRemainQT = currentsourceRemainQT + workingALIcount;
        console.log("currentsourceRemainQT" + currentsourceRemainQT )
        var newdestTotalQT = currentTotalRemainQT + workingALIcount;

        var newsourceRemainQT = pcrOpsisSheet.getRange(idx + 2, 7).setValue(newdestRemainQT);
        var newTotalRemainQT = pcrOpsisSheet.getRange(idx + 2, 6).setValue(newdestTotalQT);
        var newTotalRemainQT = pcrOpsisSheet.getRange(idx + 2, 8).setValue(inventorychangeDate);


        }



        else if (idx < 0){ //Not in record

        for (ii = 0; ii < sourcelist1.length; ii++) { //if not in record, append new record to inventory sheet.

          var myRange = source.getRange(i + 2, sourcelist1[ii]).getValues();
          console.log(myRange);

          pcrOpsisSheet.getRange(deslastRow, destinationlist1[ii]).setValues(myRange);

        }

          
        }


        source.getRange(i + 2, 13).setValue("Pass"); // finally made the row PASS.


        pcrOpsisSheet.getRange(deslastRow, 8).setValue(inventorychangeDate);  // setting last date change value

      }
      else {

      }

    }

    else if (passfailcolumn[0, i] == "Fail") {


      console.log(x + " lot " + y + " has failed QC.");

    }
    else {
      console.log("The " + x + " lot " + y + " lot has passed or cell is empty.");
    }


  }


}
