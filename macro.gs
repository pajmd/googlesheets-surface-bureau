/** @OnlyCurrentDoc */

function Insert_LineX() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('1:1').activate();
  spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveRange().getLastRow(), 1);
  spreadsheet.getActiveRange().offset(spreadsheet.getActiveRange().getNumRows(), 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
};

function bureau_individuel(persons) {
  return persons * 11;
}

function bureau_area(persons) {
  return (15+ (persons-2)*5);
}

function open_space(persons) {
  return persons * 6;
}

function salle_reunion(persons) {
  return persons * 2;
}

function surface_totale() {
  return addup("stot")
}

function surface_totale_plus() {
  return addup("stot") * 1.1
}

function surface_totale_moins() {
  return addup("stot") * 0.9
}

function espace_commun() {
  return addup("socc") * 0.2
  // var ss = SpreadsheetApp.getActiveSpreadsheet();
  // var spreadsheet = ss.getSheetByName("Surface");
  // var nm_ranges = spreadsheet.getNamedRanges();
  // Logger.log("Range names: " + nm_ranges);
  // var total = 0;
  // for (var i = 0; i < nm_ranges.length; i++) {
  //   Logger.log("nm_rg: " + nm_ranges[i]);
  //   if (nm_ranges[i].getName().includes("socc")) {
  //     var vals = nm_ranges[i].getRange().getValues();
  //     for (var j=0; j< vals.length; j++) {
  //       if (typeof vals[j][0] === 'number') {
  //         total += vals[j][0];
  //       }
  //     }
  //     break;
  //   }
  // }
  // Logger.log("total: " + total);
  // return total * 0.2;
}

function addup(what) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheet = ss.getSheetByName("Surface");
  var nm_ranges = spreadsheet.getNamedRanges();
  Logger.log("Range names: " + nm_ranges);
  var total = 0;
  for (var i = 0; i < nm_ranges.length; i++) {
    Logger.log("nm_rg: " + nm_ranges[i]);
    if (nm_ranges[i].getName().includes(what)) {
      var vals = nm_ranges[i].getRange().getValues();
      for (var j=0; j< vals.length; j++) {
        if (typeof vals[j][0] === 'number') {
          total += vals[j][0];
        }
      }
      break;
    }
  }
  Logger.log("total: " + total);
  return total;
}

function add_salle() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheet = ss.getSheetByName("Surface");
  spreadsheet.activate();
  var common_space = spreadsheet.createTextFinder('Espace commun').findNext();
  common_space.activate();  // Activate range
  spreadsheet.insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  var added_room = common_space.offset(-1,0,1,1)
  added_room.activate();
  added_room.offset(0,0,1,1).setValue(0)
  added_room.offset(0,0,1,1).setHorizontalAlignment("center");
  added_room.offset(0,1,1,1).setValue("=salle_reunion(C"+added_room.getRowIndex()+")")
  added_room.offset(0,1,1,1).setHorizontalAlignment("center");
  added_room.setBorder(true,true,true,true, true, true, "#094040", SpreadsheetApp.BorderStyle.SOLID_THICK);
  added_room.offset(0,1,1,1).setBorder(true,null,true,true, true, true, "#618c67", SpreadsheetApp.BorderStyle.SOLID_THICK);
}

function delete_salle() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheet = ss.getSheetByName("Surface");
  spreadsheet.activate();
  var open_space = spreadsheet.createTextFinder('Espace commun').findNext();
  open_space.activate();  // Activate range
  var row_to_del = open_space.offset(-2,0,1,1)
  var prev_row = row_to_del.offset(-1,0,1,1).getValue();
  if (!prev_row.toString().includes("Salle")) {
    spreadsheet.deleteRow(row_to_del.getRow())
  }
}

function delete_bureau() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheet = ss.getSheetByName("Surface");
  spreadsheet.activate();
  var open_space = spreadsheet.createTextFinder('Open space').findNext();
  // spreadsheet.getRange('8:8').activate();
  open_space.activate();  // Activate range
  var row_to_del = open_space.offset(-2,0,1,1)
  var prev_row = row_to_del.offset(-1,0,1,1).getValue();
  if (!prev_row.toString().includes("Bureau")) {
    spreadsheet.deleteRow(row_to_del.getRow())
  }
}

function add_bureau_partage() {
  // var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheet = ss.getSheetByName("Surface");
  spreadsheet.activate();
  var open_space = spreadsheet.createTextFinder('Open space').findNext();
  // spreadsheet.getRange('8:8').activate();
  open_space.activate();  // Activate range
  // spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  spreadsheet.insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  //spreadsheet.insertRows(spreadsheet.getActiveRange().getRow() -1, 1);
  //spreadsheet.insertRowsAfter(open_space.offset(-3,0).getRow(), 1);
  var added_bureau = open_space.offset(-1,0,1,1)
  added_bureau.activate();
  added_bureau.offset(0,0,1,1).setValue(1)
  added_bureau.offset(0,0,1,1).setHorizontalAlignment("center");
  added_bureau.offset(0,1,1,1).setValue("=bureau_area(C"+added_bureau.getRowIndex()+")")
  added_bureau.offset(0,1,1,1).setHorizontalAlignment("center");
  added_bureau.setBorder(true,true,true,true, true, true, "#094040", SpreadsheetApp.BorderStyle.SOLID_THICK);
  added_bureau.offset(0,1,1,1).setBorder(true,null,true,true, true, true, "#618c67", SpreadsheetApp.BorderStyle.SOLID_THICK);

  // var minus_img = insertCellImage('https://drive.google.com/uc?export=download&id=1MtHcPK-0d-KBXnMFwb1L8UwvK550B1QK')
  // added_bureau.offset(0,2,1,1).setValue(minus_img);

  // The following: Adding image over grid using blob works
  // var img_names = DriveApp.getFilesByName("small-minus-45-19.png")
  // const img_id = img_names.next().getId();
  // Logger.log("Image id=" + DriveApp.getFileById(img_id).getName());
  // over_img = spreadsheet.insertImage(
  //   DriveApp.getFileById(img_id).getBlob(), 
  //   added_bureau.offset(0,2,1,1).getColumn(), 
  //   added_bureau.offset(0,2,1,1).getRow()
  // );
  // var img_title = "rm_c"+added_bureau.offset(0,2,1,1).getColumn()+"r"+added_bureau.offset(0,2,1,1).getRow();
  // Logger.log("Image title: " + img_title);
  // over_img.setAltTextTitle(img_title);
  // over_img.assignScript("remove_me");

};

// Just not possible to col, row of overGridImage button
function remove_me(rm_but) {
  var ss = SpreadsheetApp.getActive();
  var row_to_del = ss.getActiveCell().getRow();
  var col = ss.getActiveCell().getColumn();
  var img_title = "rm_c"+col+"r"+row_to_del;
  Logger.log("remove img: " + img_title);
  ss.deleteRow(row_to_del)
  imgs = ss.getImages();
  for (const img of imgs) {
    var title = img.getAltTextTitle();
    if (title == img_title) {
      img.remove();
      break;
    }
  }

}

function insertCellImage(imageUrl, altTitle = "", altDescription = "") {
  //Doesn't help much bc we can;t assign a nacro to images in cells
 let image = SpreadsheetApp
                 .newCellImage()
                 .setSourceUrl(imageUrl)
                 .setAltTextTitle(altTitle)
                 .setAltTextDescription(altDescription)
                 .build();
  return image;
}


