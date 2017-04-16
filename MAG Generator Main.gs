/*
This code was written by Elder Derek Parsons of the New Jersey Morristown Mission (NJMM).
Its function is to create a report called "Mission At a Glance" (MAG). 
It was written for the use of NJMM. They may use it however they see fit.
*/


function getFileList(option) {
//When option = 0, returns all files in the MAG Generator folder.
//When option = 1, returns an array of files and exludes the MAG Generator spreadsheet itself.
  Logger.log("called function getFileList("+option+")");
  var currentFileId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var folders = DriveApp.getFoldersByName("MAG Generator");
  while (folders.hasNext()) {
    var folderId = 0;
    var folder = folders.next();
    var folderFiles = folder.getFiles();
    while (folderFiles.hasNext()) {
      var file = folderFiles.next();
      if (file.getName() == "MAG Generator") {
        folderId = folder.getId();
        break;
      }
    }
    if (folderId != 0) {
      break;
    }
  }
  var files = DriveApp.getFolderById(folderId).getFiles();
  if (option == 1) {
    Logger.log("finished function getFileList()");
    return files;
  }
  var fileList = [];
  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    if (fileName != "MAG Generator") {
      fileList.push(fileName);
    }
  }
  Logger.log("finished function getFileList()");
  return fileList;
}


function selectFilesDialogue() {
  var html = HtmlService.createTemplateFromFile('Select Files')
  .evaluate()
  .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  .setWidth(800)
  .setHeight(400);
  SpreadsheetApp.getUi()
  .showModalDialog(html, 'Select Files');
}


function finishedDialogue() {
  var html = HtmlService.createTemplateFromFile('Finished')
  .evaluate()
  .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  .setWidth(400)
  .setHeight(200);
  SpreadsheetApp.getUi()
  .showModalDialog(html, 'Finished');
}


function grabFiles(formData) {
  Logger.log("called function grabFiles("+formData+")");
  var roster;
  var vehicles;
  var files = getFileList(1);
  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    if (fileName == formData.organizationRoster) {
      roster = file;
    } else if (fileName == formData.vehicleAssignments) {
      vehicles = file;
    }
  }
  var requiredFiles = {rosterFile: roster, vehicleFile: vehicles};
  Logger.log("finished function grabFiles()");
  return requiredFiles;
}


function sortSheet(sheet,column) {
  Logger.log("called function sortSheet()");
  var dataRange = sheet.getDataRange();
  var height = dataRange.getHeight();
  var width = dataRange.getWidth();
  var range = sheet.getRange(3,1,height-2,width);
  range.sort(column);
  return sheet;
}


function createNewMag() { 
//Creates the file, but doesn't write anything.
  Logger.log("called function createNewMag()");
  var date = Utilities.formatDate(new Date(), "EST5EDT", "MMM dd, yyyy");
  var newMagSpreadsheet = SpreadsheetApp.create("MAG " + date);
  var url = newMagSpreadsheet.getUrl();
  PropertiesService.getScriptProperties().setProperty('url', url);
  var newMagFileId = newMagSpreadsheet.getId();
  var folderId = 0;
  var folders = DriveApp.getFolders();
  while (folders.hasNext()) {
    var folder = folders.next();
    if (folder.getName() == "Generated MAGs") {
      folderId = folder.getId();
      break;
    }
  }
  var file = DriveApp.getFileById(newMagFileId);
  var parentFolders = file.getParents();
  var sourceFolder = parentFolders.next();
  var destinationFolder = DriveApp.getFolderById(folderId);
  destinationFolder.addFile(file);
  sourceFolder.removeFile(file);
  Logger.log("finisehd function createNewMag()");
  return newMagSpreadsheet;
}


function getData(roster,vehicle,column1,column2,restriction) {
//This is a (proabably too complicated) function that returns a bunch of strings from the IMOS source reports as needed to create the data objects.
  Logger.log("called function getData()");
  var names = [];
  var height = roster.getDataRange().getHeight();
  var values = [];
  var column1Values = roster.getRange(3,column1,height-2,1).getValues();
  if (column2 !== undefined && restriction !== undefined) {
    var column2Values = roster.getRange(3,column2,height-2,1).getValues();
    for (var x in column1Values) {
      if (column2Values[x][0] == restriction) {
        values.push(column1Values[x][0]);
      }
    }
  } else {
    for (x in column1Values) {
      values.push(column1Values[x][0]);
    }
  }
  for (var x in values) {
    var value = values[x];
    if (value == '') {
      continue;
    }
    if (names.indexOf(value) == -1) {
      if (names.length > 0) {
        var number = parseInt(value.slice(0,1));
        var length = names.length;
        for (k=0; k<length; k++) {
          var nextNumber = parseInt(names[k].slice(0,1));
          if (number < nextNumber) {
            names.splice(k,0,value);
            break;
          } else if (number < 10 && k==length-1) {
            names.push(value);
            break;
          }
        }
        if (names.indexOf(value) == -1) {
          names.unshift(value);
        }
      } else {
        names.push(value);
      }
    }
  }
  return names;
}


function getZones(roster,vehicle) {
  Logger.log("called function getZones()");
  var zones = [];
  roster = sortSheet(roster,18);
  var titles = getData(roster,vehicle,18);
  for (i=0; i<titles.length; i++) {
    zones[i] = new Zone(titles[i],roster,vehicle);
  }
  Logger.log("finished function getZones()");
  return zones;
}


function Zone(name,roster,vehicle) {
  this.name = name;
  this.districts = getDistricts(name,roster,vehicle);
}


function getDistricts(zone,roster,vehicle) {
  Logger.log("called function getDistricts()");
  var districts = [];
  roster = sortSheet(roster,19);
  var titles = getData(roster,vehicle,19,18,zone);
  for (j=0; j<titles.length; j++) {
    districts[j] = new District(titles[j],roster,vehicle);
  }
  return districts;
}


function District(name,roster,vehicle) {
  this.name = name;
  this.areas = getAreas(name,roster,vehicle);
  for (var area in this.areas) {
    for (var missionary in this.areas[area].missionaries) {
      if (this.areas[area].missionaries[missionary].position == 'DL' || this.areas[area].missionaries[missionary].position == 'DT') {
        this.areas.splice(0, 0, this.areas.splice(area, 1)[0]);
      }
    }
  }
}


function getAreas(district,roster,vehicle) {
  Logger.log("called function getAreas()");
  var areas = [];
  roster = sortSheet(roster,20);//Changed from 19 to 20 by E. Clement
  var titles = getData(roster,vehicle,20,19,district);//Changed from 19,18 to 20,19 by E. Clement
  for (h=0; h<titles.length; h++) {
    areas[h] = new Area(titles[h],roster,vehicle);
  }
  return areas;
}


function Area(name,roster,vehicle) {
  this.name = name;
  var rosterRows = [];
  var rosterHeight = roster.getDataRange().getHeight();
  var rosterValues = roster.getRange(3,20,rosterHeight-2,1).getValues();//Changed from 19 to 20 by E. Clement
  for (var cell in rosterValues) {
    if (rosterValues[cell][0] == name) {
      rosterRows.push(parseInt(cell)+3);
    }
  }
  var vehicleRow;
  var vehicleHeight = vehicle.getDataRange().getHeight();
  var vehicleValues = vehicle.getRange(3,3,vehicleHeight-2,1).getValues();
  for (var cell in vehicleValues) {
    if (vehicleValues[cell][0] == name) {
        vehicleRow = parseInt(cell)+3;
        break;
    }
  }
  this.address = roster.getRange(rosterRows[0],32).getValue();//Changed from 14 to 32 by E. Clement
  this.phone = roster.getRange(rosterRows[0],13).getValue();//Changed from 9 to 13 by E. Clement
  if (vehicleRow !== undefined) {
    var rowValues = vehicle.getRange(vehicleRow,1,1,9).getValues();
    this.vehicle = rowValues[0][3] + "-" + rowValues[0][4] + "-" + rowValues[0][8];
  } else {
    this.vehicle = "No vehicle assigned";
  }
  var fullUnit = roster.getRange(rosterRows[0],17).getValue();//Changed from 16 to 17 by E. Clement
  var index = fullUnit.indexOf("(");
  this.unit = fullUnit.slice(0,index);
  this.missionaries = getMissionaries(roster,rosterRows);
  var positionValues = {
    JC:0,
    SC:5,
    DL:1,
    TR:4,
    STL1:2,
    ZL1:2,
    STL2:1,
    ZL2:1,
    DT:4,
    STLT:4,
    AP:3
  }
  for (var missionary1 in this.missionaries) {
    for (var missionary2 in this.missionaries) {
      var p1 = this.missionaries[missionary1].position;
      var p2 = this.missionaries[missionary2].position;
      if (positionValues[p1] > positionValues[p2]) {
        if (missionary1 < missionary2) {
        this.missionaries.splice(missionary2, 0, this.missionaries.splice(missionary1, 1)[0]);
        }
      } else if (positionValues[p1] < positionValues[p2]) {
        if (missionary1 > missionary2) {
          this.missionaries.splice(missionary1, 0, this.missionaries.splice(missionary2, 1)[0]);
        }
      }
    }
  }
}


function getMissionaries(roster,rows) {
  Logger.log("called function getMissionaries()");
  var missionaries = [];
  var fullNames = roster.getRange(rows[0],1,rows.length,1).getValues();
  var lastNames = [[],[]];
  for (var fullName in fullNames) {
    var index = fullNames[fullName][0].indexOf(",");
    if (index !== -1) {
      lastNames[0].push(fullNames[fullName][0].slice(0,index));
      lastNames[1].push(parseInt(rows[fullName]));
    }
  }
  for (var person in lastNames[0]) {
    missionaries[person] = new Missionary(lastNames[0][person],roster,lastNames[1][person]);
  }
  return missionaries;
}


function Missionary(name,roster,row) {
  this.name = name;
  this.position = roster.getRange(row,12).getValue();//Changed from 8 to 12 by E. Clement
}


function generateMag(formData) { 
//This is the main function that gathers the data, organizes it, and writes it to a new spreadsheet.
  Logger.log("called function generateMag()");
  SpreadsheetApp.getActiveSpreadsheet().toast('Task started', "This may take a minute or two. Please be patient and don't close this page. You will be notified when the task is completed.");
  var files = grabFiles(formData);
  var roster = SpreadsheetApp.open(files.rosterFile).getSheets()[0];
  /*var rosterHeight = roster.getDataRange().getHeight();
  var rosterRange = roster.getRange(3,17,rosterHeight-2,1).getValues();*/
  var vehicle = SpreadsheetApp.open(files.vehicleFile).getSheets()[0];
  var sortedVehicle = sortSheet(vehicle,3);
  var zones = getZones(roster,vehicle);
  Logger.log(zones);
  var mag = createNewMag();
  writeMag(mag,zones);
  Logger.log("finished function generateMag()");
  finishedDialogue();
}


function writeMag(mag,zones) {
  Logger.log("called function writeMag()");
  var sheet = mag.getActiveSheet();
  var row = 2;
  var column = 1;
  var depth = 0;
  var colors = {
    JC:"white",
    SC:"white",
    DL:"#ff9999",
    TR:"#84e184",
    STL1:"#ffb84d",
    STL2:"#ffb84d",
    ZL1:"#e699ff",
    ZL2:"#e699ff",
    DT:"#d2a679",
    STLT:"#d2a679",
    AP:"#b3ffff",
    districtTitle:"#c0c0c0",
    areaTitle:"#e6e6e6"
  }
  for (var zone in zones) {
    if (column > 10) {
      column-=8;
      row+=59;
      depth+=1;
    }
    sheet.getRange(row,column,1,2).merge().setHorizontalAlignment("center").setFontSize(12).setValue(zones[zone].name);
    sheet.setColumnWidth(column,230);
    sheet.setColumnWidth(column+1,120);
    row+=1;
    var districts = zones[zone].districts;
    var jumped = false;
    for (var district in districts) {
      if (zone == 0 && jumped == false) { //since zone 0 can be much longer than any other zone is allowed to be, this block will push part of the zone unto the next page if needed.
        if (row+(districts[district].areas.length*4) > 58) {
          row = 61;
          sheet.getRange(row,column,1,2).merge().setHorizontalAlignment("center").setFontSize(12).setValue(zones[zone].name + " (continued)");
          row = 62;
          jumped = true;
        }
      }
      sheet.getRange(row,column,1,2).merge().setHorizontalAlignment("center").setBackground(colors.districtTitle).setBorder(true,true,true,true,false,false).setValue(districts[district].name);
      row+=1;
      var areas = districts[district].areas
      for (r=0; r<areas.length; r++) {
        sheet.getRange(row,column).setBackground(colors.areaTitle).setValue(areas[r].name);
        sheet.getRange(row,column,4,2).setBorder(true,true,true,true,false,false);
        areas[r].address = areas[r].address.replace(/\r?\n|\r/g," ");
        sheet.getRange(row+1,column).setValue(areas[r].address);
        sheet.getRange(row+2,column).setValue(areas[r].phone);
        sheet.getRange(row+3,column).setValue(areas[r].vehicle);
        sheet.getRange(row,column+1).setValue(areas[r].unit);
        var missionaryRow = row+3;
        var missionaries = areas[r].missionaries;
        for (v=0; v<missionaries.length; v++) {
          sheet.getRange(missionaryRow,column+1).setBackground(colors[missionaries[v].position]).setValue(missionaries[v].name);
          missionaryRow-=1;
        }
        row+=4;
      }
    }
    if (zone == 0) { //prints the color key under zone 0
      row+=3;
      sheet.getRange(row,column).setHorizontalAlignment("center").setValue("Color Key");
      row+=1;
      for (var color in colors) {
        sheet.getRange(row,column).setBackground(colors[color]).setValue(color);
        row+=1;
      }
    }
    column+=2;
    row=2+(59*depth);
  }
  sheet.getDataRange().setVerticalAlignment("center").setFontWeight("bold"); //makes the sheet easier to read.
  sheet.getRange(1,1,1,10).merge().setFontSize(18).setHorizontalAlignment("center").setVerticalAlignment("top").setFontSize(18).setValue(mag.getName()); //sets MAG title.
  Logger.log("finished function writeMag()");
}