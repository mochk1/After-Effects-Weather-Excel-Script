#include "./deps/shim.js";
#include "./deps/xlsx.flow.js";

// UI
var mainWindow = new Window("palette", "Easy Weather", undefined);
mainWindow.orientation = "column";
mainWindow.preferredSize = [100,100];
var groupOne = mainWindow.add("group", undefined, "groupOne");
groupOne.orientation = "column";
var ProjBtn = groupOne.add("Button", undefined, "Open Template");
var getFileButton = groupOne.add("Button", undefined, "Select File...");
var img = groupOne.add("image", undefined, "./deps/Excel.png");
img.preferredSize = [100,100];
getFileButton.helpTip = "Select XLSX File";

mainWindow.center();
mainWindow.show();


//open template


 ProjBtn.onClick = function() {
                   var FileLocation = "./TEMP + TASHKIF.aep"
                var my_file = new File(FileLocation);   
             if (my_file.exists) {
                 app.beginSuppressDialogs()
                new_project = app.open(my_file);
                app.endSuppressDialogs()}
                else {
                        alert ("No File Found");
                       return false;}}

//function- select and read file

getFileButton.onClick = function() {
    
    
    
    try{
var file = new File;
var filename = file.openDlg ("Select A File");
var infile = File(filename);
infile.open("r");
infile.encoding = "binary";
var data = infile.read();
var workbook = XLSX.read(data, {type:"binary"});
var first_sheet_name = workbook.SheetNames[0];
var first_worksheet = workbook.Sheets[first_sheet_name];
var data = XLSX.utils.sheet_to_json(first_worksheet, {header:1});

//TASHKIF

var tashkif = app.project.item(2);


//set tomorrow
var dlayer = app.project.item(2).layer("DAY CTRL").property("Effects");
for(i=1;i<=dlayer.numProperties;i++){
dlayer.property(i).property("Checkbox").setValue(false);    
    }

 var d = new Date(Date(0));
var weekday=new Array(7);
weekday[0]="ראשון";
weekday[1]="שני";
weekday[2]="שלישי";
weekday[3]="רביעי";
weekday[4]="חמישי";
weekday[5]="שישי";
weekday[6]="שבת";
var tom=(d.getDay()+1);
if(tom==7) {tom=0}


tashkif.layer("DAY CTRL").property("Effects").property(weekday[tom]).property("Checkbox").setValue(true);


app.beginUndoGroup("WEATHER");


//set text, check for empty lines

try {
    var B43 = workbook.Sheets.גיליון1.B43.v;
    tashkif.layer("text 01").property("Text").property("Source Text").setValue(B43);}
catch(e){
   tashkif.layer("text 01").property("Text").property("Source Text").setValue('');
    }
try {
    var D43 = workbook.Sheets.גיליון1.D43.v;
    tashkif.layer("text 02").property("Text").property("Source Text").setValue(D43);}
catch(e){
    tashkif.layer("text 02").property("Text").property("Source Text").setValue('');
    }
try {
    var F43 = workbook.Sheets.גיליון1.F43.v;
    tashkif.layer("text 03").property("Text").property("Source Text").setValue(F43);}
catch(e){
    tashkif.layer("text 03").property("Text").property("Source Text").setValue('');
    }
try {
    var H43 = workbook.Sheets.גיליון1.H43.v;
    tashkif.layer("text 04").property("Text").property("Source Text").setValue(H43);}
catch(e){
    tashkif.layer("text 04").property("Text").property("Source Text").setValue('');
    }
try {
    var J43 = workbook.Sheets.גיליון1.J43.v;
    tashkif.layer("text 05").property("Text").property("Source Text").setValue(J43);}
catch(e){
    tashkif.layer("text 05").property("Text").property("Source Text").setValue('');
    }

//temps

var B45 =workbook.Sheets.גיליון1.B45.v
var D45 = workbook.Sheets.גיליון1.D45.v
var F45 = workbook.Sheets.גיליון1.F45.v
var H45 = workbook.Sheets.גיליון1.H45.v
var J45 = workbook.Sheets.גיליון1.J45.v
var L40 = workbook.Sheets.גיליון1.L40.v
var L45 = workbook.Sheets.גיליון1.L45.v

//set temps


tashkif.layer("TEMP1").property("Text").property("Source Text").setValue(B45);
tashkif.layer("TEMP2").property("Text").property("Source Text").setValue(D45);
tashkif.layer("TEMP3").property("Text").property("Source Text").setValue(F45);
tashkif.layer("TEMP4").property("Text").property("Source Text").setValue(H45);
tashkif.layer("TEMP5").property("Text").property("Source Text").setValue(J45);
tashkif.layer("TEMP8").property("Text").property("Source Text").setValue(L40);
tashkif.layer("TEMP7").property("Text").property("Source Text").setValue(L45);


//reset icons


for(i=1;i<=app.project.item(2).numLayers;i++){
    var ilayer = app.project.item(2).layer(i);
    if(ilayer.name.indexOf ('icon')!=-1){
            for(j=1;j<=ilayer.property("Effects").numProperties;j++){
            ilayer.property("Effects").property(j).property("Checkbox").setValue(false);
 }}}
    
for(i=1;i<=app.project.item(4).numLayers;i++){
    var ilayer = app.project.item(4).layer(i);
    if(ilayer.name.indexOf ('icon')!=-1){
            for(j=1;j<=ilayer.property("Effects").numProperties;j++){
            ilayer.property("Effects").property(j).property("Checkbox").setValue(false);
 }}}
    

//set icons


var L38 = workbook.Sheets.גיליון1. L38.v
var L43 = workbook.Sheets.גיליון1.L43.v


//01
try{
    var B32 = workbook.Sheets.גיליון1.B32.v
    var B34 = workbook.Sheets.גיליון1.B34.v
tashkif.layer("icon01").property("Effects").property(B34).property("Checkbox").setValue(true)
}

catch(e){ 
    alert("נא לשנות ידנית את האייקון של יום " + B32)
    }


//02
try{
    var D34 = workbook.Sheets.גיליון1.D34.v
    var D32 = workbook.Sheets.גיליון1.D32.v
tashkif.layer("icon02").property("Effects").property(D34).property("Checkbox").setValue(true)
}

catch(e){ 
     alert(" נא לשנות ידנית את האייקון של יום " + D32)
    }



//03
try{
    var F32 = workbook.Sheets.גיליון1.F32.v
    var F34 = workbook.Sheets.גיליון1.F34.v
tashkif.layer("icon03").property("Effects").property(F34).property("Checkbox").setValue(true)
}

catch(e){ 
     alert(" נא לשנות ידנית את האייקון של יום " + F32)
    }

//04
try{
    var H32 = workbook.Sheets.גיליון1.H32.v
    var H34 = workbook.Sheets.גיליון1.H34.v
tashkif.layer("icon04").property("Effects").property(H34).property("Checkbox").setValue(true)
}

catch(e){ 
     alert(" נא לשנות ידנית את האייקון של יום " + H32)
    }

//05
try{
    var J32 = workbook.Sheets.גיליון1.J32.v
    var J34 = workbook.Sheets.גיליון1.J34.v
tashkif.layer("icon05").property("Effects").property(J34).property("Checkbox").setValue(true)
}

catch(e){ 
     alert(" נא לשנות ידנית את האייקון של יום " + J32)
    }

//06
try{
    var J32 = workbook.Sheets.גיליון1.L38.v
    var J34 = workbook.Sheets.גיליון1.L36.v
tashkif.layer("icon06").property("Effects").property(L38).property("Checkbox").setValue(true)
}

catch(e){ 
     alert(" נא לשנות ידנית את האייקון של יום " + L36)
    }


//07
try{
    var J32 = workbook.Sheets.גיליון1.L43.v
    var J34 = workbook.Sheets.גיליון1.L41.v
tashkif.layer("icon07").property("Effects").property(L43).property("Checkbox").setValue(true)
}

catch(e){ 
     alert(" נא לשנות ידנית את האייקון של יום " + J32)
    }



//graph


var B45 = parseInt(workbook.Sheets.גיליון1.B45.v);
var D45 = parseInt(workbook.Sheets.גיליון1.D45.v);
var F45 = parseInt(workbook.Sheets.גיליון1.F45.v);
var H45 = parseInt(workbook.Sheets.גיליון1.H45.v);
var J45 = parseInt(workbook.Sheets.גיליון1.J45.v);
var L40 = parseInt(workbook.Sheets.גיליון1.L40.v);

app.project.item(2).layer("LINE CTRL").property("Effects").property("max").property("Slider").setValue(Math.max(B45, D45, F45, H45, J45)-4);
app.project.item(2).layer("LINE CTRL").property("Effects").property("min").property("Slider").setValue(Math.min(B45, D45, F45, H45, J45));


tashkif.layer("LINE CTRL").property("Effects").property("P1").property("Slider").setValue(B45);
tashkif.layer("LINE CTRL").property("Effects").property("P2").property("Slider").setValue(D45);
tashkif.layer("LINE CTRL").property("Effects").property("P3").property("Slider").setValue(F45);
tashkif.layer("LINE CTRL").property("Effects").property("P4").property("Slider").setValue(H45);
tashkif.layer("LINE CTRL").property("Effects").property("P5").property("Slider").setValue(J45);
tashkif.layer("LINE CTRL").property("Effects").property("P6").property("Slider").setValue(L40);

//TEMPS

var temp = app.project.item(4);


//zafon
temp.layer("24").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.B3.v);
temp.layer("18").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.C3.v);
//kineret
temp.layer("23").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.B4.v);
temp.layer("17").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.C4.v);
//nazrat
temp.layer("22").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.B5.v);
temp.layer("16").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.C5.v);
//haifa
temp.layer("21").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.B6.v);
temp.layer("15").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.C6.v);
//ariel
temp.layer("20").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.B9.v);
temp.layer("14").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.C9.v);
//shfela
temp.layer("19").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.B8.v);
temp.layer("13").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.C8.v);
//ta
app.project.item(4).layer("12").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.B7.v);
app.project.item(4).layer("11").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.C7.v);
//jerus
temp.layer("10").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.B10.v);
temp.layer("09").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.C10.v);
//bs
temp.layer("08").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.B12.v);
temp.layer("07").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.C12.v);
//dead sea
temp.layer("06").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.B14.v);
temp.layer("05").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.C14.v);
//ramon
temp.layer("04").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.B13.v);
temp.layer("03").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.C13.v);
//eilat
temp.layer("02").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.B15.v);
temp.layer("01").property("Text").property("Source Text").setValue(workbook.Sheets.גיליון1.C15.v);



//set icons


temp.layer("icons  הרי הצפון").property("Effects").property(workbook.Sheets.גיליון1.D3.v).property("Checkbox").setValue(true)
temp.layer("icons  כינרת").property("Effects").property(workbook.Sheets.גיליון1.D4.v).property("Checkbox").setValue(true)
temp.layer("icons  נצרת").property("Effects").property(workbook.Sheets.גיליון1.D5.v).property("Checkbox").setValue(true)
temp.layer("icons  חיפה").property("Effects").property(workbook.Sheets.גיליון1.D6.v).property("Checkbox").setValue(true)
temp.layer("icons  אריאל").property("Effects").property(workbook.Sheets.גיליון1.D9.v).property("Checkbox").setValue(true)
temp.layer("icons השפלה").property("Effects").property(workbook.Sheets.גיליון1.D8.v).property("Checkbox").setValue(true)
temp.layer("icons תל-אביב").property("Effects").property(workbook.Sheets.גיליון1.D7.v).property("Checkbox").setValue(true)
temp.layer("icons ירושלים").property("Effects").property(workbook.Sheets.גיליון1.D10.v).property("Checkbox").setValue(true)
temp.layer("icons באר שבע").property("Effects").property(workbook.Sheets.גיליון1.D12.v).property("Checkbox").setValue(true)
temp.layer("icons ים המלח").property("Effects").property(workbook.Sheets.גיליון1.D14.v).property("Checkbox").setValue(true)
temp.layer("icons מצפה רמון").property("Effects").property(workbook.Sheets.גיליון1.D13.v).property("Checkbox").setValue(true)
temp.layer("icons אילת").property("Effects").property(workbook.Sheets.גיליון1.D15.v).property("Checkbox").setValue(true)


app.endUndoGroup();

alert('Done !');
mainWindow.close();
    }
catch(e){
alert('Something is wrong :( \nPlease make sure your file is in xls format!')    

    }
};



/*
    
    
    expression for nulls in Group1 comp


slider = comp("Tashkif").layer("LINE CTRL").effect("P1")("Slider")
min =  comp("Tashkif").layer("LINE CTRL").effect("min")("Slider")
max = comp("Tashkif").layer("LINE CTRL").effect("max")("Slider")

x=transform.position[0];
y=linear(slider, min, max,0,150);

[x,150-y]

   */
