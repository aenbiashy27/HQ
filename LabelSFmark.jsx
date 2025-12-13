call();


 // 13/12/2025 
function call(){ 


var appleScript2 = new File('/Applications/HQMaster/Data.app');

if (appleScript2.exists) {

    appleScript2.execute();
    }else{
        alert ( 'error code : app mount error DD');
          app.documents.close(SaveOptions.DONOTSAVECHANGES); 
      return ;

    };
};


main2();  
function main2(){  
var f = new File('/Applications/Adobe Photoshop CS6/Legal/en_US/Statistics.jsx');  
if (f.exists){


}else{
    alert ('error code : sq script error main222')
    app.documents.close(SaveOptions.DONOTSAVECHANGES); 
   return; 
    }};


var doc = app.activeDocument;  

zoom();


doc.layers.getByName('pic').move( activeDocument, ElementPlacement.PLACEATEND );  
doc.layers.getByName('Cut Layer').move( activeDocument, ElementPlacement.PLACEATBEGINNING );  
doc.layers.getByName('pic').hasSelectedArtwork = true; 



var arthne;
var artwne;


var selected = doc.selection;  

selectedHeight1 = selection[0].height;
selectedHeight1 = new UnitValue(selectedHeight1, 'pt').as('mm');
selectedHeight1 = Math.round(selectedHeight1*10)/10;

selectedWidth2 = selection[0].width;
selectedWidth2 = new UnitValue(selectedWidth2, 'pt').as('mm');
selectedWidth2 = Math.round(selectedWidth2*10)/10;


var selectedHeightcvlb2 = selectedHeight1-2;
var selectedWidthcvlb2 = selectedWidth2-2;


var arth = selectedHeight1;

var artw = selectedWidth2;


var rotateAmount1 ;
var rotateAmount2 ;


var selectedHeight ;
var selectedWidth ;

var selectedHeightcv ;
var selectedHeightcvlb ;

var selectedWidthcv ;
var selectedWidthcvlb ;


var selectedHeight11 ;

var selectedWidth22 ;


var OriginX ;
var OriginY ;

var selectedHeightlb ;
var selectedWidthlb ;

           
var WShape;
var cutline1;

var margin ;
var margin2 ;

var repeatAmountw;
var repeatAmounth;

var repeatAmountwl ;
var repeatAmounthl ;  

var LineW ;
var LineH ;

var sLineW;
var LLineW ;

var sLineH;
var LLineH ;


var RCDSize1;


var activeAB = doc.artboards[doc.artboards.getActiveArtboardIndex()];
var artboardRight = activeAB.artboardRect[2];
var artboardBottom = activeAB.artboardRect[3];
var myPageItem = doc.selection[0];
myPageItem.position ;

var RCD = ((selectedHeight1*2)+(selectedWidth2*2))/35;
var RCDCC = RCD;
RCDCC = new UnitValue(RCDCC, 'pt').as('mm');
RCDCC = Math.round(RCDCC*10)/10;

var tt;

var i1 = File ("/Applications/HQMaster/Sticker_Label_Auto.app/Contents/1.png");
var i2 = File ("/Applications/HQMaster/Sticker_Label_Auto.app/Contents/2.png");
var i3 = File ("/Applications/HQMaster/Sticker_Label_Auto.app/Contents/3.png");
var i4 = File ("/Applications/HQMaster/Sticker_Label_Auto.app/Contents/4.png");
var i5 = File ("/Applications/HQMaster/Sticker_Label_Auto.app/Contents/5.png");
var i6 = File ("/Applications/HQMaster/Sticker_Label_Auto.app/Contents/6.png");
var i7 = File ("/Applications/HQMaster/Sticker_Label_Auto.app/Contents/7.png");
var i8 = File ("/Applications/HQMaster/Sticker_Label_Auto.app/Contents/8.png");
var i9 = File ("/Applications/HQMaster/Sticker_Label_Auto.app/Contents/9.png");
var i10 = File ("/Applications/HQMaster/Sticker_Label_Auto.app/Contents/10.png");




myDlg = new Window('dialog', 'Cut Auto V13.1 - 23/02/23');  
  myDlg.alignChildren = 'center';

   

  shapePanel = myDlg.add('panel', undefined, '317 Sticker Kiss Cut');
  shapePanel.orientation = 'row';
  shapePanel.alignment = 'center';
  shapePanel.margins = [20, 20, 20, 10];
  shapeGroup = shapePanel.add('group');
  shapeGroup.orientation = 'row';
  shapeGroup.alignChildren = 'center';
  shapeGroup.graphics.font = 'Myriad Pro:14';  
  


var t1 = shapeGroup.add ("iconbutton", undefined, i1, {style: "toolbutton", toggle: true}); // c
t1.onClick= function (){
if (t1.value == true){
    
    WShape = 'Round';
arthne = arth - 2;
artwne = artw - 2;
cuttingsize.text = 'Cut Size:  '+artwne+'  x  '+arthne+' mm ' ;
Sizedetails.text = '2mm Bleed Circle';
    
          t10.value = false; 
        t2.value = false; 
     t3.value = false; 
      t4.value = false; 
       t5.value = false; 
         t6.value = false; 
        t7.value = false; 
         t8.value = false; 
           t9.value = false; 
            Savesaga.value = false ;
            Savesaga.show();
           fmarkcut.value = true;
           fmarkcut.show();
           rcGroup.hide();
           
           
           }
  
   };

var t2 = shapeGroup.add ("iconbutton", undefined, i2, {style: "toolbutton", toggle: true}); // d
t2.value = false;
t2.onClick= function (){
if (t2.value == true){
    
    WShape = 'RoundN';
    arthne = arth ;
artwne = artw ;
cuttingsize.text = 'Cut Size:  '+artwne+'  x  '+arthne+' mm ' ;
    Sizedetails.text = 'No Bleed Circle';
    
  t1.value = false; 
    t1.value = false; 
        t10.value = false; 
     t3.value = false; 
      t4.value = false; 
       t5.value = false; 
         t6.value = false; 
        t7.value = false; 
         t8.value = false; 
           t9.value = false; 
            Savesaga.value = false ;
            Savesaga.show();
           fmarkcut.value = true;
            fmarkcut.show();
           rcGroup.hide();
         }
   
   };

var t3 = shapeGroup.add ("iconbutton", undefined,i3, {style: "toolbutton", toggle: true}); // c
t3.onClick= function (){
if (t3.value == true){
    
    WShape = 'Rectangle';
arthne = arth - 2;
artwne = artw - 2;
cuttingsize.text = 'Cut Size:  '+artwne+'  x  '+arthne+' mm ' ;
    Sizedetails.text = '2mm Bleed Rectangle';
    
    
        t1.value = false; 
        t2.value = false; 
     t10.value = false; 
      t4.value = false; 
       t5.value = false; 
         t6.value = false; 
        t7.value = false; 
         t8.value = false; 
           t9.value = false; 
            Savesaga.value = false ;
            Savesaga.show();
           fmarkcut.value = true;
            fmarkcut.show();
           rcGroup.hide();
          }
  

   };
var t4 = shapeGroup.add ("iconbutton", undefined, i4, {style: "toolbutton", toggle: true}); // c
t4.onClick= function (){
if (t4.value == true){
    
    WShape = 'RRectangle';
arthne = arth - 2;
artwne = artw - 2;
cuttingsize.text = 'Cut Size:  '+artwne+'  x  '+arthne+' mm ' ;
Sizedetails.text = '2mm Bleed Rectangle Round Corner Size: '+RCDCC+' mm';

        t1.value = false; 
     t2.value = false; 
      t3.value = false; 
       t5.value = false; 
        t6.value = false; 
          t7.value = false;
           Savesaga.value = false ;
           Savesaga.show();
           fmarkcut.value = true;
            fmarkcut.show();
           rcGroup.hide();
         }
  
   };

var t5 = shapeGroup.add ("iconbutton", undefined, i5, {style: "toolbutton", toggle: true}); // c
t5.onClick= function (){
if (t5.value == true){
    
     WShape = 'Line';
    arthne = arth ;
artwne = artw ;
cuttingsize.text = 'Cut Size:  '+artwne+'  x  '+arthne+' mm ' ;
    Sizedetails.text = 'No Bleed Rectangle Line Cut';
    
         t1.value = false; 
        t2.value = false; 
     t3.value = false; 
      t4.value = false; 
       t10.value = false; 
         t6.value = false; 
        t7.value = false; 
         t8.value = false; 
           t9.value = false; 
            Savesaga.value = false ;
            Savesaga.show();
           fmarkcut.value = true;
            fmarkcut.show();
           rcGroup.hide();
          }

   };

var t6 = shapeGroup.add ("iconbutton", undefined, i6, {style: "toolbutton", toggle: true}); // c
t6.onClick= function (){
if (t6.value == true){
    
       WShape = 'LineB';
 arthne = arth - 2;
artwne = artw - 2;
cuttingsize.text = 'Cut Size:  '+artwne+'  x  '+arthne+' mm ' ;
Sizedetails.text = '2mm Bleed Single Color Background Rectangle';

       t1.value = false; 
        t2.value = false; 
     t3.value = false; 
      t4.value = false; 
       t5.value = false; 
         t10.value = false; 
        t7.value = false; 
         t8.value = false; 
           t9.value = false; 
               Savesaga.value = false ;
               Savesaga.show();
           fmarkcut.value = true;
            fmarkcut.show();
           rcGroup.hide();
           
           
         }
 
   };




   shapePanel2 = myDlg.add('panel', undefined, 'Paper Die Cut');
  shapePanel2.orientation = 'row';
  shapePanel2.alignment = 'center';
  shapePanel2.margins = [20, 20, 20, 10];
  shapeGroup2 = shapePanel2.add('group');
  shapeGroup2.orientation = 'row';
  shapeGroup2.alignChildren = 'center';
  shapeGroup2.graphics.font = 'Myriad Pro:14';  
  
  
  
  var t7 = shapeGroup2.add ("iconbutton", undefined, i7, {style: "toolbutton", toggle: true}); // c
t7.onClick= function (){
if (t7.value == true){
    

    
      WShape = '317Roundsaga';
arthne = arth - 0;
artwne = artw - 0;
cuttingsize.text = 'Cut Size:  '+artwne+'  x  '+arthne+' mm ' ;
Sizedetails.text = 'Die-cut Round No Bleed';

      t1.value = false; 
        t2.value = false; 
     t3.value = false; 
      t4.value = false; 
       t5.value = false; 
         t6.value = false; 
        t10.value = false; 
         t8.value = false; 
           t9.value = false; 
           Savesaga.value = true ;
       Savesaga.hide();
           fmarkcut.value = false;
           fmarkcut.hide();
        rcGroup.hide();
           
           
         }
  
   };

  var t8 = shapeGroup2.add ("iconbutton", undefined, i8, {style: "toolbutton", toggle: true}); // c
t8.onClick= function (){
if (t8.value == true){
    
    WShape = '317Roundbsaga';
arthne = arth - 3;
artwne = artw - 3;
cuttingsize.text = 'Cut Size:  '+artwne+'  x  '+arthne+' mm ' ;
Sizedetails.text = 'Die-cut Round 3mm Bleed';

        t1.value = false; 
        t2.value = false; 
     t3.value = false; 
      t4.value = false; 
       t5.value = false; 
         t6.value = false; 
        t7.value = false; 
         t10.value = false; 
           t9.value = false; 
           Savesaga.value = true ;
           fmarkcut.value = false;
            Savesaga.hide();
               fmarkcut.hide();
           rcGroup.hide();
           
         }
   
   };

  var t9 = shapeGroup2.add ("iconbutton", undefined, i9, {style: "toolbutton", toggle: true}); // c
t9.onClick= function (){
if (t9.value == true){
    
     WShape = '317RRectanglesaga';
arthne = arth - 0;
artwne = artw - 0;
cuttingsize.text = 'Cut Size:  '+artwne+'  x  '+arthne+' mm ' ;
Sizedetails.text = 'Die-cut Rectangle No Bleed';

    
       t1.value = false; 
        t2.value = false; 
     t3.value = false; 
      t4.value = false; 
       t5.value = false; 
         t6.value = false; 
        t7.value = false; 
         t8.value = false; 
           t10.value = false; 
           Savesaga.value = true ;
           fmarkcut.value = false;
           Savesaga.hide();
               fmarkcut.hide();
           rcGroup.show();
         }
   
   };

  var t10 = shapeGroup2.add ("iconbutton", undefined, i10, {style: "toolbutton", toggle: true}); // c
t10.onClick= function (){
if (t10.value == true){
    

    
    WShape = '317RRectanglebsaga';
arthne = arth - 3;
artwne = artw - 3;
cuttingsize.text = 'Cut Size:  '+artwne+'  x  '+arthne+' mm ' ;
Sizedetails.text = 'Die-cut Rectangle 3mm Bleed';

 
    
    t1.value = false; 
        t2.value = false; 
     t3.value = false; 
      t4.value = false; 
       t5.value = false; 
         t6.value = false; 
        t7.value = false; 
         t8.value = false; 
           t9.value = false; 
           Savesaga.value = true ;
           fmarkcut.value = false;
         Savesaga.hide();

               fmarkcut.hide();
            rcGroup.show();
         }
   
   };
  
  
 ppsize = shapeGroup2.add('group');
ppsize.orientation = 'column';
ppsize.alignChildren = 'left';




var _list_array = ['A3 size paper', '317 x 450 paper'];
var _list = ppsize.add('dropdownlist', undefined, undefined, { name: '_list', items: _list_array });
_list.size = [150,35]
_list.selection = 1;
_list.alignment = ['center', 'center'];
_list.graphics.font = 'Myriad Pro:18';  

sizePanel =myDlg.add('panel', undefined, 'Setting');
  sizePanel.orientation = 'row';
  sizePanel.margins = 15;
   sizePanel.alignment = 'fill';
  sizePanel.alignChildren = 'center';
    artworkGroup = sizePanel.add('group');
artworkGroup.orientation = 'column';

artworksize = artworkGroup.add('statictext', undefined, 'File Size:   '+artw+'  x   '+arth+' mm ');
artworksize.graphics.font = 'Myriad Pro:14';  
  artworksize.alignment = 'left';
 
cuttingsize = artworkGroup.add('statictext', undefined, 'Cutting: Please Select Shape');
cuttingsize.graphics.font = 'Myriad Pro:16';  
cuttingsize.characters = 25;
  cuttingsize.alignment = 'left';

Sizedetails = artworkGroup.add('statictext', undefined, '');
Sizedetails.graphics.font = 'Myriad Pro:14';  
Sizedetails.characters = 38;
  Sizedetails.alignment = 'left';


rcGroup = sizePanel.add('group');
rcGroup.orientation = 'column';
rcGroup.alignChildren = 'center';
rcGroup.graphics.font = 'Myriad Pro:14';  




  var RCDsizeauto = rcGroup.add('button', undefined, 'Auto RC');
  RCDsizeauto.graphics.font = 'Myriad Pro:16';  
 
  RCDsizeauto.onClick = function () { 
      SizedetailsRC.text = RCDCC;
      RCDSize1 = new UnitValue(SizedetailsRC.text, 'mm').as('pt');
      
      };
  

Sizedetails1 = rcGroup.add('statictext', undefined, 'Round Corner ');
Sizedetails1.graphics.font = 'Myriad Pro:15';  

SizedetailsRC = rcGroup.add('edittext',undefined,'0' );


RCDSize1 = new UnitValue(SizedetailsRC.text, 'mm').as('pt');

SizedetailsRC.onChange = function (){  
    
RCDSize1 = new UnitValue(SizedetailsRC.text, 'mm').as('pt');

}


Sizedetails3 = rcGroup.add('statictext', undefined, ' mm');

SizedetailsRC.graphics.font = 'Myriad Pro:20';  
SizedetailsRC.characters = 4;



fmarkGroup = sizePanel.add('group');
  fmarkGroup.orientation = 'column';
  fmarkGroup.alignChildren = 'center';
  fmarkGroup.graphics.font = 'Myriad Pro:14';  



fmarkcut = fmarkGroup.add('checkbox', undefined, ' Fmark AI' );
fmarkcut.graphics.font = 'Myriad Pro:18';  
fmarkcut.value = true;

   Savesaga = fmarkGroup.add('checkbox', undefined, ' Saga AI' );
      Savesaga.graphics.font = 'Myriad Pro:18';  
      Savesaga.value = false;



lastPanel =myDlg.add('panel', undefined, '');
  lastPanel.orientation = 'column';
  lastPanel.margins = 20;
  lastPanel.alignment = 'fill';
  lastPanel.alignChildren = 'center';
  
 okGroup = lastPanel.add('group');
okGroup.orientation = 'row';

  var prefixSt = okGroup.add('statictext', undefined, 'File Name:'); 
		prefixSt.size = [70,20]
       

		this.prefixEt = okGroup.add('edittext', undefined, this.prefix); 
		this.prefixEt.characters = 25;
       this.prefixEt.graphics.font = 'Myriad Pro:20';  


 (myDlg.ok = okGroup.add('button', undefined, 'OK' ));
 myDlg.ok.graphics.font = 'Arial:20';  
    myDlg.ok.onClick = shapeC;

 (myDlg.cancel = okGroup.add('button', undefined, 'Cancel' ));
 myDlg.cancel.graphics.font = 'Arial:20';  
    myDlg.cancel.onClick = actionCanceled; 

myDlg.show();  



function shapeC(){ 
           var myGrey= new CMYKColor()//initial default grey
myGrey.black=100;
 switch (WShape)
       {
           case 'Round' : 
           

           
           
rotateAmount1 = (Math.floor(395/selectedWidth2) ) *(Math.floor(288/selectedHeight1) );
 rotateAmount2 = (Math.floor(395/selectedHeight1) )*(Math.floor(288/selectedWidth2) );


  myObject1 = selected[0];  
    
if (rotateAmount1<rotateAmount2){
myObject1.rotate(90); 
doc.selection = null;
}

function removeArtboards ( arr ) { 

        var i = arr.length; while ( i-- ) activeDocument.artboards[ arr[i] ].remove(); 

    } 

    removeArtboards( [1] ); 
 doc.layers.getByName('Mark3').locked = false; 
                   doc.layers.getByName('Mark3').remove();


doc.layers.getByName('pic').hasSelectedArtwork = true; 

 selectedHeight = selection[0].height;
 selectedWidth = selection[0].width;

 selectedHeightcv = selectedHeight;
 selectedHeightcvlb = selection[0].height-5.7499675;

 selectedWidthcv = selectedWidth;
 selectedWidthcvlb = selection[0].width-5.7499675;


 selectedHeight11 = selectedHeight;
selectedHeight11 = new UnitValue(selectedHeight11, 'pt').as('mm');
selectedHeight11 = Math.round(selectedHeight11*10)/10;

 selectedWidth22 = selectedWidth;
selectedWidth22 = new UnitValue(selectedWidth22, 'pt').as('mm');
selectedWidth22 = Math.round(selectedWidth22*10)/10;


 OriginX = selection[0].position[0];
 OriginY = selection[0].position[1];

 selectedHeightlb = selection[0].height-5.7499675; 
 selectedWidthlb = selection[0].width-5.7499675;

           
           cutline1 = doc.pathItems.ellipse(-2.8346456692913207,2.8346456692913207,selectedWidth-5.669291338582641,selectedHeight-5.669291338582641);
           cutline1.stroked = true;
           cutline1.strokeWidth = 1;
           cutline1.filled = false;
           cutline1.strokeColor = myGrey;
           selectedPosition = selection.position;
           
            sLineW=0;
         sLineH=0;

myPageItem.position = [sLineW, -sLineH];


           repeatAmountw = Math.floor(395/selectedWidth22)-1;
           repeatAmounth = Math.floor(288/selectedHeight11)-1;
           
           margin =  (816.3779527559003-((repeatAmounth+1)*selectedHeightcv))/repeatAmounth;
          margin2 =  (1119.6850393700715-((repeatAmountw+1)*selectedWidthcv))/repeatAmountw;
          
           set1();
           set1cut();
           v = selectedWidth2;  
           h = selectedHeight1;  
           

           
 break;
 case 'RoundN' : 
 
          
           
rotateAmount1 = (Math.floor(395/selectedWidth2) ) *(Math.floor(288/selectedHeight1) );
 rotateAmount2 = (Math.floor(395/selectedHeight1) )*(Math.floor(288/selectedWidth2) );


  myObject1 = selected[0];  
    
if (rotateAmount1<rotateAmount2){
myObject1.rotate(90); 
doc.selection = null;
}
function removeArtboards ( arr ) { 

        var i = arr.length; while ( i-- ) activeDocument.artboards[ arr[i] ].remove(); 

    } 

    removeArtboards( [1] ); 

 doc.layers.getByName('Mark3').locked = false; 
                   doc.layers.getByName('Mark3').remove();

doc.layers.getByName('pic').hasSelectedArtwork = true; 

 selectedHeight = selection[0].height;
 selectedWidth = selection[0].width;

 selectedHeightcv = selectedHeight;
 selectedHeightcvlb = selection[0].height-5.7499675;

 selectedWidthcv = selectedWidth;
 selectedWidthcvlb = selection[0].width-5.7499675;


 selectedHeight11 = selectedHeight;
selectedHeight11 = new UnitValue(selectedHeight11, 'pt').as('mm');
selectedHeight11 = Math.round(selectedHeight11*10)/10;

 selectedWidth22 = selectedWidth;
selectedWidth22 = new UnitValue(selectedWidth22, 'pt').as('mm');
selectedWidth22 = Math.round(selectedWidth22*10)/10;


 OriginX = selection[0].position[0];
 OriginY = selection[0].position[1];

 selectedHeightlb = selection[0].height-5.7499675; 
 selectedWidthlb = selection[0].width-5.7499675;
 
 
 
cutline1 =  doc.pathItems.ellipse(0,0,selectedWidth,selectedHeight);
 cutline1.stroked = true;
           cutline1.strokeWidth = 1;
           cutline1.filled = false;
           cutline1.strokeColor = myGrey;
           selectedPosition = selection.position;
           
            sLineW=0;
         sLineH=0;

myPageItem.position = [sLineW, -sLineH];
           
           
     repeatAmountw = Math.floor(395/selectedWidth22)-1;
           repeatAmounth = Math.floor(288/selectedHeight11)-1;

 margin =  (816.3779527559003-((repeatAmounth+1)*selectedHeightcv))/repeatAmounth;
          margin2 =  (1119.6850393700715-((repeatAmountw+1)*selectedWidthcv))/repeatAmountw;
 set1();
           set1cut();
           v = selectedWidth2;  
           h = selectedHeight1;  


 break;
 case 'Rectangle' : 
 
          
           
rotateAmount1 = (Math.floor(395/selectedWidth2) ) *(Math.floor(288/selectedHeight1) );
 rotateAmount2 = (Math.floor(395/selectedHeight1) )*(Math.floor(288/selectedWidth2) );


  myObject1 = selected[0];  
    
if (rotateAmount1<rotateAmount2){
myObject1.rotate(90); 
doc.selection = null;
}

function removeArtboards ( arr ) { 

        var i = arr.length; while ( i-- ) activeDocument.artboards[ arr[i] ].remove(); 

    } 

    removeArtboards( [1] ); 
 doc.layers.getByName('Mark3').locked = false; 
                   doc.layers.getByName('Mark3').remove();


doc.layers.getByName('pic').hasSelectedArtwork = true; 

 selectedHeight = selection[0].height;
 selectedWidth = selection[0].width;

 selectedHeightcv = selectedHeight;
 selectedHeightcvlb = selection[0].height-5.7499675;

 selectedWidthcv = selectedWidth;
 selectedWidthcvlb = selection[0].width-5.7499675;


 selectedHeight11 = selectedHeight;
selectedHeight11 = new UnitValue(selectedHeight11, 'pt').as('mm');
selectedHeight11 = Math.round(selectedHeight11*10)/10;

 selectedWidth22 = selectedWidth;
selectedWidth22 = new UnitValue(selectedWidth22, 'pt').as('mm');
selectedWidth22 = Math.round(selectedWidth22*10)/10;


 OriginX = selection[0].position[0];
 OriginY = selection[0].position[1];

 selectedHeightlb = selection[0].height-5.7499675; 
 selectedWidthlb = selection[0].width-5.7499675;
 
 
cutline1 =  doc.pathItems.rectangle(-2.8346456692913207,2.8346456692913207,selectedWidth-5.669291338582641,selectedHeight-5.669291338582641);
 cutline1.stroked = true;
           cutline1.strokeWidth = 1;
           cutline1.filled = false;
           cutline1.strokeColor = myGrey;
           selectedPosition = selection.position;
           
            sLineW=0;
         sLineH=0;

myPageItem.position = [sLineW, -sLineH];
           
           
     repeatAmountw = Math.floor(395/selectedWidth22)-1;
           repeatAmounth = Math.floor(288/selectedHeight11)-1;

 margin =  (816.3779527559003-((repeatAmounth+1)*selectedHeightcv))/repeatAmounth;
          margin2 =  (1119.6850393700715-((repeatAmountw+1)*selectedWidthcv))/repeatAmountw;
 set1();
           set1cut();

v = selectedWidth2;  
           h = selectedHeight1;  

 break;
 case 'RRectangle' : 

         
           
rotateAmount1 = (Math.floor(395/selectedWidth2) ) *(Math.floor(288/selectedHeight1) );
 rotateAmount2 = (Math.floor(395/selectedHeight1) )*(Math.floor(288/selectedWidth2) );


  myObject1 = selected[0];  
    
if (rotateAmount1<rotateAmount2){
myObject1.rotate(90); 
doc.selection = null;
}

function removeArtboards ( arr ) { 

        var i = arr.length; while ( i-- ) activeDocument.artboards[ arr[i] ].remove(); 

    } 

    removeArtboards( [1] ); 

 doc.layers.getByName('Mark3').locked = false; 
                   doc.layers.getByName('Mark3').remove();

doc.layers.getByName('pic').hasSelectedArtwork = true; 

 selectedHeight = selection[0].height;
 selectedWidth = selection[0].width;

 selectedHeightcv = selectedHeight;
 selectedHeightcvlb = selection[0].height-5.7499675;

 selectedWidthcv = selectedWidth;
 selectedWidthcvlb = selection[0].width-5.7499675;


 selectedHeight11 = selectedHeight;
selectedHeight11 = new UnitValue(selectedHeight11, 'pt').as('mm');
selectedHeight11 = Math.round(selectedHeight11*10)/10;

 selectedWidth22 = selectedWidth;
selectedWidth22 = new UnitValue(selectedWidth22, 'pt').as('mm');
selectedWidth22 = Math.round(selectedWidth22*10)/10;


 OriginX = selection[0].position[0];
 OriginY = selection[0].position[1];

 selectedHeightlb = selection[0].height-5.7499675; 
 selectedWidthlb = selection[0].width-5.7499675;
 
cutline1 =  doc.pathItems.roundedRectangle(-2.8346456692913207,2.8346456692913207,selectedWidth-5.669291338582641,selectedHeight-5.669291338582641,RCD,RCD);
 cutline1.stroked = true;
           cutline1.strokeWidth = 1;
           cutline1.filled = false;
           cutline1.strokeColor = myGrey;
           selectedPosition = selection.position;
           
            sLineW=0;
         sLineH=0;

myPageItem.position = [sLineW, -sLineH];
           
           
     repeatAmountw = Math.floor(395/selectedWidth22)-1;
           repeatAmounth = Math.floor(288/selectedHeight11)-1;

 margin =  (816.3779527559003-((repeatAmounth+1)*selectedHeightcv))/repeatAmounth;
          margin2 =  (1119.6850393700715-((repeatAmountw+1)*selectedWidthcv))/repeatAmountw;
 set1();
           set1cut();
           v = selectedWidth2;  
           h = selectedHeight1;  


 break;
 case 'Line' : 

         
           
rotateAmount1 = (Math.floor(395/selectedWidth2) ) *(Math.floor(288/selectedHeight1) );
 rotateAmount2 = (Math.floor(395/selectedHeight1) )*(Math.floor(288/selectedWidth2) );


  myObject1 = selected[0];  
    
if (rotateAmount1<rotateAmount2){
myObject1.rotate(90); 
doc.selection = null;
}

function removeArtboards ( arr ) { 

        var i = arr.length; while ( i-- ) activeDocument.artboards[ arr[i] ].remove(); 

    } 

    removeArtboards( [1] ); 

 doc.layers.getByName('Mark3').locked = false; 
                   doc.layers.getByName('Mark3').remove();

doc.layers.getByName('pic').hasSelectedArtwork = true; 

 selectedHeight = selection[0].height;
 selectedWidth = selection[0].width;

 selectedHeightcv = selectedHeight;
 selectedHeightcvlb = selection[0].height-5.7499675;

 selectedWidthcv = selectedWidth;
 selectedWidthcvlb = selection[0].width-5.7499675;


 selectedHeight11 = selectedHeight;
selectedHeight11 = new UnitValue(selectedHeight11, 'pt').as('mm');
selectedHeight11 = Math.round(selectedHeight11*10)/10;

 selectedWidth22 = selectedWidth;
selectedWidth22 = new UnitValue(selectedWidth22, 'pt').as('mm');
selectedWidth22 = Math.round(selectedWidth22*10)/10;


 OriginX = selection[0].position[0];
 OriginY = selection[0].position[1];

 selectedHeightlb = selection[0].height-5.7499675; 
 selectedWidthlb = selection[0].width-5.7499675;

repeatAmountw = Math.floor(395/selectedWidth22)-1;
repeatAmounth = Math.floor(288/selectedHeight11)-1;

repeatAmountwl = Math.floor(395/selectedWidth22);
repeatAmounthl = Math.floor(288/selectedHeight11);

LineW = selectedWidthcv*(repeatAmountw+1);
LineH = selectedHeightcv*(repeatAmounth+1);

sLineW=(1119.6850393700715-LineW)/2;
LLineW = LineW + sLineW+2.8346456692913207;

sLineH=(816.3779527559003-LineH)/2;
LLineH = LineH + sLineH+2.8346456692913207;

myPageItem.position = [sLineW, -sLineH];

if (repeatAmounthl  % 2 == 0){
    
    cutline1 =  doc.pathItems.add();
cutline1.setEntirePath([ 
    [sLineW-1.4,-sLineH],
         [LLineW-1.4,-sLineH],

]);
 cutline1.stroked = true;
           cutline1.strokeWidth = 1;
           cutline1.filled = false;
           cutline1.strokeColor = myGrey;
           
           }
else { 

    cutline1 =  doc.pathItems.add();
cutline1.setEntirePath([ 
    [LLineW-1.4,-sLineH],
         [sLineW-1.4,-sLineH],

]);
 cutline1.stroked = true;
           cutline1.strokeWidth = 1;
           cutline1.filled = false;
           cutline1.strokeColor = myGrey;



    };


        



      margin = 0 ;
          margin2 =0 ;

set1();
line();
v = selectedWidth2;  
           h = selectedHeight1;  



 break;
 case 'LineB' : 
 
          
           
rotateAmount1 = (Math.floor(395/selectedWidth2) ) *(Math.floor(288/selectedHeight1) );
 rotateAmount2 = (Math.floor(395/selectedHeight1) )*(Math.floor(288/selectedWidth2) );


  myObject1 = selected[0];  
    
if (rotateAmount1<rotateAmount2){
myObject1.rotate(90); 
doc.selection = null;
}

function removeArtboards ( arr ) { 

        var i = arr.length; while ( i-- ) activeDocument.artboards[ arr[i] ].remove(); 

    } 

    removeArtboards( [1] ); 
 doc.layers.getByName('Mark3').locked = false; 
                   doc.layers.getByName('Mark3').remove();


doc.layers.getByName('pic').hasSelectedArtwork = true; 

 selectedHeight = selection[0].height;
 selectedWidth = selection[0].width;

 selectedHeightcv = selectedHeight;
 selectedHeightcvlb = selection[0].height-5.7499675;

 selectedWidthcv = selectedWidth;
 selectedWidthcvlb = selection[0].width-5.7499675;


 selectedHeight11 = selectedHeight;
selectedHeight11 = new UnitValue(selectedHeight11, 'pt').as('mm');
selectedHeight11 = Math.round(selectedHeight11*10)/10;

 selectedWidth22 = selectedWidth;
selectedWidth22 = new UnitValue(selectedWidth22, 'pt').as('mm');
selectedWidth22 = Math.round(selectedWidth22*10)/10;


 OriginX = selection[0].position[0];
 OriginY = selection[0].position[1];

 selectedHeightlb = selection[0].height-5.7499675; 
 selectedWidthlb = selection[0].width-5.7499675;
 
 
repeatAmountw = Math.floor(1116.8503937007804/selectedWidthcvlb)-1;
repeatAmounth = Math.floor(813.543307086609/selectedHeightcvlb)-1;         
           
repeatAmountwl = Math.floor(1116.8503937007804/selectedWidthcvlb);
repeatAmounthl = Math.floor(813.543307086609/selectedHeightcvlb);  
 
LineW = (selectedWidthcvlb*(repeatAmountw+1))+2.8346456692913207 ;
LineH = (selectedHeightcvlb*(repeatAmounth+1))+2.8346456692913207;

sLineW=(1119.6850393700715-LineW)/2;
LLineW = LineW + sLineW+2.8346456692913207;

sLineH=(816.3779527559003-LineH)/2;
LLineH = LineH + sLineH+2.8346456692913207;

myPageItem.position = [sLineW, -sLineH];


if (repeatAmounthl  % 2 == 0){



 
cutline1 =  doc.pathItems.add();
cutline1.setEntirePath([ 
[sLineW-1.4+2.8346456692913207,-sLineH-2.8346456692913207],
         [LLineW-1.4,-sLineH-2.8346456692913207],

]);
 cutline1.stroked = true;
           cutline1.strokeWidth = 1;
           cutline1.filled = false;
           cutline1.strokeColor = myGrey;
      
      
                 }
else { 
      
      
      cutline1 =  doc.pathItems.add();
cutline1.setEntirePath([ 
[LLineW-1.4,-sLineH-2.8346456692913207],
         [sLineW-1.4+2.8346456692913207,-sLineH-2.8346456692913207],

]);
 cutline1.stroked = true;
           cutline1.strokeWidth = 1;
           cutline1.filled = false;
           cutline1.strokeColor = myGrey;
      
   };

  margin = 0 ;
          margin2 =0 ;

set2();
lineb();
v = selectedWidthcvlb2;  
           h = selectedHeightcvlb2;

break;
 case '317Roundsaga' : 
 
       if(_list.selection.index == 1) {

selectedHeight1 = selection[0].height;
selectedHeight1 = new UnitValue(selectedHeight1, 'pt').as('mm');
selectedHeight1 = Math.round(selectedHeight1*10)/10;

selectedWidth2 = selection[0].width;
selectedWidth2 = new UnitValue(selectedWidth2, 'pt').as('mm');
selectedWidth2 = Math.round(selectedWidth2*10)/10;

    
rotateAmount1 = (Math.floor(420/selectedWidth2) ) *(Math.floor(297/selectedHeight1) );
 rotateAmount2 = (Math.floor(420/selectedHeight1) )*(Math.floor(297/selectedWidth2) );
 
  myObject1 = selected[0];  
    
if (rotateAmount1<rotateAmount2){
myObject1.rotate(90); 
doc.selection = null;
}

function removeArtboards ( arr ) { 

        var i = arr.length; while ( i-- ) activeDocument.artboards[ arr[i] ].remove(); 

    } 

    removeArtboards( [1] ); 

   
                doc.layers.getByName('Mark3').locked = false; 
                   doc.layers.getByName('Mark3').remove();

doc.layers.getByName('pic').hasSelectedArtwork = true; 

 selectedHeight = selection[0].height;
 selectedWidth = selection[0].width;

 selectedHeightcv = selectedHeight;
 selectedHeightcvlb = selection[0].height-5.7499675;

 selectedWidthcv = selectedWidth;
 selectedWidthcvlb = selection[0].width-5.7499675;


 selectedHeight11 = selectedHeight;
selectedHeight11 = new UnitValue(selectedHeight11, 'pt').as('mm');
selectedHeight11 = Math.round(selectedHeight11*10)/10;

 selectedWidth22 = selectedWidth;
selectedWidth22 = new UnitValue(selectedWidth22, 'pt').as('mm');
selectedWidth22 = Math.round(selectedWidth22*10)/10;


 OriginX = selection[0].position[0];
 OriginY = selection[0].position[1];

 selectedHeightlb = selection[0].height-5.7499675; 
 selectedWidthlb = selection[0].width-5.7499675;
 
cutline1 =  doc.pathItems.ellipse(-0.00001511811024,-56.692908503937,selectedWidth,selectedHeight);
 cutline1.stroked = true;
           cutline1.strokeWidth = 1;
           cutline1.filled = false;
           cutline1.strokeColor = myGrey;
           selectedPosition = selection.position;
           
            sLineW=-56.6929;
         sLineH=-0.00001511811024;
         
         

myPageItem.position = [sLineW, -sLineH];
           
           
     repeatAmountw = Math.floor(420/selectedWidth22)-1;
           repeatAmounth = Math.floor(297/selectedHeight11)-1;

 margin =  (841.889763779528-((repeatAmounth+1)*selectedHeightcv))/repeatAmounth;
          margin2 =  (1190.55118110236-((repeatAmountw+1)*selectedWidthcv))/repeatAmountw;
 set1();
           set1cut();
           v = selectedWidth2;  
           h = selectedHeight1;  

 } else {
     
     
selectedHeight1 = selection[0].height;
selectedHeight1 = new UnitValue(selectedHeight1, 'pt').as('mm');
selectedHeight1 = Math.round(selectedHeight1*10)/10;

selectedWidth2 = selection[0].width;
selectedWidth2 = new UnitValue(selectedWidth2, 'pt').as('mm');
selectedWidth2 = Math.round(selectedWidth2*10)/10;

    
rotateAmount1 = (Math.floor(396/selectedWidth2) ) *(Math.floor(276/selectedHeight1) );
 rotateAmount2 = (Math.floor(396/selectedHeight1) )*(Math.floor(276/selectedWidth2) );
 
  myObject1 = selected[0];  
    
if (rotateAmount1<rotateAmount2){
myObject1.rotate(90); 
doc.selection = null;
}

function removeArtboards ( arr ) { 

        var i = arr.length; while ( i-- ) activeDocument.artboards[ arr[i] ].remove(); 

    } 

    removeArtboards( [0] ); 

   
                doc.layers.getByName('Mark Layer').locked = false; 
                   doc.layers.getByName('Mark Layer').remove();
                   var currentLayer =  doc.layers.getByName('Mark3');  
            currentLayer.name= 'Mark Layer'; 
                   

doc.layers.getByName('pic').hasSelectedArtwork = true; 

 selectedHeight = selection[0].height;
 selectedWidth = selection[0].width;

 selectedHeightcv = selectedHeight;
 selectedHeightcvlb = selection[0].height-5.7499675;

 selectedWidthcv = selectedWidth;
 selectedWidthcvlb = selection[0].width-5.7499675;


 selectedHeight11 = selectedHeight;
selectedHeight11 = new UnitValue(selectedHeight11, 'pt').as('mm');
selectedHeight11 = Math.round(selectedHeight11*10)/10;

 selectedWidth22 = selectedWidth;
selectedWidth22 = new UnitValue(selectedWidth22, 'pt').as('mm');
selectedWidth22 = Math.round(selectedWidth22*10)/10;


 OriginX = selection[0].position[0];
 OriginY = selection[0].position[1];

 selectedHeightlb = selection[0].height-5.7499675; 
 selectedWidthlb = selection[0].width-5.7499675;
 
cutline1 =  doc.pathItems.ellipse(-0.00001511811024,20.3527559055118,selectedWidth,selectedHeight);
 cutline1.stroked = true;
           cutline1.strokeWidth = 1;
           cutline1.filled = false;
           cutline1.strokeColor = myGrey;
           selectedPosition = selection.position;
           
            sLineW=20.3527559055118;
         sLineH=-0.00001511811024;
         
         

myPageItem.position = [sLineW, -sLineH];
           
           
     repeatAmountw = Math.floor(396/selectedWidth22)-1;
           repeatAmounth = Math.floor(276/selectedHeight11)-1;

 margin =  (782.36220472441-((repeatAmounth+1)*selectedHeightcv))/repeatAmounth;
          margin2 =  (1122.51968503937-((repeatAmountw+1)*selectedWidthcv))/repeatAmountw;
 set1();
           set1cut();
           v = selectedWidth2;  
           h = selectedHeight1;  
           
           }
     
     
     
     

 break;
 
 
 case '317Roundbsaga' : 
 
 
   if(_list.selection.index == 1) {

selectedHeight1 = selection[0].height;
selectedHeight1 = new UnitValue(selectedHeight1, 'pt').as('mm');
selectedHeight1 = Math.round(selectedHeight1*10)/10;

selectedWidth2 = selection[0].width;
selectedWidth2 = new UnitValue(selectedWidth2, 'pt').as('mm');
selectedWidth2 = Math.round(selectedWidth2*10)/10;
           
           
rotateAmount1 = (Math.floor(420/selectedWidth2) ) *(Math.floor(297/selectedHeight1) );
 rotateAmount2 = (Math.floor(420/selectedHeight1) )*(Math.floor(297/selectedWidth2) );


  myObject1 = selected[0];  
    
if (rotateAmount1<rotateAmount2){
myObject1.rotate(90); 
doc.selection = null;
}


function removeArtboards ( arr ) { 

        var i = arr.length; while ( i-- ) activeDocument.artboards[ arr[i] ].remove(); 

    } 

    removeArtboards( [1] ); 

   
                doc.layers.getByName('Mark3').locked = false; 
                   doc.layers.getByName('Mark3').remove();



doc.layers.getByName('pic').hasSelectedArtwork = true; 

 selectedHeight = selection[0].height;
 selectedWidth = selection[0].width;

 selectedHeightcv = selectedHeight;
 selectedHeightcvlb = selection[0].height-5.7499675;

 selectedWidthcv = selectedWidth;
 selectedWidthcvlb = selection[0].width-5.7499675;


 selectedHeight11 = selectedHeight;
selectedHeight11 = new UnitValue(selectedHeight11, 'pt').as('mm');
selectedHeight11 = Math.round(selectedHeight11*10)/10;

 selectedWidth22 = selectedWidth;
selectedWidth22 = new UnitValue(selectedWidth22, 'pt').as('mm');
selectedWidth22 = Math.round(selectedWidth22*10)/10;


 OriginX = selection[0].position[0];
 OriginY = selection[0].position[1];

 selectedHeightlb = selection[0].height-5.7499675; 
 selectedWidthlb = selection[0].width-5.7499675;
 
cutline1 =  doc.pathItems.ellipse(-4.25195338582677,-52.44094,selectedWidth-8.5024,selectedHeight-8.5024);
 cutline1.stroked = true;
           cutline1.strokeWidth = 1;
           cutline1.filled = false;
           cutline1.strokeColor = myGrey;
           selectedPosition = selection.position;
           
            sLineW=-56.6929;
            
         sLineH=-0.00001511811024;
         


myPageItem.position = [sLineW, -sLineH];
           
           
     repeatAmountw = Math.floor(420/selectedWidth22)-1;
           repeatAmounth = Math.floor(297/selectedHeight11)-1;

 margin =  (841.889763779528-((repeatAmounth+1)*selectedHeightcv))/repeatAmounth;
          margin2 =  (1190.55118110236-((repeatAmountw+1)*selectedWidthcv))/repeatAmountw;
 set1();
           set1cut();
           v = selectedWidth2;  
           h = selectedHeight1;  
           
            } else {
     
           selectedHeight1 = selection[0].height;
selectedHeight1 = new UnitValue(selectedHeight1, 'pt').as('mm');
selectedHeight1 = Math.round(selectedHeight1*10)/10;

selectedWidth2 = selection[0].width;
selectedWidth2 = new UnitValue(selectedWidth2, 'pt').as('mm');
selectedWidth2 = Math.round(selectedWidth2*10)/10;

rotateAmount1 = (Math.floor(396/selectedWidth2) ) *(Math.floor(276/selectedHeight1) );
 rotateAmount2 = (Math.floor(396/selectedHeight1) )*(Math.floor(276/selectedWidth2) );
 
  myObject1 = selected[0];  
    
if (rotateAmount1<rotateAmount2){
myObject1.rotate(90); 
doc.selection = null;
}

function removeArtboards ( arr ) { 

        var i = arr.length; while ( i-- ) activeDocument.artboards[ arr[i] ].remove(); 

    } 

    removeArtboards( [0] ); 

   
                doc.layers.getByName('Mark Layer').locked = false; 
                   doc.layers.getByName('Mark Layer').remove();
                   var currentLayer =  doc.layers.getByName('Mark3');  
            currentLayer.name= 'Mark Layer'; 
            
            doc.layers.getByName('pic').hasSelectedArtwork = true; 

 selectedHeight = selection[0].height;
 selectedWidth = selection[0].width;

 selectedHeightcv = selectedHeight;
 selectedHeightcvlb = selection[0].height-5.7499675;

 selectedWidthcv = selectedWidth;
 selectedWidthcvlb = selection[0].width-5.7499675;


 selectedHeight11 = selectedHeight;
selectedHeight11 = new UnitValue(selectedHeight11, 'pt').as('mm');
selectedHeight11 = Math.round(selectedHeight11*10)/10;

 selectedWidth22 = selectedWidth;
selectedWidth22 = new UnitValue(selectedWidth22, 'pt').as('mm');
selectedWidth22 = Math.round(selectedWidth22*10)/10;


 OriginX = selection[0].position[0];
 OriginY = selection[0].position[1];

 selectedHeightlb = selection[0].height-5.7499675; 
 selectedWidthlb = selection[0].width-5.7499675;
 
cutline1 =  doc.pathItems.ellipse(-4.25195338582677,24.6047292913386,selectedWidth-8.5024,selectedHeight-8.5024);
 cutline1.stroked = true;
           cutline1.strokeWidth = 1;
           cutline1.filled = false;
           cutline1.strokeColor = myGrey;
           selectedPosition = selection.position;
           
            sLineW=20.3527692913386;
            
         sLineH=-0.00001511811024;
         


myPageItem.position = [sLineW, -sLineH];
           
           
         
     repeatAmountw = Math.floor(396/selectedWidth22)-1;
           repeatAmounth = Math.floor(276/selectedHeight11)-1;

 margin =  (782.36220472441-((repeatAmounth+1)*selectedHeightcv))/repeatAmounth;
          margin2 =  (1122.51968503937-((repeatAmountw+1)*selectedWidthcv))/repeatAmountw;
 set1();
           set1cut();
           v = selectedWidth2;  
           h = selectedHeight1;  

}
           

 break;
 
 
  case '317RRectanglesaga' : 


 if(_list.selection.index == 1) {

selectedHeight1 = selection[0].height;
selectedHeight1 = new UnitValue(selectedHeight1, 'pt').as('mm');
selectedHeight1 = Math.round(selectedHeight1*10)/10;

selectedWidth2 = selection[0].width;
selectedWidth2 = new UnitValue(selectedWidth2, 'pt').as('mm');
selectedWidth2 = Math.round(selectedWidth2*10)/10;

    
rotateAmount1 = (Math.floor(420/selectedWidth2) ) *(Math.floor(297/selectedHeight1) );
 rotateAmount2 = (Math.floor(420/selectedHeight1) )*(Math.floor(297/selectedWidth2) );
 
  myObject1 = selected[0];  
    
if (rotateAmount1<rotateAmount2){
myObject1.rotate(90); 
doc.selection = null;
}

function removeArtboards ( arr ) { 

        var i = arr.length; while ( i-- ) activeDocument.artboards[ arr[i] ].remove(); 

    } 

    removeArtboards( [1] ); 

   
                doc.layers.getByName('Mark3').locked = false; 
                   doc.layers.getByName('Mark3').remove();

doc.layers.getByName('pic').hasSelectedArtwork = true; 

 selectedHeight = selection[0].height;
 selectedWidth = selection[0].width;

 selectedHeightcv = selectedHeight;
 selectedHeightcvlb = selection[0].height-5.7499675;

 selectedWidthcv = selectedWidth;
 selectedWidthcvlb = selection[0].width-5.7499675;


 selectedHeight11 = selectedHeight;
selectedHeight11 = new UnitValue(selectedHeight11, 'pt').as('mm');
selectedHeight11 = Math.round(selectedHeight11*10)/10;

 selectedWidth22 = selectedWidth;
selectedWidth22 = new UnitValue(selectedWidth22, 'pt').as('mm');
selectedWidth22 = Math.round(selectedWidth22*10)/10;


 OriginX = selection[0].position[0];
 OriginY = selection[0].position[1];

 selectedHeightlb = selection[0].height-5.7499675; 
 selectedWidthlb = selection[0].width-5.7499675;
 
cutline1 =  doc.pathItems.roundedRectangle(-0.00001511811024,-56.692908503937,selectedWidth,selectedHeight,RCDSize1,RCDSize1);
cutline1.stroked = true;
           cutline1.strokeWidth = 1;
           cutline1.filled = false;
           cutline1.strokeColor = myGrey;
           selectedPosition = selection.position;
           
            sLineW=-56.6929;
         sLineH=-0.00001511811024;
         
         

myPageItem.position = [sLineW, -sLineH];
           
           
     repeatAmountw = Math.floor(420/selectedWidth22)-1;
           repeatAmounth = Math.floor(297/selectedHeight11)-1;

 margin =  (841.889763779528-((repeatAmounth+1)*selectedHeightcv))/repeatAmounth;
          margin2 =  (1190.55118110236-((repeatAmountw+1)*selectedWidthcv))/repeatAmountw;
 set1();
           set1cut();
           v = selectedWidth2;  
           h = selectedHeight1;  

 } else {
     
     
selectedHeight1 = selection[0].height;
selectedHeight1 = new UnitValue(selectedHeight1, 'pt').as('mm');
selectedHeight1 = Math.round(selectedHeight1*10)/10;

selectedWidth2 = selection[0].width;
selectedWidth2 = new UnitValue(selectedWidth2, 'pt').as('mm');
selectedWidth2 = Math.round(selectedWidth2*10)/10;

    
rotateAmount1 = (Math.floor(396/selectedWidth2) ) *(Math.floor(276/selectedHeight1) );
 rotateAmount2 = (Math.floor(396/selectedHeight1) )*(Math.floor(276/selectedWidth2) );
 
  myObject1 = selected[0];  
    
if (rotateAmount1<rotateAmount2){
myObject1.rotate(90); 
doc.selection = null;
}

function removeArtboards ( arr ) { 

        var i = arr.length; while ( i-- ) activeDocument.artboards[ arr[i] ].remove(); 

    } 

    removeArtboards( [0] ); 

   
                doc.layers.getByName('Mark Layer').locked = false; 
                   doc.layers.getByName('Mark Layer').remove();
                   var currentLayer =  doc.layers.getByName('Mark3');  
            currentLayer.name= 'Mark Layer'; 
                   

doc.layers.getByName('pic').hasSelectedArtwork = true; 

 selectedHeight = selection[0].height;
 selectedWidth = selection[0].width;

 selectedHeightcv = selectedHeight;
 selectedHeightcvlb = selection[0].height-5.7499675;

 selectedWidthcv = selectedWidth;
 selectedWidthcvlb = selection[0].width-5.7499675;


 selectedHeight11 = selectedHeight;
selectedHeight11 = new UnitValue(selectedHeight11, 'pt').as('mm');
selectedHeight11 = Math.round(selectedHeight11*10)/10;

 selectedWidth22 = selectedWidth;
selectedWidth22 = new UnitValue(selectedWidth22, 'pt').as('mm');
selectedWidth22 = Math.round(selectedWidth22*10)/10;


 OriginX = selection[0].position[0];
 OriginY = selection[0].position[1];

 selectedHeightlb = selection[0].height-5.7499675; 
 selectedWidthlb = selection[0].width-5.7499675;
 
cutline1 =  doc.pathItems.roundedRectangle(-0.00001511811024,20.3527559055118,selectedWidth,selectedHeight,RCDSize1,RCDSize1);
 cutline1.stroked = true;
           cutline1.strokeWidth = 1;
           cutline1.filled = false;
           cutline1.strokeColor = myGrey;
           selectedPosition = selection.position;
           
            sLineW=20.3527559055118;
         sLineH=-0.00001511811024;
         
         

myPageItem.position = [sLineW, -sLineH];
           
           
     repeatAmountw = Math.floor(396/selectedWidth22)-1;
           repeatAmounth = Math.floor(276/selectedHeight11)-1;

 margin =  (782.36220472441-((repeatAmounth+1)*selectedHeightcv))/repeatAmounth;
          margin2 =  (1122.51968503937-((repeatAmountw+1)*selectedWidthcv))/repeatAmountw;
 set1();
           set1cut();
           v = selectedWidth2;  
           h = selectedHeight1;  
           
           }
     


 

 break;
 
  case '317RRectanglebsaga' : 
  

 
   if(_list.selection.index == 1) {

selectedHeight1 = selection[0].height;
selectedHeight1 = new UnitValue(selectedHeight1, 'pt').as('mm');
selectedHeight1 = Math.round(selectedHeight1*10)/10;

selectedWidth2 = selection[0].width;
selectedWidth2 = new UnitValue(selectedWidth2, 'pt').as('mm');
selectedWidth2 = Math.round(selectedWidth2*10)/10;
           
           
rotateAmount1 = (Math.floor(420/selectedWidth2) ) *(Math.floor(297/selectedHeight1) );
 rotateAmount2 = (Math.floor(420/selectedHeight1) )*(Math.floor(297/selectedWidth2) );


  myObject1 = selected[0];  
    
if (rotateAmount1<rotateAmount2){
myObject1.rotate(90); 
doc.selection = null;
}


function removeArtboards ( arr ) { 

        var i = arr.length; while ( i-- ) activeDocument.artboards[ arr[i] ].remove(); 

    } 

    removeArtboards( [1] ); 

   
                doc.layers.getByName('Mark3').locked = false; 
                   doc.layers.getByName('Mark3').remove();



doc.layers.getByName('pic').hasSelectedArtwork = true; 

 selectedHeight = selection[0].height;
 selectedWidth = selection[0].width;

 selectedHeightcv = selectedHeight;
 selectedHeightcvlb = selection[0].height-5.7499675;

 selectedWidthcv = selectedWidth;
 selectedWidthcvlb = selection[0].width-5.7499675;


 selectedHeight11 = selectedHeight;
selectedHeight11 = new UnitValue(selectedHeight11, 'pt').as('mm');
selectedHeight11 = Math.round(selectedHeight11*10)/10;

 selectedWidth22 = selectedWidth;
selectedWidth22 = new UnitValue(selectedWidth22, 'pt').as('mm');
selectedWidth22 = Math.round(selectedWidth22*10)/10;


 OriginX = selection[0].position[0];
 OriginY = selection[0].position[1];

 selectedHeightlb = selection[0].height-5.7499675; 
 selectedWidthlb = selection[0].width-5.7499675;
 
cutline1 =  doc.pathItems.roundedRectangle(-4.25195338582677,-52.44094,selectedWidth-8.5024,selectedHeight-8.5024,RCDSize1,RCDSize1);
 cutline1.stroked = true;
           cutline1.strokeWidth = 1;
           cutline1.filled = false;
           cutline1.strokeColor = myGrey;
           selectedPosition = selection.position;
           
            sLineW=-56.6929;
            
         sLineH=-0.00001511811024;
         


myPageItem.position = [sLineW, -sLineH];
           
           
     repeatAmountw = Math.floor(420/selectedWidth22)-1;
           repeatAmounth = Math.floor(297/selectedHeight11)-1;

 margin =  (841.889763779528-((repeatAmounth+1)*selectedHeightcv))/repeatAmounth;
          margin2 =  (1190.55118110236-((repeatAmountw+1)*selectedWidthcv))/repeatAmountw;
 set1();
           set1cut();
           v = selectedWidth2;  
           h = selectedHeight1;  
           
            } else {
     
           selectedHeight1 = selection[0].height;
selectedHeight1 = new UnitValue(selectedHeight1, 'pt').as('mm');
selectedHeight1 = Math.round(selectedHeight1*10)/10;

selectedWidth2 = selection[0].width;
selectedWidth2 = new UnitValue(selectedWidth2, 'pt').as('mm');
selectedWidth2 = Math.round(selectedWidth2*10)/10;

rotateAmount1 = (Math.floor(396/selectedWidth2) ) *(Math.floor(276/selectedHeight1) );
 rotateAmount2 = (Math.floor(396/selectedHeight1) )*(Math.floor(276/selectedWidth2) );
 
  myObject1 = selected[0];  
    
if (rotateAmount1<rotateAmount2){
myObject1.rotate(90); 
doc.selection = null;
}

function removeArtboards ( arr ) { 

        var i = arr.length; while ( i-- ) activeDocument.artboards[ arr[i] ].remove(); 

    } 

    removeArtboards( [0] ); 

   
                doc.layers.getByName('Mark Layer').locked = false; 
                   doc.layers.getByName('Mark Layer').remove();
                   var currentLayer =  doc.layers.getByName('Mark3');  
            currentLayer.name= 'Mark Layer'; 
            
            doc.layers.getByName('pic').hasSelectedArtwork = true; 

 selectedHeight = selection[0].height;
 selectedWidth = selection[0].width;

 selectedHeightcv = selectedHeight;
 selectedHeightcvlb = selection[0].height-5.7499675;

 selectedWidthcv = selectedWidth;
 selectedWidthcvlb = selection[0].width-5.7499675;


 selectedHeight11 = selectedHeight;
selectedHeight11 = new UnitValue(selectedHeight11, 'pt').as('mm');
selectedHeight11 = Math.round(selectedHeight11*10)/10;

 selectedWidth22 = selectedWidth;
selectedWidth22 = new UnitValue(selectedWidth22, 'pt').as('mm');
selectedWidth22 = Math.round(selectedWidth22*10)/10;


 OriginX = selection[0].position[0];
 OriginY = selection[0].position[1];

 selectedHeightlb = selection[0].height-5.7499675; 
 selectedWidthlb = selection[0].width-5.7499675;
 
cutline1 =  doc.pathItems.roundedRectangle(-4.25195338582677,24.6047292913386,selectedWidth-8.5024,selectedHeight-8.5024,RCDSize1,RCDSize1);
 cutline1.stroked = true;
           cutline1.strokeWidth = 1;
           cutline1.filled = false;
           cutline1.strokeColor = myGrey;
           selectedPosition = selection.position;
           
            sLineW=20.3527692913386;
            
         sLineH=-0.00001511811024;
         


myPageItem.position = [sLineW, -sLineH];
           
           
         
     repeatAmountw = Math.floor(396/selectedWidth22)-1;
           repeatAmounth = Math.floor(276/selectedHeight11)-1;

 margin =  (782.36220472441-((repeatAmounth+1)*selectedHeightcv))/repeatAmounth;
          margin2 =  (1122.51968503937-((repeatAmountw+1)*selectedWidthcv))/repeatAmountw;
 set1();
           set1cut();
           v = selectedWidth2;  
           h = selectedHeight1;  

}
 

 break;
 

 
 
default:

alert('大哥，至少也要点一个呀' )
return; 


}





myDlg.close(0);  

}



function actionCanceled() { 

    myDlg.close();
    app.executeMenuCommand ('undo'); 
 

}


function set1(){ 
    
    

    


selectedPosition = selection.position;
	var newItem = selection[0].duplicate( doc, ElementPlacement.PLACEATEND);  
	for(i=0; i< repeatAmounth ; i++){
		var newPosY = selectedHeight - newItem.position[1] + margin;
        
		newItem.position = [newItem.position[0],-newPosY];
       
		doc.selection = null;
		newItem.selected = true;
		var selectedPosition = selection.position;
		newItem.duplicate( doc, ElementPlacement.PLACEATEND );  
      

     }  
newItem.remove();
doc.layers.getByName('pic').hasSelectedArtwork = true; 
app.executeMenuCommand ('group');

selectedPosition = selection.position;
	var newItem = selection[0].duplicate( doc, ElementPlacement.PLACEATEND);  
    
	for(i=0; i< repeatAmountw ; i++){
		var newPosX = selectedWidth + newItem.position[0] + margin2;
       
		newItem.position = [newPosX,newItem.position[1]];
       
		doc.selection = null;
		newItem.selected = true;
		var selectedPosition = selection.position;
		newItem.duplicate( doc, ElementPlacement.PLACEATEND );  
    }
   newItem.remove();
doc.layers.getByName('pic').hasSelectedArtwork = true; 

app.executeMenuCommand ('ungroup');

tt = selection.length;

 myDlg.close(0);  
alert('total: ' + tt+' pcs' )
}


function set2(){ 
    
    

    


selectedPosition = selection.position;
	var newItem = selection[0].duplicate( doc, ElementPlacement.PLACEATEND);  
	for(i=0; i< repeatAmounth ; i++){
		var newPosY = selectedHeightlb - newItem.position[1] + margin;
        
		newItem.position = [newItem.position[0],-newPosY];
       
		doc.selection = null;
		newItem.selected = true;
		var selectedPosition = selection.position;
		newItem.duplicate( doc, ElementPlacement.PLACEATEND );  
      

     }  
newItem.remove();
doc.layers.getByName('pic').hasSelectedArtwork = true; 
app.executeMenuCommand ('group');

selectedPosition = selection.position;
	var newItem = selection[0].duplicate( doc, ElementPlacement.PLACEATEND);  
    
	for(i=0; i< repeatAmountw ; i++){
		var newPosX = selectedWidthlb + newItem.position[0] + margin2;
       
		newItem.position = [newPosX,newItem.position[1]];
       
		doc.selection = null;
		newItem.selected = true;
		var selectedPosition = selection.position;
		newItem.duplicate( doc, ElementPlacement.PLACEATEND );  
    }
   newItem.remove();
doc.layers.getByName('pic').hasSelectedArtwork = true; 

app.executeMenuCommand ('ungroup');

tt = selection.length;
 myDlg.close(0);  

alert('total: ' + tt+' pcs' )

}





function set1cut(){ 


selectedPosition = selection.position;
    var newItem2 = cutline1.duplicate( doc, ElementPlacement.PLACEATBEGINNING);  
for(i=0; i< repeatAmounth ; i++){
var newPosY2 = selectedHeight - newItem2.position[1] + margin;
 newItem2.position = [newItem2.position[0],-newPosY2];
var selectedPosition = selection.position;
  newItem2.duplicate( doc, ElementPlacement.PLACEATBEGINNING);  

}
newItem2.remove();
doc.selection = null;
doc.layers.getByName('Cut Layer').hasSelectedArtwork = true; 

app.executeMenuCommand ('group');

selectedPosition = selection.position;
	
    var newItem2 = selection[0].duplicate( doc, ElementPlacement.PLACEATBEGINNING);  
	for(i=0; i< repeatAmountw ; i++){
		
        var newPosX2 = selectedWidth + newItem2.position[0] + margin2;
		
        newItem2.position = [newPosX2,newItem2.position[1]];
		doc.selection = null;
		newItem2.selected = true;
		var selectedPosition = selection.position;
		
        newItem2.duplicate( doc, ElementPlacement.PLACEATBEGINNING );  
	}


newItem2.remove();

doc.layers.getByName('Cut Layer').hasSelectedArtwork = true; 
app.executeMenuCommand ('ungroup');


}




function line(){ 
  var myGrey= new CMYKColor()//initial default grey
myGrey.black=100;
    
selectedPosition = selection.position;
    var newItem2 = cutline1.duplicate( doc, ElementPlacement.PLACEATBEGINNING);  
for(i= 0; i< repeatAmounthl ; i++){
var newPosY2 = selectedHeight - newItem2.position[1] ;
 newItem2.rotate(180);
 newItem2.position = [newItem2.position[0],-newPosY2];
var selectedPosition = selection.position;
  newItem2.duplicate( doc, ElementPlacement.PLACEATBEGINNING );  

}
newItem2.remove();


var cutline2 = doc.pathItems.add();
cutline2.setEntirePath([ 
    [sLineW,-sLineH+1.4],
         [sLineW,-LLineH+1.4],

]);

cutline2.stroked = true;
cutline2.filled = false;
cutline2.strokeWidth = 1;
cutline2.strokeColor = myGrey;
selectedPosition = selection.position;
	
    var newItem2 = cutline2.duplicate( doc, ElementPlacement.PLACEATBEGINNING);  
	for(i=0; i< repeatAmountwl ; i++){
		
        var newPosX2 = selectedWidth + newItem2.position[0] ;
        newItem2.rotate(180);
	
        newItem2.position = [newPosX2,newItem2.position[1]];
		doc.selection = null;
		newItem2.selected = true;
		var selectedPosition = selection.position;
		
        newItem2.duplicate( doc, ElementPlacement.PLACEATBEGINNING );  
	}
newItem2.remove();

doc.layers.getByName('Cut Layer').hasSelectedArtwork = true; 
app.executeMenuCommand ('ungroup');
 
 }


function lineb(){ 
    
  var myGrey= new CMYKColor()//initial default grey
myGrey.black=100;
    

    selectedPosition = selection.position;
    var newItem2 = cutline1.duplicate( doc, ElementPlacement.PLACEATBEGINNING);  
for(i=0; i< repeatAmounthl ; i++){
var newPosY2 = selectedHeightlb - newItem2.position[1];
  newItem2.rotate(180);
 newItem2.position = [newItem2.position[0],-newPosY2];
 
var selectedPosition = selection.position;
  newItem2.duplicate( doc, ElementPlacement.PLACEATBEGINNING );  

}
newItem2.remove();


var cutline2 = doc.pathItems.add();
cutline2.setEntirePath([ 
 [sLineW+2.8346456692913207,-sLineH-2.8346456692913207+1.4],
         [sLineW+2.8346456692913207,-LLineH+1.4],

]);

cutline2.stroked = true;
cutline2.filled = false;
cutline2.strokeWidth = 1;
cutline2.strokeColor = myGrey;
selectedPosition = selection.position;
	
    var newItem2 = cutline2.duplicate( doc, ElementPlacement.PLACEATBEGINNING);  
	for(i=0; i< repeatAmountwl ; i++){
		
        var newPosX2 = selectedWidthlb + newItem2.position[0];
		  newItem2.rotate(180);
        newItem2.position = [newPosX2,newItem2.position[1]];
        
		doc.selection = null;
		newItem2.selected = true;
		var selectedPosition = selection.position;
		
        newItem2.duplicate( doc, ElementPlacement.PLACEATBEGINNING );  
	}
newItem2.remove();

doc.layers.getByName('Cut Layer').hasSelectedArtwork = true; 
app.executeMenuCommand ('ungroup');



}





 sizewp();

function sizewp(){ 
   

   
   

    var text2 = doc.textFrames.add();
text2.position = [1119.6850393700715,15];
text2.story.textRange.justification = Justification.RIGHT;
text2.textRange.characterAttributes.size = 8;
text2.contents = tt+'pcs  |  '+arthne+'  x '+artwne+'mm'; 
text2.moveToBeginning(doc.layers['pic']);



}



doc.layers.getByName('Mark').locked = false; 
doc.layers.getByName('Mark').move( activeDocument, ElementPlacement.PLACEATBEGINNING );  
doc.layers.getByName('Mark').locked = true; 






try {
    // uncomment to suppress Illustrator warning dialogs
    // app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

    if (app.documents.length > 0 ) {

        // Get the folder to save the files into
        var destFolder = null;
        destFolder = Folder.selectDialog( 'Select Folder to Save PDF & .ai Cutting File ', '~/Desktop/' );

        if (destFolder != null) {
            var options, i, sourceDoc, targetFile;  
         sourceDoc = app.activeDocument;
            // Get the PDF options to be used
         
                   
                  
            // You can tune these by changing the code in the getOptions() function.

        // returns the document object

                // Get the file to save the document as pdf into
          
                

                savedata ();

function savedata(){ 

if(fmarkcut.value == true & Savesaga.value == false ) {  
    
             options = this.getOptions();
                targetFile = this.getTargetFile(sourceDoc.name, '.pdf', destFolder);
                sourceDoc.saveAs( targetFile, options );     
    
    
         targetFile2 = this.getTargetFile2(sourceDoc.name, '.ai',destFolder);
                  options2 = this.getOptions2();
                
                doc.layers.getByName('pic').remove();
                doc.layers.getByName('Mark Layer').locked = false; 
                doc.layers.getByName('Mark Layer').remove();
                        doc.layers.getByName('Extra Layer').locked = false; 
                doc.layers.getByName('Extra Layer').remove();
                sourceDoc.saveAs( targetFile2, options2 );
}
else if (Savesaga.value == true & fmarkcut.value == false) {
    
    doc.layers.getByName('Mark').locked = false; 
                doc.layers.getByName('Mark').remove();
    
    
             options = this.getOptions();
                targetFile = this.getTargetFile(sourceDoc.name, '.pdf', destFolder);
                sourceDoc.saveAs( targetFile, options );     
    
    
 targetFile3 = this.getTargetFile3(sourceDoc.name, '.ai',destFolder);
                
                   options3 = this.getOptions3();

              
doc.layers.getByName('pic').remove();

        
                  doc.layers.getByName('Mark Layer').locked = false; 
                  doc.layers.getByName('Extra Layer').locked = false; 


doc.layers.getByName('Mark Layer').move( activeDocument, ElementPlacement.PLACEATBEGINNING );  
doc.layers.getByName('Extra Layer').move( activeDocument, ElementPlacement.PLACEATBEGINNING ); 




roa();
             doc.layers.getByName('Extra Layer').locked = true; 
                  doc.layers.getByName('Mark Layer').locked = true; 
             
             
               sourceDoc.saveAs( targetFile3, options3);

   }
else if (Savesaga.value == true & fmarkcut.value == true) {
    
    
    options = this.getOptions();
                targetFile = this.getTargetFile(sourceDoc.name, '.pdf', destFolder);
                sourceDoc.saveAs( targetFile, options );     
    
    
         targetFile2 = this.getTargetFile2(sourceDoc.name, '.ai',destFolder);
                  options2 = this.getOptions2();
                
                doc.layers.getByName('pic').remove();
                doc.layers.getByName('Mark Layer').locked = false; 
                doc.layers.getByName('Mark Layer').remove();
                        doc.layers.getByName('Extra Layer').locked = false; 
                doc.layers.getByName('Extra Layer').remove();
                sourceDoc.saveAs( targetFile2, options2 );
    
    
    
    
 targetFile3 = this.getTargetFile3(sourceDoc.name, '.ai',destFolder);
                
                   options3 = this.getOptions3();
                app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
                app.executeMenuCommand("undo");
              
doc.layers.getByName('pic').remove();
doc.layers.getByName('Mark').locked = false; 
                doc.layers.getByName('Mark').remove();
        
                  doc.layers.getByName('Mark Layer').locked = false; 
                  doc.layers.getByName('Extra Layer').locked = false; 


doc.layers.getByName('Mark Layer').move( activeDocument, ElementPlacement.PLACEATBEGINNING );  
doc.layers.getByName('Extra Layer').move( activeDocument, ElementPlacement.PLACEATBEGINNING ); 




roa();
             doc.layers.getByName('Extra Layer').locked = true; 
                  doc.layers.getByName('Mark Layer').locked = true; 
             
             
               sourceDoc.saveAs( targetFile3, options3);
               
}

};

                
                
                
                
                
       

               
               sourceDoc.close();
             
             
        }
    }
    else{
        throw new Error('There are no document open!');
    }
}
catch(e) {
    alert( e.message, 'Script Alert', true);
}

/** Returns the options to be used for the generated files.
    @return PDFSaveOptions object
*/
function getOptions() {
    var NamePreset = 'Print PDF';
    // Create the required options object
    var options = new PDFSaveOptions();

    
  options.pDFPreset='[High Quality Print]';


    return options;
}


function getOptions2() {
    var NamePreset = 'ai';
    // Create the required options object

    var options = new IllustratorSaveOptions();
    

  options.compatibility = Compatibility.ILLUSTRATOR8;

    return options;
}

function getOptions3() {
    var NamePreset = 'ai';
    // Create the required options object

    var options = new IllustratorSaveOptions();
    

  options.compatibility = Compatibility.ILLUSTRATOR16;

    return options;
}




/** Returns the file to save or export the document into.
    @param docName the name of the document
    @param ext the extension the file extension to be applied
    @param destFolder the output folder
    @return File object
*/


function roa() {
var doc = app.activeDocument,
    myArtboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
rotateArtboard(doc, myArtboard, 90);



/**
 * Rotates artboard by mutiples of 90°.
 * @author m1b
 * @param {Artboard} ab - an Illustrator Artboard
 * @param {Number} angle - 90°, 180°, 270°
 */
function rotateArtboard(doc, ab, angle) {
    if (
        ab == undefined
        || ab.constructor.name != 'Artboard'
    )
        return;

    if (
        angle == undefined
        || (angle % 90 != 0)
    ) {
        alert('Angle parameter must be multiple of 90 degrees.');
        return;
    }

    var items = itemsOnArtboard(doc, ab),
        center = centerOfBounds(ab.artboardRect);

    if (angle == 90 || angle == 270) {
        var r = ab.artboardRect;
        ab.artboardRect = boundsByCenterWidthHeight(center, -(r[3] - r[1]), r[2] - r[0]);
    }

    for (var i = 0; i < items.length; i++) {
        rotateItemAroundPoint(items[i], center, angle);
    }

}



/**
 * Get the center point of bounds.
 * @param {Array[Number]} bounds - [L, T, R, B]
 * @returns {Array[Number]} [x, y]
*/
function centerOfBounds(bounds) {
    var l = bounds[0], t = bounds[1], r = bounds[2], b = bounds[3],
        x = l + ((r - l) / 2),
        y = t + ((b - t) / 2);
    return [x, y];
};



/**
 * Rotate page item around a point [x, y].
 * @author m1b
 * @param {PageItem} item - an Illustrator page item
 * @param {Array[Number]} point - a point [x, y]
 * @param {Number} angleInDegrees - rotation angle
 */
function rotateItemAroundPoint(item, point, angleInDegrees) {
    // sanity
    if (
        item == undefined
        || typeof item.translate != 'function'
    )
        return;
    if (
        point == undefined
        || !point.hasOwnProperty('1')
    )
        return;
    if (
        angleInDegrees == undefined
    )
        return;

    // move 'point' to the document origin, then rotate, then move back
    item.translate(-point[0], -point[1]);
    item.rotate(angleInDegrees, true, false, false, false, Transformation.DOCUMENTORIGIN);
    item.translate(point[0], point[1]);
}



/**
 * Returns page items on an Illustrator artboard.
 * NOTE: will modify current selection.
 * @author m1b
 * @param {Artboard} ab - an Illustrator Artboard
 * @returns {Array}
 */
function itemsOnArtboard(doc, ab) {
    if (
        ab == undefined
        || ab.constructor.name != 'Artboard'
    )
        return [];

    var found = [],
        index = getArtboardIndexByName(doc, ab.name);

    if (index == -1)
        return [];

    doc.artboards.setActiveArtboardIndex(index);
    doc.selectObjectsOnActiveArtboard();

    for (var i = 0; i < selection.length; i++) {
        found.push(selection[i]);
    }

    return found;
}



/**
 * Returns the index of Named artboard.
 * @author m1b
 * @param {Document} doc - an Illustrator Document
 * @param {String} name - the Artboard name
 * @returns {Number}
 */
function getArtboardIndexByName(doc, name) {
    for (var i = 0; i < doc.artboards.length; i++) {
        if (doc.artboards[i].name == name)
            return i;
    }
    return -1;
}



/**
 * Calculates bounds of width x height centered on a point.
 * @author m1b
 * @param {point} center - a point [x, y]
 * @param {Number} width - width in points
 * @param {Number} height - height in points
 * @returns {Array} [L, T, R, B]
 */
function boundsByCenterWidthHeight(center, width, height) {
    var cx = center[0], cy = center[1];
    var l = cx - (width / 2);
    var t = cy + (height / 2);
    var r = cx + (width / 2);
    var b = cy - (height / 2);
    return [l, t, r, b];
}

};







function getTargetFile(docName, ext, destFolder) {
    

this.prefix	   = this.prefixEt.text; 
    
    var newName = tt+'pcs '+arthne+'x'+artwne+'_'+this.prefix;

    // if name has no dot (and hence no extension),
    // just append the extension
    if (docName.indexOf('.') < 0) {
        newName = docName + ext;
    } else {
        var dot = docName.lastIndexOf('.');
        newName = newName;
        newName += ext;
    }

    // Create the file object to save to
    var myFile = new File( destFolder + '/' + newName );

    // Preflight access rights
    if (myFile.open('w')) {
        myFile.close();
    }
    else {
        throw new Error('Access is denied');
    }
    return myFile;
}

function getTargetFile2(docName, ext, destFolder) {
    

this.prefix2   = this.prefixEt.text; 
    
    var newName2= 'GT_'+tt+'pcs '+arthne+'x'+artwne+'_'+this.prefix2;

    // if name has no dot (and hence no extension),
    // just append the extension
    if (docName.indexOf('.') < 0) {
        newName2= docName + ext;
    } else {
        var dot = docName.lastIndexOf('.');
        newName2= newName2
        newName2+= ext;
    }

    // Create the file object to save to
    var myFile = new File( destFolder + '/' + newName2);

    // Preflight access rights
    if (myFile.open('w')) {
        myFile.close();
    }
    else {
        throw new Error('Access is denied');
    }
    return myFile;
}
function getTargetFile3(docName, ext, destFolder) {

this.prefix3   = this.prefixEt.text; 
    
    var newName3= 'Saga_'+tt+'pcs '+arthne+'x'+artwne+'_'+this.prefix3;

    // if name has no dot (and hence no extension),
    // just append the extension
    if (docName.indexOf('.') < 0) {
        newName3= docName + ext;
    } else {
        var dot = docName.lastIndexOf('.');
        newName3= newName3
        newName3+= ext;
    }

    // Create the file object to save to
    var myFile = new File( destFolder + '/' + newName3);

    // Preflight access rights
    if (myFile.open('w')) {
        myFile.close();
    }
    else {
        throw new Error('Access is denied');
    }
    return myFile;
}

function zoom(){
    var zoomLevel =  '100';

if ( zoomLevel ) {
    zoomLevel = zoomLevel / 100;
    app.executeMenuCommand('fitin'); // Centers the artboard to the window
    app.activeDocument.activeView.zoom = zoomLevel;
}
}
