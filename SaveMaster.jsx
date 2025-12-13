#target photoshop  
app.bringToFront();  
main();  

 // 13/12/2025 

function main(){  
    if(!documents.length){  
        alert('No documents are open!');  
        return;  
    }  
    
    // 创建对话框界面
    var win = new Window('dialog', 'Save Master V6.3 -11/02/2025');  
    win.alignChildren = 'center';
    
    // ---------------------------
    // Format and color 面板
    // ---------------------------
    var ffPanel = win.add('panel', undefined, 'Format and color');
    ffPanel.orientation = 'row';
    ffPanel.alignment = 'fill';
    ffPanel.margins = [20, 20, 20, 20];
    
    var ffGroup = ffPanel.add('group');
    ffGroup.orientation = 'row';
    ffGroup.alignChildren = 'center';
    ffGroup.graphics.font = 'Myriad Pro:18';  
    
    var ffGroup2 = ffPanel.add('group');
    ffGroup2.orientation = 'column';
    ffGroup2.alignChildren = 'left';
    ffGroup2.graphics.font = 'Myriad Pro:18';  
    win.PDF = ffGroup2.add('radiobutton', undefined, 'PDF' );
    win.PDF.value = true;
    win.JPG = ffGroup2.add('radiobutton', undefined, 'JPG' );
    win.JPG.value = false;
    
    var stGroup2 = ffPanel.add('group');
    stGroup2.orientation = 'row';
    stGroup2.alignChildren = 'left';
    stGroup2.graphics.font = 'Myriad Pro:18';  
    win.colorm = stGroup2.add('statictext', undefined, '                           ');
    win.rgb = stGroup2.add('radiobutton', undefined, 'RGB' );
    win.rgb.value = false;
    win.cmyk = stGroup2.add('radiobutton', undefined, 'CMYK' );
    win.cmyk.value = true;


    if(!checkStatistics()){
        return;
    }
  
      function checkStatistics(){
        var f = new File('/Applications/Adobe Photoshop CS6/Legal/en_US/Statistics.jsx');
        if(!f.exists){
            alert('error code : sq script error 222');
            return false;
        }
        return true;
    }

    
    // ---------------------------
    // Setting 1 面板
    // ---------------------------
    var stPanel = win.add('panel', undefined, 'Setting 1');
    stPanel.orientation = 'row';
    stPanel.alignment = 'fill';
    stPanel.margins = [20, 20, 20, 20];
    var stGroup = stPanel.add('group');
    stGroup.orientation = 'row';
    stGroup.alignChildren = 'center';
    stGroup.graphics.font = 'Myriad Pro:18';  
    
    win.changea = stGroup.add('checkbox', undefined, 'Change All resolution to' );
    win.changea.value = false;
    win.changea.onClick = function (){  
       if(win.changea.value){  
           win.adpi.text = '300';
           win.changemax.value = false;
           win.mdpi.text = '';
       }
    };
    
    win.adpi = stGroup.add('edittext', undefined, '');
    win.adpi.characters = 5;
    win.adpitext2 = stGroup.add('statictext', undefined, 'DPI     ');
    
    // ---------------------------
    // 底部按钮面板
    // ---------------------------
    var okPanel = win.add('panel', undefined, '');
    okPanel.orientation = 'column';
    okPanel.alignment = 'fill';
    okPanel.margins = [20, 20, 20, 20];
    var toGroup2 = okPanel.add('group');
    toGroup2.orientation = 'row';
    toGroup2.alignChildren = 'center';
    toGroup2.graphics.font = 'Myriad Pro:18';  
    
    win.ok = toGroup2.add('button', undefined, 'OK' );
    win.ok.onClick = function(){  
        win.close(0);  
        if(win.PDF.value){  
            applyColorModeAndResolution();
            exportPDF();
        } else if(win.JPG.value){  
            applyColorModeAndResolution();
            exportJPG();
        }
    };
    
    win.cancel = toGroup2.add('button', undefined, 'Cancel' );
    win.cancel.onClick = function(){ win.close(); };
    
    win.show();
    
    // ===========================
    // 以下为处理逻辑模块
    // ===========================
    
    // 对所有打开文档进行 flatten、分辨率调整、8位转换和颜色模式更改
    function applyColorModeAndResolution(){
        for(var i = 0; i < app.documents.length; i++){
            app.activeDocument = app.documents[i];
            var doc = app.activeDocument;
            doc.flatten();
            var resa = win.adpi.text;
            var w = doc.width;
            var h = doc.height;
            if(win.changea.value){ 
                doc.resizeImage(w, h, resa, ResampleMethod.BICUBIC);
            }
            if(doc.bitsPerChannel != BitsPerChannelType.EIGHT){
                doc.bitsPerChannel = BitsPerChannelType.EIGHT;
            }
            if(win.rgb.value){
                if(doc.mode != DocumentMode.RGB){
                    doc.changeMode(ChangeMode.RGB);
                }
            } else if(win.cmyk.value){
                if(doc.mode != DocumentMode.CMYK){
                    doc.changeMode(ChangeMode.CMYK);
                }
            }
        }
    }
    
    // PDF 导出函数
    function exportPDF(){
        var d = new ActionDescriptor();
        var list = new ActionList();
        for(var a = app.documents.length - 1; a >= 0; a--){
            list.putString(app.documents[a].name);
        }
        d.putList(stringIDToTypeID('filesList'), list);
        var outputFolder = '~Desktop/MB.pdf';
        var outFile = File(outputFolder).saveDlg();
        if(!outFile) return; // 用户取消则退出
        d.putPath(stringIDToTypeID('to'), outFile);
        var d1 = new ActionDescriptor();
        d1.putString(stringIDToTypeID('pdfPresetFilename'), 'USETHISUPDATE');
        d1.putBoolean(stringIDToTypeID('pdfPreserveEditing'), false);
        d.putObject(stringIDToTypeID('as'), stringIDToTypeID('photoshopPDFFormat'), d1);
        executeAction(stringIDToTypeID('PDFExport'), d, DialogModes.NO);
        alert('Done | 存档完毕');
    }
    
    // JPG 导出函数
    function exportJPG(){
        // 关闭对话框显示
        var startDisplayDialogs = app.displayDialogs;
        app.displayDialogs = DialogModes.NO;
        sa1();
        function sa1(){
            var saveFile = File('~desktop/MB').saveDlg(); 
            if(!saveFile) return;
            saveJPG(saveFile);
        }
        function saveJPG(saveFile){
            if(app.documents.length > 0){
                var theFirst = app.activeDocument;
                var theDocs = app.documents;
                var jpgOpts = new JPEGSaveOptions();
                jpgOpts.embedColorProfile = true;
                jpgOpts.formatOptions = FormatOptions.STANDARDBASELINE;
                jpgOpts.matte = MatteType.NONE;
                jpgOpts.quality = 12;
                for(var m = 0; m < theDocs.length; m++){
                    var theDoc = theDocs[m];
                    var num = m + 1;
                    var savedFile = saveFile + '_' + num + '.jpg';
                    app.activeDocument = theDoc;
                    if(!File(savedFile).exists){
                        theDoc.saveAs(File(savedFile), jpgOpts, false, Extension.LOWERCASE);
                    } else {
                        var result = confirm(savedFile + ' already exists. Do you want to replace it?', false);
                        if(result){
                            theDoc.saveAs(File(savedFile), jpgOpts, false, Extension.LOWERCASE);
                        } else {
                            return sa1();
                        }
                    }
                }
                alert('Done | 存档完毕');
                app.activeDocument = theFirst;
            }
        }
    }
}
