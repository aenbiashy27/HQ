#target photoshop
app.bringToFront();

main();


 // 13/12/2025 
function main(){
    if(!documents.length) {
        alert('No documents are open!');
        return;
    }
    
    // 修改对象属性名称，将 default 改为 primary
    var ripLocations = {
        primary: 'HP Latex 360 B/1 - Colorvizio PP poster',
        optionA: 'HP Latex 360 A/1 - Colorvizio PP poster',
        optionB: 'HP Latex 360 B/1 - Colorvizio PP poster'
    };
    var boardFN = "";
    var boardDC = "";
    var boardCP = "";
    var allFN = "";
    
    // 创建对话框
    var win = new Window('dialog', 'Poster Onyx Auto V9 - 12/03/2025');
    win.alignChildren = 'center';
    
    // ===== 第一面板：Color Mode =====
    var colPanel = win.add('panel', undefined, 'Color Mode');
    colPanel.orientation = 'row';
    colPanel.alignment = 'fill';
    colPanel.margins = [20, 20, 120, 20];
    var colGroup = colPanel.add('group');
    colGroup.orientation = 'row';
    colGroup.alignChildren = 'center';
    colGroup.graphics.font = 'Myriad Pro:18';
    win.cmyk = colGroup.add('radiobutton', undefined, 'CMYK');
    win.cmyk.value = true;
    win.rgb  = colGroup.add('radiobutton', undefined, 'RGB');
    win.rgb.value = false;
    
    // ===== 第二面板（用于海报选项）=====
    var ppPanel = win.add('panel', undefined, 'Color Mode');
    ppPanel.orientation = 'row';
    ppPanel.alignment = 'fill';
    ppPanel.margins = [20, 20, 120, 20];
    var ppGroup = ppPanel.add('group');
    ppGroup.orientation = 'row';
    ppGroup.alignChildren = 'center';
    ppGroup.graphics.font = 'Myriad Pro:18';
    
    var posterOptions = [
        'Matt Board', 'Matt Poster', 'Matt Sticker',
        'Matt pullup banner','Matt Xbanner',
        'Gloss Board', 'Gloss Poster', 'Gloss Sticker',
        'Gloss Pullup banner','Gloss Xbanner',
        'PVC','Fabric','Artist Canvas',
        'Backlit Film','Backlit Sticker',
        'Matt ORAJET','Gloss ORAJET',
        'Coated Paper'
    ];
    win.ppoptions = ppGroup.add('dropdownlist', undefined, posterOptions);
    win.ppoptions.selection = 0;
    win.ppoptions.onChange = function(){
        switch(this.selection.index){
            case 0:
                win.pptext.text = 'M  B';
                ripLocations.primary = 'HP Latex 360 B/1 - Colorvizio PP poster';
                ripLocations.optionA = 'HP Latex 360 A/1 - Colorvizio PP poster';
                ripLocations.optionB = 'HP Latex 360 B/1 - Colorvizio PP poster';
                showExtraOptions(true);
                break;
            case 1:
                win.pptext.text = 'M  P';
                ripLocations.primary = 'HP Latex 360 A/1 - Colorvizio PP poster';
                ripLocations.optionA = 'HP Latex 360 A/1 - Colorvizio PP poster';
                ripLocations.optionB = 'HP Latex 360 B/1 - Colorvizio PP poster';
                showExtraOptions(false);
                allFN = "";
                break;
            case 2:
                win.pptext.text = 'M  SSSSS';
                ripLocations.primary = 'HP Latex 360 B/1 - Colorvizio PP poster';
                ripLocations.optionA = 'HP Latex 360 A/1 - Colorvizio PP poster';
                ripLocations.optionB = 'HP Latex 360 B/1 - Colorvizio PP poster';
                showExtraOptions(false);
                allFN = "";
                break;
            case 3:
                win.pptext.text = 'M  Pullup';
                ripLocations.primary = 'HP Latex 360 A/1 - Colorvizio PP poster';
                ripLocations.optionA = 'HP Latex 360 A/1 - Colorvizio PP poster';
                ripLocations.optionB = 'HP Latex 360 B/1 - Colorvizio PP poster';
                showExtraOptions(false);
                allFN = "";
                break;
            case 4:
                win.pptext.text = 'M  Xbanner';
                ripLocations.primary = 'HP Latex 360 A/1 - Colorvizio PP poster';
                ripLocations.optionA = 'HP Latex 360 A/1 - Colorvizio PP poster';
                ripLocations.optionB = 'HP Latex 360 B/1 - Colorvizio PP poster';
                showExtraOptions(false);
                allFN = "";
                break;
            case 5:
                win.pptext.text = 'G  B';
                ripLocations.primary = 'HP Latex 360 B/1 - Colorvizio PP poster';
                ripLocations.optionA = 'HP Latex 360 A/1 - Colorvizio PP poster';
                ripLocations.optionB = 'HP Latex 360 B/1 - Colorvizio PP poster';
                showExtraOptions(true);
                allFN = "";
                break;     
            case 6:
                win.pptext.text = 'G  P';
                ripLocations.primary = 'HP Latex 360 A/1 - Colorvizio PP poster';
                ripLocations.optionA = 'HP Latex 360 A/1 - Colorvizio PP poster';
                ripLocations.optionB = 'HP Latex 360 B/1 - Colorvizio PP poster';
                showExtraOptions(false);
                allFN = "";
                break;     
            case 7:
                win.pptext.text = 'G  SSSSS';
                ripLocations.primary = 'HP Latex 360 B/1 - Colorvizio PP poster';
                ripLocations.optionA = 'HP Latex 360 A/1 - Colorvizio PP poster';
                ripLocations.optionB = 'HP Latex 360 B/1 - Colorvizio PP poster';
                showExtraOptions(false);
                allFN = "";
                break;     
            case 8:
                win.pptext.text = 'G  Pullup';
                ripLocations.primary = 'HP Latex 360 A/1 - Colorvizio PP poster';
                ripLocations.optionA = 'HP Latex 360 A/1 - Colorvizio PP poster';
                ripLocations.optionB = 'HP Latex 360 B/1 - Colorvizio PP poster';
                showExtraOptions(false);
                allFN = "";
                break;     
            case 9:
                win.pptext.text = 'G  Xbanner';
                ripLocations.primary = 'HP Latex 360 A/1 - Colorvizio PP poster';
                ripLocations.optionA = 'HP Latex 360 A/1 - Colorvizio PP poster';
                ripLocations.optionB = 'HP Latex 360 B/1 - Colorvizio PP poster';
                showExtraOptions(false);
                allFN = "";
                break;     
            case 10:
                win.pptext.text = 'PVC';
                ripLocations.primary = 'HP Latex 360 A/3 - PVC_BANNER';
                ripLocations.optionA = 'HP Latex 360 A/3 - PVC_BANNER';
                ripLocations.optionB = 'HP Latex 360 B/3 - PVC_BANNER';
                showExtraOptions(false);
                allFN = "";
                break;     
            case 11:
                win.pptext.text = 'Fabic';
                ripLocations.primary = 'HP Latex 360 A/9 - Fabric';
                ripLocations.optionA = 'HP Latex 360 A/9 - Fabric';
                ripLocations.optionB = 'HP Latex 360 B/9 - Fabric';
                showExtraOptions(false);
                allFN = "";
                break;     
            case 12:
                win.pptext.text = 'Artist Canvas';
                ripLocations.primary = 'HP Latex 360 B/8 - CANVAS';
                ripLocations.optionA = 'HP Latex 360 A/8 - CANVAS';
                ripLocations.optionB = 'HP Latex 360 B/8 - CANVAS';
                showExtraOptions(false);
                allFN = "";
                break;     
            case 13:
                win.pptext.text = 'BacklitFilm';
                ripLocations.primary = 'HP Latex 360 B/4 - PET BACKLIT FILM';
                ripLocations.optionA = 'HP Latex 360 A/4 - PET BACKLIT FILM';
                ripLocations.optionB = 'HP Latex 360 B/4 - PET BACKLIT FILM';
                showExtraOptions(false);
                allFN = "";
                break;
            case 14:
                win.pptext.text = 'BacklitSticker';
                ripLocations.primary = 'HP Latex 360 B/4 - PET BACKLIT FILM';
                ripLocations.optionA = 'HP Latex 360 A/4 - PET BACKLIT FILM';
                ripLocations.optionB = 'HP Latex 360 B/4 - PET BACKLIT FILM';
                showExtraOptions(false);
                allFN = "";
                break;          
            case 15:
                win.pptext.text = 'M  ORAJET';
                ripLocations.primary = 'HP Latex 360 B/2 - ORAJET';
                ripLocations.optionA = 'HP Latex 360 A/2 - ORAJET';
                ripLocations.optionB = 'HP Latex 360 B/2 - ORAJET';
                showExtraOptions(false);
                allFN = "";
                break;     
            case 16:
                win.pptext.text = 'G  ORAJET';
                ripLocations.primary = 'HP Latex 360 B/2 - ORAJET';
                ripLocations.optionA = 'HP Latex 360 A/2 - ORAJET';
                ripLocations.optionB = 'HP Latex 360 B/2 - ORAJET';
                showExtraOptions(false);
                allFN = "";
                break;     
            case 17:
                win.pptext.text = 'CoatedPaper';
                ripLocations.primary = 'HP Latex 360 B/6 - COATED PAPER';
                ripLocations.optionA = 'HP Latex 360 A/6 - COATED PAPER';
                ripLocations.optionB = 'HP Latex 360 B/6 - COATED PAPER';
                showExtraOptions(false);
                allFN = "";
                break;
        }
        allFN = boardFN + boardDC + boardCP;
    };
    
    function showExtraOptions(flag){
        if(flag){
            win.SFds.show();
            win.SFdc.show();
            win.SFcp.show();
        } else {
            win.SFds.hide();
            win.SFdc.hide();
            win.SFcp.hide();
        }
    }
    
    win.pptext = ppGroup.add('statictext', undefined, 'M B');
    win.pptext.characters = 11;
    
    // ===== 第三面板：Staff Name =====
    var stPanel = win.add('panel', undefined, 'Staff Name');
    stPanel.orientation = 'row';
    stPanel.alignment = 'fill';
    stPanel.margins = [20, 20, 20, 20];
    var stGroup = stPanel.add('group');
    stGroup.orientation = 'row';
    stGroup.alignChildren = 'center';
    stGroup.graphics.font = 'Myriad Pro:18';
    win.stname = stGroup.add('dropdownlist', undefined, ['BONG', 'ALEX','An','MING', 'GUEST1','GUEST2','GUEST3']);
    win.stname.selection = 0;
    
    win.SFds = stGroup.add('checkbox', undefined, 'Double Sided');
    win.SFds.value = false;
    win.SFds.onClick = function(){
        boardFN = win.SFds.value ? "DoubleSided_" : "";
        allFN = boardFN + boardDC + boardCP;
    };
    
    win.SFdc = stGroup.add('checkbox', undefined, 'Die-cut');
    win.SFdc.value = false;
    win.SFdc.onClick = function(){
        if(win.SFdc.value){
            win.SFcp.hide();
            boardDC = "DieCut_";
        } else {
            win.SFcp.show();
            boardDC = "";
        }
        allFN = boardFN + boardDC + boardCP;
    };
    
    win.SFcp = stGroup.add('dropdownlist', undefined, ['No Capping','Black Capping', 'White Capping', 'Gold Capping','Silver Capping']);
    win.SFcp.selection = 0;
    win.SFcp.onChange = function(){
        switch(this.selection.index){
            case 0:
                win.SFdc.show();
                boardCP = "";
                break;
            case 1:
                win.SFdc.hide();
                boardCP = "BlackCap_";
                break;
            case 2:
                win.SFdc.hide();
                boardCP = "WhiteCap_";
                break;
            case 3:
                win.SFdc.hide();
                boardCP = "GoldCap_";
                break;
            case 4:
                win.SFdc.hide();
                boardCP = "SilverCap_";
                break;
        }
        allFN = boardFN + boardDC + boardCP;
    };
    
    // ===== 第四面板：Printer Rip And Roate =====
    var printPanel = win.add('panel', undefined, 'Printer Rip And Roate');
    printPanel.orientation = 'row';
    printPanel.alignment = 'fill';
    printPanel.margins = [20, 20, 20, 20];
    var printGroup = printPanel.add('group');
    printGroup.orientation = 'row';
    printGroup.alignChildren = 'center';
    printGroup.graphics.font = 'Myriad Pro:18';
    
    var printerOptions = ['Auto Rip', 'Printer A','Printer B'];
    win.autoprinter = printGroup.add('dropdownlist', undefined, printerOptions);
    win.autoprinter.selection = 0;
    
    var rotateOptions = ['Auto Rotate', 'No Rotate'];
    win.autorotate = printGroup.add('dropdownlist', undefined, rotateOptions);
    win.autorotate.selection = 0;
    
    var dpiOptions = ['Auto DPI', 'File DPI'];
    win.autodpi = printGroup.add('dropdownlist', undefined, dpiOptions);
    win.autodpi.selection = 0;
    
    // ===== 第五面板：Auto Preset =====
    var presetPanel = win.add('panel', undefined, 'Auto Preset');
    presetPanel.orientation = 'row';
    presetPanel.alignment = 'fill';
    presetPanel.margins = [20, 20, 20, 20];
    var presetGroup = presetPanel.add('group');
    presetGroup.orientation = 'column';
    presetGroup.alignChildren = 'center';
    presetGroup.graphics.font = 'Myriad Pro:18';
    var presetGroup2 = presetPanel.add('group');
    presetGroup2.orientation = 'row';
    presetGroup2.alignChildren = 'left';
    presetGroup2.graphics.font = 'Myriad Pro:18';
    var presetGroup3 = presetPanel.add('group');
    presetGroup3.orientation = 'row';
    presetGroup3.alignChildren = 'left';
    presetGroup3.graphics.font = 'Myriad Pro:18';
    
    win.bits = presetGroup.add('radiobutton', undefined, '8 bits /channel');
    win.bits.value = true;
    win.JPG  = presetGroup2.add('radiobutton', undefined, 'JPG');
    win.JPG.value = true;
    
    var artPages = app.documents.length;
    win.Pagess = presetGroup2.add('statictext', undefined, '                   Proceesing Print: ' + artPages + ' pages');
    
    // 检查 Statistics.jsx 文件是否存在
    if(!checkStatistics()){
        return;
    }
    
    // ===== 最后一面板：OK / Cancel =====
    var actionPanel = win.add('panel', undefined, '');
    actionPanel.orientation = 'column';
    actionPanel.alignment = 'fill';
    actionPanel.margins = [20, 20, 20, 20];
    var actionGroup = actionPanel.add('group');
    actionGroup.orientation = 'row';
    actionGroup.alignChildren = 'center';
    actionGroup.graphics.font = 'Myriad Pro:18';
    
    win.ok = actionGroup.add('button', undefined, 'OK');
    win.cancel = actionGroup.add('button', undefined, 'Cancel');
    win.cancel.onClick = function(){ main.quit(); };
    
    win.ok.onClick = function(){
        app.displayDialogs = DialogModes.NO;
        win.close();
        processDocuments();
    };
    
    win.show();
    
    // =================== 功能模块 ===================
    
    function checkStatistics(){
        var f = new File('/Applications/Adobe Photoshop CS6/Legal/en_US/Statistics.jsx');
        if(!f.exists){
            alert('error code : sq script error 222');
            return false;
        }
        return true;
    }
    
    function processDocuments(){
        var originalDoc = app.activeDocument;
        for(var i=0; i<app.documents.length; i++){
             app.activeDocument = app.documents[i];
             var doc = app.activeDocument;
             doc.flatten();
             // 如果尺寸检查不通过，则中断处理并重新显示界面
             if(!rotateDocument()){
                 alert("Processing aborted due to size check failure. Returning to interface.");
                 main();
                 return;
             }
             adjustDPI();
             // 若海报选项为 case 0 或 5 且选择了 Capping（非 No Capping），则调用 adjustCapping()
             if((win.ppoptions.selection.index === 0 || win.ppoptions.selection.index === 5) && win.SFcp.selection.index > 0){
                 adjustCapping();
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
        saveDocuments();
        app.activeDocument = originalDoc;
    }
    
    // 修改后的 rotateDocument 函数，返回 true 表示成功，false 表示尺寸检查未通过
    function rotateDocument(){
        var success = true;
        if(win.autorotate.selection.index === 0){
            var w = app.activeDocument.width.as('mm');
            var h = app.activeDocument.height.as('mm');
            var artw = Math.round(w*10)/10;
            var arth = Math.round(h*10)/10;
            
            if(artw > 1220 && arth < 1220){
                app.activeDocument.rotateCanvas(90);
            } else if(artw > 1220 && arth > 1220){
                alert('Large size?, Size: ' + artw + ' x ' + arth + ' mm');
            } else {
                var qtyP = 1220/artw;
                var qtyL = 1220/arth;
                var qtyPmr = Math.floor(qtyP);
                var qtyLmr = Math.floor(qtyL);
                var sqfP = ((1200 * arth) / 93025) / qtyPmr;
                var sqfL = ((1200 * artw) / 93025) / qtyLmr;
                if(sqfP > sqfL){
                    app.activeDocument.rotateCanvas(90);
                }
            }
            // 根据海报选项下拉列表的选择项进行额外操作
            switch(win.ppoptions.selection.index){
                case 3:
                case 8:
                    // Function 2：要求尺寸必须是600x1650 或850x2050 或1000x2050 或1200x2050（允许±5mm误差）
                    if(!checkCase2()){
                        success = false;
                    } else {
                        app.activeDocument.rotateCanvas(180);
                    }
                    break;
                case 4:
                case 9:
                    // Function 1：要求尺寸必须是600x1600（允许±5mm误差）
                    if(!checkCase1()){
                        success = false;
                    }
                    break;
                default:
                    break;
            }
        }
        return success;
    }
    
    // Function 1：用于 case 4 和 9
    // 检查当前文档尺寸是否在600×1600mm（允许±5mm误差）；如果符合则扩展画布至600×1670mm（背景白色）
    function checkCase1(){
        var doc = app.activeDocument;
        var w = doc.width.as('mm');
        var h = doc.height.as('mm');
        if(Math.abs(w - 600) > 5 || Math.abs(h - 1600) > 5){
             alert("size not in 600 x 1600mm, please resize");
             return false;
        } else {
             var whiteColor = new SolidColor();
             whiteColor.rgb.red = 255;
             whiteColor.rgb.green = 255;
             whiteColor.rgb.blue = 255;
             app.backgroundColor = whiteColor;
             try{
                 doc.resizeCanvas(UnitValue(600, "mm"), UnitValue(1670, "mm"), AnchorPosition.MIDDLECENTER);
             } catch(e){
                 alert("Error resizing canvas: " + e);
                 return false;
             }
             return true;
        }
    }
    
    // Function 2：用于 case 3 和 8
    // 检查当前文档尺寸是否为以下之一（允许±5mm误差）：600×1650、850×2050、1000×2050 或1200×2050
    function checkCase2(){
        var doc = app.activeDocument;
        var w = doc.width.as('mm');
        var h = doc.height.as('mm');
        var validSizes = [
           {w:600, h:1650},
           {w:850, h:2050},
           {w:1000, h:2050},
           {w:1200, h:2050}
        ];
        var found = false;
        var closestSize = null;
        var minDiff = Infinity;
        for(var i=0; i<validSizes.length; i++){
             var diff = Math.abs(w - validSizes[i].w) + Math.abs(h - validSizes[i].h);
             if(diff < minDiff){
                 minDiff = diff;
                 closestSize = validSizes[i];
             }
             if(Math.abs(w - validSizes[i].w) <= 5 && Math.abs(h - validSizes[i].h) <= 5){
                 found = true;
                 break;
             }
        }
        if(!found){
             alert("size not in " + closestSize.w + " x " + closestSize.h + "mm, please resize");
             return false;
        }
        return true;
    }
    
    function adjustDPI(){
        if(win.autodpi.selection.index === 0){
            var w = app.activeDocument.width.as('mm');
            var h = app.activeDocument.height.as('mm');
            var artw = Math.round(w*10)/10;
            var arth = Math.round(h*10)/10;
            var sqf = (artw * arth) / 93025;
            var currRes = app.activeDocument.resolution;
            if(sqf < 2 && currRes > 300){
                app.activeDocument.resizeImage(undefined, undefined, 300, ResampleMethod.BICUBIC);
            } else if(sqf >= 2 && sqf < 20 && currRes > 200){
                app.activeDocument.resizeImage(undefined, undefined, 200, ResampleMethod.BICUBIC);
            } else if(sqf >= 20 && currRes > 150){
                app.activeDocument.resizeImage(undefined, undefined, 150, ResampleMethod.BICUBIC);
            }
        }
    }
    
    // 新增的 Function：用于 case 0 和 case 5，当选择了 Capping（非 No Capping）时，
    // 将当前文档的图像尺寸缩小 4mm（宽、高各减少4mm），再扩展 canvas 至原始尺寸（背景白色）
    function adjustCapping(){
        var doc = app.activeDocument;
        var origWidth = doc.width.as("mm");
        var origHeight = doc.height.as("mm");
        // 设置背景为白色
        var whiteColor = new SolidColor();
        whiteColor.rgb.red = 255;
        whiteColor.rgb.green = 255;
        whiteColor.rgb.blue = 255;
        app.backgroundColor = whiteColor;
        // 缩小图像尺寸 4mm（注意：此处使用 resizeImage 会更改文档尺寸）
        doc.resizeImage(UnitValue(origWidth - 4, "mm"), UnitValue(origHeight - 4, "mm"), doc.resolution, ResampleMethod.BICUBIC);
        // 扩展 canvas 至原始尺寸，扩展区域以中间为基准，背景为白色
        doc.resizeCanvas(UnitValue(origWidth, "mm"), UnitValue(origHeight, "mm"), AnchorPosition.MIDDLECENTER);
    }
    
    function saveDocuments(){
        var originalDoc = app.activeDocument;
        var saveDir;
        if(win.autoprinter.selection.index === 0){
            saveDir = ripLocations.primary;
        } else if(win.autoprinter.selection.index === 1){
            saveDir = ripLocations.optionA;
        } else if(win.autoprinter.selection.index === 2){
            saveDir = ripLocations.optionB;
        }
        var docPath = '~/Desktop/' + saveDir;
        var backupPath = '~/Desktop/1_Master Poster';
        
        var jpgOpts = new JPEGSaveOptions();
        jpgOpts.embedColorProfile = true;
        jpgOpts.formatOptions = FormatOptions.STANDARDBASELINE;
        jpgOpts.matte = MatteType.NONE;
        jpgOpts.quality = 12;
        
        for(var m=0; m<app.documents.length; m++){
            var today = new Date();
            var dateStr = today.getDate() + '' + (today.getMonth()+1);
            var timeStr = today.getHours() + '' + today.getMinutes() + '' + today.getSeconds();
            var dateTime = dateStr + timeStr;
            var theDoc = app.documents[m];
            var num = m+1;
            var fileName = win.pptext.text + '_' + allFN + win.stname.selection.text + '_' + num + '.jpg';
            var savedFile = new File(docPath + '/' + fileName);
            var savedFileBackup = new File(backupPath + '/' + win.pptext.text + '_' + allFN + win.stname.selection.text + '_' + num + '_' + dateTime + '.jpg');
            app.activeDocument = theDoc;
            if(!savedFile.exists){
                theDoc.saveAs(savedFile, jpgOpts, false, Extension.LOWERCASE);
                theDoc.saveAs(savedFileBackup, jpgOpts, false, Extension.LOWERCASE);
            } else {
                var result = confirm(savedFile.fsName + ' already exists. Do you want to replace it?', false);
                if(result){
                    theDoc.saveAs(savedFile, jpgOpts, false, Extension.LOWERCASE);
                    theDoc.saveAs(savedFileBackup, jpgOpts, false, Extension.LOWERCASE);
                } else {
                    main();
                    return;
                }
            }
        }
        alert('Done | 存档完毕');
        app.activeDocument = originalDoc;
    }
}
