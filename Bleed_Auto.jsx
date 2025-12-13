#target photoshop

 // 13/12/2025 


function main(){
    app.bringToFront();
    app.preferences.rulerUnits = Units.MM;
    app.preferences.typeUnits = TypeUnits.MM;

    if (!documents.length) {
        alert('No documents are open!');
        return;
    }
    
    // 内容识别填充函数
    function contentAwareFill() {
        try {
            activeDocument.selection.bounds;
        } catch (e) {
            alert('Error: Artwork may already include bleed size');
            return;
        }
        var desc = new ActionDescriptor();
        desc.putEnumerated(charIDToTypeID('Usng'), charIDToTypeID('FlCn'), stringIDToTypeID('contentAware'));
        executeAction(charIDToTypeID('Fl  '), desc, DialogModes.NO);
    }
    
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


    // 单个文档的处理函数（先 flatten，再加出血，最后（可选）添加切线）
    function processDocument(doc, bleedW_input, bleedH_input, cutLineEnabled) {
        app.activeDocument = doc;
        doc.flatten();
        if (doc.bitsPerChannel != BitsPerChannelType.EIGHT) {
            doc.bitsPerChannel = BitsPerChannelType.EIGHT;
        }
        // 移除所有参考线
        for (var i = doc.guides.length - 1; i >= 0; i--) {
            doc.guides[i].remove();
        }
        
        var docWidth = Math.round(doc.width);
        var docHeight = Math.round(doc.height);
        var newWidth = docWidth + Number(bleedW_input);
        var newHeight = docHeight + Number(bleedH_input);
        var bleedLeftTop = ((newWidth - docWidth) / 2) + 0.5;
        var bleedTop = ((newHeight - docHeight) / 2) + 0.5;
        var bleedRight = newWidth - bleedLeftTop - 0.1;
        var bleedBottom = newHeight - bleedTop - 0.1;
        
        // 修改画布尺寸，锚点设为中间
        doc.resizeCanvas(new UnitValue(newWidth, 'mm'), new UnitValue(newHeight, 'mm'), AnchorPosition.MIDDLECENTER);
        
        // 计算毫米与像素的换算比例
        app.preferences.rulerUnits = Units.MM;
        var mmWidth = Number(doc.width);
        app.preferences.rulerUnits = Units.PIXELS;
        var pxWidth = Number(doc.width);
        var cvmm = pxWidth / mmWidth;
        
        // 定义四个出血区域（上、左、下、右）
        var regionTop = [ [0, 0],
                          [0, bleedTop * cvmm],
                          [newWidth * cvmm, bleedTop * cvmm],
                          [newWidth * cvmm, 0] ];
        var regionLeft = [ [0, 0],
                           [0, newHeight * cvmm],
                           [bleedLeftTop * cvmm, newHeight * cvmm],
                           [bleedLeftTop * cvmm, 0] ];
        var regionBottom = [ [0, bleedBottom * cvmm],
                             [0, newHeight * cvmm],
                             [newWidth * cvmm, newHeight * cvmm],
                             [newWidth * cvmm, bleedBottom * cvmm] ];
        var regionRight = [ [bleedRight * cvmm, 0],
                            [newWidth * cvmm, 0],
                            [newWidth * cvmm, newHeight * cvmm],
                            [bleedRight * cvmm, newHeight * cvmm] ];
        
        var regions = [regionTop, regionLeft, regionBottom, regionRight];
        for (var i = 0; i < regions.length; i++) {
            doc.selection.select(regions[i], SelectionType.REPLACE, 0, false);
            contentAwareFill();
        }
        
        // 如果选中“Cut Line”，则添加切线（描边）
        if (cutLineEnabled) {
            app.preferences.rulerUnits = Units.PIXELS;
            var strokeColor = new SolidColor();
            strokeColor.cmyk.cyan = 0;
            strokeColor.cmyk.magenta = 0;
            strokeColor.cmyk.yellow = 0;
            strokeColor.cmyk.black = 100;
            doc.selection.selectAll();
            doc.selection.stroke(strokeColor, 1, StrokeLocation.INSIDE);
            doc.selection.deselect();
            app.preferences.rulerUnits = Units.MM;
        }
        app.preferences.rulerUnits = Units.MM;
    }
    
    // 对话框界面
    function showDialog() {
        var win = new Window('dialog', 'Bleed Auto V9.2 -31/01/2025');
        win.alignChildren = 'center';

        // Canvas Size 面板
        var ffPanel = win.add('panel', undefined, 'Canvas Size');
        ffPanel.orientation = 'row';
        ffPanel.margins = [200, 20, 200, 20];
        var stGroup2 = ffPanel.add('group');
        stGroup2.orientation = 'row';
        stGroup2.alignChildren = 'center';
        stGroup2.graphics.font = 'Myriad Pro:20';
        
        var docWidthVal = Math.round(app.activeDocument.width);
        var docHeightVal = Math.round(app.activeDocument.height);
        
        var canvasW = stGroup2.add('edittext', undefined, docWidthVal.toString());
        canvasW.characters = 6;
        canvasW.onChange = function() {
            bleedW.text = Number(canvasW.text) - docWidthVal;
        };
        stGroup2.add('statictext', undefined, 'mm (w) ');
        
        var canvasH = stGroup2.add('edittext', undefined, docHeightVal.toString());
        canvasH.characters = 6;
        canvasH.onChange = function() {
            bleedH.text = Number(canvasH.text) - docHeightVal;
        };
        stGroup2.add('statictext', undefined, 'mm (h) ');
        
        // Bleed 面板
        var ppPanel = win.add('panel', undefined, 'Bleed');
        ppPanel.orientation = 'row';
        ppPanel.margins = [200, 20, 200, 20];
        var ppGroup = ppPanel.add('group');
        ppGroup.orientation = 'row';
        ppGroup.alignChildren = 'center';
        ppGroup.graphics.font = 'Myriad Pro:20';
        
        ppGroup.add('statictext', undefined, '+  ');
        var bleedW = ppGroup.add('edittext', undefined, '0');
        bleedW.characters = 6;
        bleedW.onChange = function() {
            canvasW.text = docWidthVal + Number(bleedW.text);
        };
        ppGroup.add('statictext', undefined, 'mm (w) ');
        
        var bleedH = ppGroup.add('edittext', undefined, '0');
        bleedH.characters = 6;
        bleedH.onChange = function() {
            canvasH.text = docHeightVal + Number(bleedH.text);
        };
        ppGroup.add('statictext', undefined, 'mm (h) ');
        
        // Setting 面板
        var settingPanel = win.add('panel', undefined, 'Setting');
        settingPanel.orientation = 'row';
        var settingGroup = settingPanel.add('group');
        settingGroup.orientation = 'row';
        settingGroup.alignChildren = 'left';
        settingGroup.graphics.font = 'Myriad Pro:22';
        
        var chkBleedPreset = settingGroup.add('checkbox', undefined, ' 5mm Bleed');
        chkBleedPreset.value = false;
        chkBleedPreset.onClick = function() {
            if (chkBleedPreset.value) {
                bleedW.text = '5';
                bleedH.text = '5';
            } else {
                bleedW.text = '0';
                bleedH.text = '0';
            }
        };
        
        var chkCutLine = settingGroup.add('checkbox', undefined, ' Cut Line');
        chkCutLine.value = false;
        
        // Auto Preset 面板
        var presetPanel = win.add('panel', undefined, 'Auto Preset');
        presetPanel.orientation = 'row';
        presetPanel.margins = [20, 20, 20, 20];
        var presetGroup = presetPanel.add('group');
        presetGroup.orientation = 'column';
        presetGroup.alignChildren = 'center';
        presetGroup.graphics.font = 'Myriad Pro:18';
 
        
        var pagesInfo = presetPanel.add('group');
        pagesInfo.orientation = 'row';
        pagesInfo.graphics.font = 'Myriad Pro:18';
        pagesInfo.add('statictext', undefined, 'Processing Resize: ' + app.documents.length + ' pages');
     
        
        // 底部按钮面板
        var btnPanel = win.add('panel', undefined, '');
        btnPanel.orientation = 'column';
        btnPanel.alignment = 'fill';
        btnPanel.margins = [20, 20, 20, 20];
        var btnGroup = btnPanel.add('group');
        btnGroup.orientation = 'row';
        btnGroup.alignChildren = 'center';
        btnGroup.graphics.font = 'Myriad Pro:18';
        
        var btnOk = btnGroup.add('button', undefined, 'Ok');
        btnOk.onClick = function() {
            win.close();
            var docs = [];
            for (var i = 0; i < app.documents.length; i++) {
                docs.push(app.documents[i]);
            }
            for (var j = 0; j < docs.length; j++) {
                processDocument(docs[j], bleedW.text, bleedH.text, chkCutLine.value);
            }
            alert('Bleed Added');
        };
        
        var btnCancel = btnGroup.add('button', undefined, 'Cancel');
        btnCancel.onClick = function() { win.close(); };
        
        win.show();
    }
    
    showDialog();
}

main();
