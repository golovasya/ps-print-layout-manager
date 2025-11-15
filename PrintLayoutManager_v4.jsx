#target photoshop

// ======================================================================
// Print Layout Manager v4.0 - Enhanced with Excel Integration
// Based on Sborshchik v.3.5 with added Excel XLSX support
// ======================================================================

app.preferences.rulerUnits = Units.MM;

// ================== ВАЖНО: Включить библиотеку SheetJS ==================
// Скачайте файлы с https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/
// и поместите их в ту же папку, что и этот скрипт:
// - shim.min.js
// - jszip.js  
// - xlsx.full.min.js
// 
// Раскомментируйте строки ниже после загрузки библиотек:
// #include "shim.min.js"
// #include "jszip.js"
// #include "xlsx.full.min.js"

// ================== Глобальные переменные для Excel ===================
var excelData = [];          // Данные из таблицы
var excelFilePath = null;    // Путь к загруженному файлу
var useExcelMode = false;    // Режим работы с Excel

// ================== ФУНКЦИИ ДЛЯ РАБОТЫ С EXCEL ========================

// Загрузка Excel файла
function loadExcelFile() {
    var file = File.openDialog("Выберите файл XLSX", "*.xlsx;*.xls");
    if (!file) return false;
    
    try {
        file.open("r");
        file.encoding = "binary";
        var data = file.read();
        file.close();
        
        var workbook = XLSX.read(data, {type: "binary"});
        var firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        var jsonData = XLSX.utils.sheet_to_json(firstSheet, {header: 1});
        
        excelData = [];
        for (var i = 1; i < jsonData.length; i++) {
            var row = jsonData[i];
            if (!row || row.length < 6) continue;
            
            var item = {
                photo: row[0] || "",
                size: row[1] || "",
                orderId: row[2] || "",
                name: row[3] || "",
                color: row[4] || "",
                article: row[5] || "",
                layerId: null,
                realSize: null
            };
            
            excelData.push(item);
        }
        
        excelFilePath = file.fsName;
        alert("Загружено " + excelData.length + " записей из таблицы");
        return true;
        
    } catch (e) {
        alert("Ошибка чтения Excel: " + e.message + "\n\nУбедитесь, что библиотеки SheetJS подключены!");
        return false;
    }
}

// Привязка слоя к данным таблицы
function attachExcelDataToLayer(layer, excelItem, itemIndex) {
    if (!layer || !excelItem) return;
    
    try {
        var safeName = (excelItem.article || "item") + "_" + (excelItem.size || "") + "_" + itemIndex;
        layer.name = safeName;
        excelItem.layerId = layer.id;
    } catch (e) {}
}

// Синхронизация размера слоя
function syncLayerSize(layer, newSizeMM) {
    if (!layer) return false;
    
    try {
        var doc = app.activeDocument;
        var bounds = layer.bounds;
        var currentWidth = bounds[2].as("mm") - bounds[0].as("mm");
        var currentHeight = bounds[3].as("mm") - bounds[1].as("mm");
        
        var maxCurrentSize = Math.max(currentWidth, currentHeight);
        var targetSizePx = (newSizeMM / 25.4) * doc.resolution;
        var scaleFactor = (targetSizePx / maxCurrentSize);
        
        layer.resize(scaleFactor * 100, scaleFactor * 100, AnchorPosition.TOPLEFT);
        return true;
    } catch (e) {
        alert("Ошибка изменения размера: " + e.message);
        return false;
    }
}

// Получение всех слоёв
function getAllLayers(doc) {
    var layers = [];
    function traverse(layerSet) {
        for (var i = 0; i < layerSet.layers.length; i++) {
            var layer = layerSet.layers[i];
            if (layer.typename === "LayerSet") {
                traverse(layer);
            } else {
                layers.push(layer);
            }
        }
    }
    traverse(doc);
    return layers;
}

// UI панель управления принтами
function showPrintManagerPanel() {
    if (!app.documents.length) {
        alert("Откройте документ с раскладкой");
        return;
    }
    
    if (!excelData || excelData.length === 0) {
        alert("Сначала загрузите таблицу Excel");
        return;
    }
    
    var win = new Window("palette", "Print Manager", undefined, {closeButton: true});
    win.orientation = "column";
    win.alignChildren = ["fill", "top"];
    win.preferredSize = [400, 600];
    
    var infoGroup = win.add("group");
    infoGroup.add("statictext", undefined, "Записей: " + excelData.length);
    
    var listGroup = win.add("group");
    listGroup.orientation = "column";
    listGroup.alignChildren = ["fill", "top"];
    listGroup.add("statictext", undefined, "Список принтов:");
    
    var printList = listGroup.add("listbox", undefined, [], {
        numberOfColumns: 4,
        showHeaders: true,
        columnTitles: ["Артикул", "Размер", "Цвет", "Физ.размер"]
    });
    printList.preferredSize = [380, 350];
    
    function updatePrintList() {
        printList.removeAll();
        for (var i = 0; i < excelData.length; i++) {
            var item = excelData[i];
            var sizeStr = item.realSize ? item.realSize + " мм" : "—";
            var listItem = printList.add("item", item.article || "—");
            listItem.subItems[0].text = item.size || "—";
            listItem.subItems[1].text = item.color || "—";
            listItem.subItems[2].text = sizeStr;
        }
    }
    
    updatePrintList();
    
    var detailsGroup = win.add("panel", undefined, "Детали");
    detailsGroup.orientation = "column";
    detailsGroup.alignChildren = ["fill", "top"];
    detailsGroup.margins = 10;
    
    var detailArticle = detailsGroup.add("statictext", undefined, "Артикул: —");
    var detailSize = detailsGroup.add("statictext", undefined, "Размер: —");
    var detailColor = detailsGroup.add("statictext", undefined, "Цвет: —");
    
    var btnGroup = win.add("group");
    btnGroup.orientation = "row";
    
    var btnSelect = btnGroup.add("button", undefined, "Выделить");
    var btnResize = btnGroup.add("button", undefined, "Изменить размер");
    
    printList.onChange = function() {
        if (printList.selection) {
            var idx = printList.selection.index;
            var item = excelData[idx];
            detailArticle.text = "Артикул: " + (item.article || "—");
            detailSize.text = "Размер: " + (item.size || "—");
            detailColor.text = "Цвет: " + (item.color || "—");
        }
    };
    
    btnSelect.onClick = function() {
        if (!printList.selection) {
            alert("Выберите принт");
            return;
        }
        
        var idx = printList.selection.index;
        var item = excelData[idx];
        
        if (!item.layerId) {
            alert("Принт не размещен");
            return;
        }
        
        try {
            var doc = app.activeDocument;
            var layers = getAllLayers(doc);
            
            for (var i = 0; i < layers.length; i++) {
                if (layers[i].id === item.layerId) {
                    doc.activeLayer = layers[i];
                    alert("Выделено!");
                    return;
                }
            }
            alert("Слой не найден");
        } catch (e) {
            alert("Ошибка: " + e.message);
        }
    };
    
    btnResize.onClick = function() {
        if (!printList.selection) {
            alert("Выберите принт");
            return;
        }
        
        var idx = printList.selection.index;
        var item = excelData[idx];
        
        if (!item.layerId) {
            alert("Принт не размещен");
            return;
        }
        
        var newSize = prompt("Новый размер (мм):", item.realSize || "300");
        if (!newSize) return;
        
        var sizeMM = parseFloat(newSize);
        if (isNaN(sizeMM) || sizeMM <= 0) {
            alert("Некорректный размер");
            return;
        }
        
        try {
            var doc = app.activeDocument;
            var layers = getAllLayers(doc);
            
            for (var i = 0; i < layers.length; i++) {
                if (layers[i].id === item.layerId) {
                    if (syncLayerSize(layers[i], sizeMM)) {
                        item.realSize = sizeMM;
                        updatePrintList();
                        alert("Размер изменен!");
                    }
                    return;
                }
            }
            alert("Слой не найден");
        } catch (e) {
            alert("Ошибка: " + e.message);
        }
    };
    
    win.show();
}

// Главный диалог
function showMainDialog() {
    var dlg = new Window("dialog", "Print Layout Manager v4.0");
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];
    dlg.margins = [10, 15, 10, 10];
    
    var excelPanel = dlg.add("panel", undefined, "Работа с таблицей Excel");
    excelPanel.orientation = "column";
    excelPanel.alignChildren = ["fill", "top"];
    excelPanel.margins = 10;
    
    var excelStatus = excelPanel.add("statictext", undefined, "Таблица не загружена");
    excelStatus.characters = 35;
    
    var btnLoadExcel = excelPanel.add("button", undefined, "Загрузить XLSX");
    var btnManagePrints = excelPanel.add("button", undefined, "Управление принтами");
    btnManagePrints.enabled = false;
    
    btnLoadExcel.onClick = function() {
        if (loadExcelFile()) {
            excelStatus.text = "Загружено: " + excelData.length + " записей";
            btnManagePrints.enabled = true;
            useExcelMode = true;
        }
    };
    
    btnManagePrints.onClick = function() {
        showPrintManagerPanel();
    };
    
    dlg.add("panel");
    
    var btnClose = dlg.add("button", undefined, "Закрыть");
    btnClose.onClick = function() {
        dlg.close();
    };
    
    dlg.show();
}

// Запуск
showMainDialog();