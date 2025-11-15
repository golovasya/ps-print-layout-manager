const { storage, core } = require('uxp');
const { app, action } = require('photoshop');

let excelData = [];
let layerPrintMap = {};
let selectedPrintIndex = null;

document.addEventListener('DOMContentLoaded', init);

function init() {
    document.getElementById('loadExcelBtn').addEventListener('click', loadExcelFile);
    document.getElementById('runLayoutBtn').addEventListener('click', runLayoutScript);
    setupLayerSelectionSync();
}

async function loadExcelFile() {
    try {
        const fs = storage.localFileSystem;
        const file = await fs.getFileForOpening({ types: ['xlsx', 'xls'] });
        if (!file) return;
        
        const arrayBuffer = await file.read({ format: storage.formats.binary });
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
        
        excelData = parseExcelData(data);
        updateStatus(`Загружено ${excelData.length} записей`);
        document.getElementById('runLayoutBtn').disabled = false;
        renderPrintsList();
    } catch (error) {
        updateStatus('Ошибка: ' + error.message);
        console.error(error);
    }
}

function parseExcelData(rawData) {
    const parsed = [];
    for (let i = 1; i < rawData.length; i++) {
        const row = rawData[i];
        if (!row || row.length < 6) continue;
        parsed.push({
            photo: row[0] || '',
            size: row[1] || '',
            article: row[5] || '',
            color: row[4] || '',
            realSize: null,
            layerId: null
        });
    }
    return parsed;
}

function renderPrintsList() {
    const container = document.getElementById('printsList');
    container.innerHTML = '';
    
    if (excelData.length === 0) {
        container.innerHTML = '<div style="padding: 20px; text-align: center; color: #888;">Нет данных</div>';
        return;
    }
    
    excelData.forEach((item, index) => {
        const div = document.createElement('div');
        div.className = 'print-item';
        div.dataset.index = index;
        
        const photo = document.createElement('div');
        photo.className = 'print-item-photo';
        photo.textContent = '📷';
        
        const info = document.createElement('div');
        info.className = 'print-item-info';
        
        const header = document.createElement('div');
        header.className = 'print-item-header';
        header.innerHTML = `
            <span class="print-item-article">${item.article || 'Без артикула'}</span>
            <span class="print-item-size">${item.size || 'N/A'}</span>
        `;
        
        const details = document.createElement('div');
        details.className = 'print-item-details';
        details.innerHTML = `
            <div>Цвет: ${item.color || 'N/A'}</div>
            <div>Размер: ${item.realSize ? item.realSize + ' мм' : 'не задан'}</div>
        `;
        
        if (item.layerId) {
            const sizeInput = document.createElement('input');
            sizeInput.type = 'text';
            sizeInput.className = 'size-input';
            sizeInput.placeholder = 'Введите размер в мм';
            sizeInput.value = item.realSize || '';
            sizeInput.addEventListener('change', (e) => updatePrintSize(index, e.target.value));
            details.appendChild(sizeInput);
        }
        
        info.appendChild(header);
        info.appendChild(details);
        div.appendChild(photo);
        div.appendChild(info);
        div.addEventListener('click', () => selectPrint(index));
        container.appendChild(div);
    });
}

function selectPrint(index) {
    selectedPrintIndex = index;
    document.querySelectorAll('.print-item').forEach((el, i) => {
        el.classList.toggle('selected', i === index);
    });
    const item = excelData[index];
    if (item.layerId) selectLayerById(item.layerId);
}

async function selectLayerById(layerId) {
    try {
        await core.executeAsModal(async () => {
            const doc = app.activeDocument;
            if (!doc) return;
            const layers = await getAllLayers(doc);
            const targetLayer = layers.find(l => l.id === layerId);
            if (targetLayer) doc.activeLayers = [targetLayer];
        });
    } catch (error) {
        console.error('Error selecting layer:', error);
    }
}

async function getAllLayers(doc) {
    const layers = [];
    function traverse(layerSet) {
        for (const layer of layerSet.layers) {
            if (layer.typename === 'LayerSet') traverse(layer);
            else layers.push(layer);
        }
    }
    traverse(doc);
    return layers;
}

function setupLayerSelectionSync() {
    setInterval(async () => {
        try {
            const doc = app.activeDocument;
            if (!doc || !doc.activeLayers || doc.activeLayers.length === 0) return;
            const activeLayerId = doc.activeLayers[0].id;
            const index = excelData.findIndex(item => item.layerId === activeLayerId);
            if (index !== -1 && index !== selectedPrintIndex) selectPrintInUI(index);
        } catch (error) {}
    }, 500);
}

function selectPrintInUI(index) {
    selectedPrintIndex = index;
    document.querySelectorAll('.print-item').forEach((el, i) => {
        el.classList.toggle('selected', i === index);
    });
    const container = document.getElementById('printsList');
    const item = container.children[index];
    if (item) item.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

async function updatePrintSize(index, newSize) {
    const item = excelData[index];
    const sizeValue = parseFloat(newSize);
    if (isNaN(sizeValue) || sizeValue <= 0) {
        alert('Некорректный размер');
        return;
    }
    item.realSize = sizeValue;
    if (item.layerId) await resizeLayerInPhotoshop(item.layerId, sizeValue);
    renderPrintsList();
}

async function resizeLayerInPhotoshop(layerId, targetSizeMM) {
    try {
        await core.executeAsModal(async () => {
            const doc = app.activeDocument;
            if (!doc) return;
            const layers = await getAllLayers(doc);
            const layer = layers.find(l => l.id === layerId);
            if (!layer) return;
            
            const bounds = layer.bounds;
            const currentWidth = bounds.right - bounds.left;
            const currentHeight = bounds.bottom - bounds.top;
            const dpi = doc.resolution;
            const targetSizePx = (targetSizeMM / 25.4) * dpi;
            const maxCurrentSize = Math.max(currentWidth.value, currentHeight.value);
            const scale = (targetSizePx / maxCurrentSize) * 100;
            layer.scale(scale, scale);
        }, { commandName: 'Resize Layer' });
    } catch (error) {
        console.error('Error resizing layer:', error);
    }
}

async function runLayoutScript() {
    updateStatus('Запуск раскладки...');
    try {
        await core.executeAsModal(async () => {
            await createLayoutDocument();
        }, { commandName: 'Run Layout Script' });
        updateStatus('Раскладка завершена!');
        await mapLayersToData();
        renderPrintsList();
    } catch (error) {
        updateStatus('Ошибка: ' + error.message);
        console.error(error);
    }
}

async function createLayoutDocument() {
    const sheetWidth = 570;
    const sheetHeight = 2000;
    const dpi = 200;
    await app.documents.add({
        width: sheetWidth * 25.4 / 72,
        height: sheetHeight * 25.4 / 72,
        resolution: dpi,
        mode: 'RGBColorMode',
        fill: 'backgroundColor'
    });
}

async function mapLayersToData() {
    try {
        const doc = app.activeDocument;
        if (!doc) return;
        const layers = await getAllLayers(doc);
        for (const layer of layers) {
            const layerName = layer.name;
            const matchIndex = excelData.findIndex(item => {
                return layerName.includes(item.article) || layerName.includes(item.size);
            });
            if (matchIndex !== -1) {
                excelData[matchIndex].layerId = layer.id;
                const bounds = layer.bounds;
                const width = bounds.right.value - bounds.left.value;
                const height = bounds.bottom.value - bounds.top.value;
                const maxSize = Math.max(width, height);
                excelData[matchIndex].realSize = Math.round((maxSize / doc.resolution) * 25.4 * 10) / 10;
            }
        }
    } catch (error) {
        console.error('Error mapping layers:', error);
    }
}

function updateStatus(message) {
    document.getElementById('statusText').textContent = message;
}

console.log('Print Layout Manager initialized');