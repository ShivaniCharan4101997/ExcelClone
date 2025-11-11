const TOTAL_COLS = 26;
const TOTAL_ROWS = 100;

document.addEventListener('DOMContentLoaded', () => {
    // --- Get Elements ---
    const colNameContainer = document.querySelector('.column-name-container');
    const rowNameContainer = document.querySelector('.row-name-container');
    const sheet = document.querySelector('#sheet');
    const formulaInput = document.querySelector('#formula-input');
    const selectedCellDisplay = document.querySelector('#selected-cell');
    const addSheetBtn = document.querySelector('.icon-add');
    const sheetBar = document.querySelector('.sheet-bar');
    const leftScroll = document.querySelector('.icon-left-scroll');
    const rightScroll = document.querySelector('.icon-right-scroll');
    const tabsContainer = document.querySelector('.tabs-container');

    const iconBold = document.querySelector('.icon-bold');
    const iconItalic = document.querySelector('.icon-italic');
    const iconUnderline = document.querySelector('.icon-underline');
    const fontFamilySelector = document.querySelector('.font-family-selector');
    const fontSizeSelector = document.querySelector('.font-size-selector');
    const alignLeft = document.querySelector('.icon-align-left');
    const alignCenter = document.querySelector('.icon-align-center');
    const alignRight = document.querySelector('.icon-align-right');
    const textColorPicker = document.querySelector('.text-color-picker');
    const bgColorPicker = document.querySelector('.bg-color-picker');
    const iconClear = document.querySelector('.icon-clear');
    const iconCut = document.querySelector('.icon-cut');
    const iconCopy = document.querySelector('.icon-copy');
    const iconPaste = document.querySelector('.icon-paste');

    // --- App State ---
    let selectedCell = null;
    let sheetData = {};
    let currentSheet = localStorage.getItem('activeSheet') || 'Sheet1';
    let clipboard = null;
    let isCutAction = false;

    // --- Load & Save ---
    const loadSheetData = () => {
        const saved = localStorage.getItem('sheetData');
        sheetData = saved ? JSON.parse(saved) : { Sheet1: {} };
    };
    const saveSheetData = () => localStorage.setItem('sheetData', JSON.stringify(sheetData));
    loadSheetData();

    // --- Helpers ---
    const getColumnName = n => {
        let ans = '';
        while (n > 0) {
            let rem = n % 26;
            if (rem === 0) {
                ans = 'Z' + ans;
                n = Math.floor(n / 26) - 1;
            } else {
                ans = String.fromCharCode(64 + rem) + ans;
                n = Math.floor(n / 26);
            }
        }
        return ans;
    };

    const getCellData = ref => {
        if (!sheetData[currentSheet][ref])
            sheetData[currentSheet][ref] = { value: '', formula: '', style: {}, dependents: [] };
        return sheetData[currentSheet][ref];
    };

    const applyStyle = (cell, style) => {
        cell.style.fontWeight = style.bold ? 'bold' : '';
        cell.style.fontStyle = style.italic ? 'italic' : '';
        cell.style.textDecoration = style.underline ? 'underline' : '';
        cell.style.fontFamily = style.fontFamily || '';
        cell.style.fontSize = style.fontSize || '';
        cell.style.textAlign = style.textAlign || '';
        cell.style.color = style.color || '';
        cell.style.backgroundColor = style.backgroundColor || '';
    };

    const toggleStyle = (styleKey, cssProp, value) => {
        if (!selectedCell) return;
        const cellData = getCellData(selectedCell.dataset.cell);
        const isActive = cellData.style[styleKey] === value;
        cellData.style[styleKey] = isActive ? false : value;
        selectedCell.style[cssProp] = isActive ? '' : value;
        saveSheetData();
    };

    const setAlignment = align => {
        if (!selectedCell) return;
        const cellData = getCellData(selectedCell.dataset.cell);
        cellData.style.textAlign = align;
        selectedCell.style.textAlign = align;
        saveSheetData();
    };

    const switchSheet = tabElement => {
        document.querySelectorAll('.sheet-tab').forEach(tab => tab.classList.remove('selected'));
        tabElement.classList.add('selected');
        currentSheet = tabElement.textContent;
        renderSavedSheet();
    };

    // --- Render UI ---
    const renderColumns = () => {
        const frag = document.createDocumentFragment();
        for (let i = 1; i <= TOTAL_COLS; i++) {
            const el = document.createElement('div');
            el.className = 'column-name';
            el.textContent = getColumnName(i);
            frag.appendChild(el);
        }
        colNameContainer.appendChild(frag);
    };

    const renderRows = () => {
        const frag = document.createDocumentFragment();
        for (let i = 1; i <= TOTAL_ROWS; i++) {
            const el = document.createElement('div');
            el.className = 'row-name';
            el.textContent = i;
            frag.appendChild(el);
        }
        rowNameContainer.appendChild(frag);
    };

    const renderCells = () => {
        const frag = document.createDocumentFragment();
        for (let r = 1; r <= TOTAL_ROWS; r++) {
            for (let c = 1; c <= TOTAL_COLS; c++) {
                const cell = document.createElement('div');
                cell.className = 'cell';
                cell.contentEditable = true;
                cell.dataset.cell = `${getColumnName(c)}${r}`;
                frag.appendChild(cell);
            }
        }
        sheet.appendChild(frag);
    };

    const renderSavedSheet = () => {
        document.querySelectorAll('.cell').forEach(cell => {
            const cellData = sheetData[currentSheet][cell.dataset.cell] || { value: '', style: {} };
            cell.textContent = cellData.value;
            applyStyle(cell, cellData.style);
        });
    };

    const renderSheetTabs = () => {
        document.querySelectorAll('.sheet-tab').forEach(tab => tab.remove());
        Object.keys(sheetData).forEach(name => {
            const tab = document.createElement('div');
            tab.className = 'menu-item sheet-tab';
            tab.textContent = name;
            if (name === currentSheet) tab.classList.add('selected');
            tab.addEventListener('click', () => switchSheet(tab));
            tabsContainer.appendChild(tab);
        });
    };

    // --- Initialize ---
    renderColumns();
    renderRows();
    renderCells();
    renderSavedSheet();
    renderSheetTabs();

    // --- Cell Selection & Formula Sync ---
    sheet.addEventListener('click', e => {
        if (!e.target.classList.contains('cell')) return;
        if (selectedCell) selectedCell.classList.remove('selected');
        selectedCell = e.target;
        selectedCell.classList.add('selected');
        selectedCellDisplay.textContent = selectedCell.dataset.cell;
        formulaInput.textContent = getCellData(selectedCell.dataset.cell).value;
    });

    // --- Formula / Cell Input ---
    const evaluateFormula = (formula, currentCell) => {
        if (!formula.startsWith('=')) return formula;
        formula = formula.slice(1);
        const refs = formula.match(/([A-Z]+[0-9]+)/g) || [];
        refs.forEach(ref => {
            if (!sheetData[currentSheet][ref])
                sheetData[currentSheet][ref] = { value: '', formula: '', style: {}, dependents: [] };
            if (!sheetData[currentSheet][ref].dependents.includes(currentCell))
                sheetData[currentSheet][ref].dependents.push(currentCell);
        });
        const replacedFormula = formula.replace(/([A-Z]+[0-9]+)/g, match => sheetData[currentSheet][match]?.value || 0);
        try { return eval(replacedFormula); } catch { return 'ERROR'; }
    };

    const updateCell = ref => {
        const cell = document.querySelector(`[data-cell="${ref}"]`);
        const cellData = sheetData[currentSheet][ref];
        if (cellData.formula) cellData.value = evaluateFormula(cellData.formula, ref);
        cell.textContent = cellData.value;
        cellData.dependents.forEach(dep => updateCell(dep));
    };

    formulaInput.addEventListener('input', () => {
        if (!selectedCell) return;
        const ref = selectedCell.dataset.cell;
        const input = formulaInput.textContent;
        const cellData = getCellData(ref);
        if (input.startsWith('=')) cellData.formula = input;
        else { cellData.formula = ''; cellData.value = input; }
        updateCell(ref);
        saveSheetData();
    });

    sheet.addEventListener('input', e => {
        if (!e.target.classList.contains('cell')) return;
        const ref = e.target.dataset.cell;
        const input = e.target.textContent;
        const cellData = getCellData(ref);
        if (input.startsWith('=')) cellData.formula = input;
        else { cellData.formula = ''; cellData.value = input; }
        updateCell(ref);
        formulaInput.textContent = input;
        saveSheetData();
    });

    // --- Text Styling ---
    iconBold.addEventListener('click', () => toggleStyle('bold', 'fontWeight', 'bold'));
    iconItalic.addEventListener('click', () => toggleStyle('italic', 'fontStyle', 'italic'));
    iconUnderline.addEventListener('click', () => toggleStyle('underline', 'textDecoration', 'underline'));
    fontFamilySelector.addEventListener('change', () => {
        if (!selectedCell) return;
        const val = fontFamilySelector.value;
        const cellData = getCellData(selectedCell.dataset.cell);
        cellData.style.fontFamily = val;
        selectedCell.style.fontFamily = val;
        saveSheetData();
    });
    fontSizeSelector.addEventListener('change', () => {
        if (!selectedCell) return;
        const val = fontSizeSelector.value + 'px';
        const cellData = getCellData(selectedCell.dataset.cell);
        cellData.style.fontSize = val;
        selectedCell.style.fontSize = val;
        saveSheetData();
    });
    alignLeft.addEventListener('click', () => setAlignment('left'));
    alignCenter.addEventListener('click', () => setAlignment('center'));
    alignRight.addEventListener('click', () => setAlignment('right'));
    textColorPicker.addEventListener('change', () => {
        if (!selectedCell) return;
        const val = textColorPicker.value;
        const cellData = getCellData(selectedCell.dataset.cell);
        cellData.style.color = val;
        selectedCell.style.color = val;
        saveSheetData();
    });
    bgColorPicker.addEventListener('change', () => {
        if (!selectedCell) return;
        const val = bgColorPicker.value;
        const cellData = getCellData(selectedCell.dataset.cell);
        cellData.style.backgroundColor = val;
        selectedCell.style.backgroundColor = val;
        saveSheetData();
    });
    iconClear.addEventListener('click', () => {
        if (!selectedCell) return;
        selectedCell.removeAttribute('style');
        getCellData(selectedCell.dataset.cell).style = {};
        saveSheetData();
    });

    // --- Cut / Copy / Paste ---
    const copyCell = () => { if (!selectedCell) return; clipboard = JSON.parse(JSON.stringify(getCellData(selectedCell.dataset.cell))); isCutAction = false; };
    const cutCell = () => {
        if (!selectedCell) return;
        clipboard = JSON.parse(JSON.stringify(getCellData(selectedCell.dataset.cell)));
        isCutAction = true;
        selectedCell.textContent = '';
        selectedCell.removeAttribute('style');
        sheetData[currentSheet][selectedCell.dataset.cell] = { value: '', style: {}, formula: '', dependents: [] };
        saveSheetData();
    };
    const pasteCell = () => {
        if (!selectedCell || !clipboard) return;
        sheetData[currentSheet][selectedCell.dataset.cell] = JSON.parse(JSON.stringify(clipboard));
        selectedCell.textContent = clipboard.value;
        applyStyle(selectedCell, clipboard.style);
        if (isCutAction) clipboard = null;
        isCutAction = false;
        saveSheetData();
    };
    iconCopy.addEventListener('click', copyCell);
    iconCut.addEventListener('click', cutCell);
    iconPaste.addEventListener('click', pasteCell);

    // --- Sheet Bar Scroll ---
    leftScroll.addEventListener('click', () => tabsContainer.scrollBy({ left: -100, behavior: 'smooth' }));
    rightScroll.addEventListener('click', () => tabsContainer.scrollBy({ left: 100, behavior: 'smooth' }));

    // --- Add New Sheet ---
    addSheetBtn.addEventListener('click', () => {
        const newSheetName = `Sheet${Object.keys(sheetData).length + 1}`;
        sheetData[newSheetName] = {};
        currentSheet = newSheetName;
        saveSheetData();
        renderSheetTabs();
        renderSavedSheet();
    });

    // --- Auto-save every 5s ---
    setInterval(saveSheetData, 5000);
});
