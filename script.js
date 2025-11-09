const TOTAL_COLS = 26;
const TOTAL_ROWS = 100;

document.addEventListener('DOMContentLoaded', function () {
    // get the elements
    const colNameContainer = document.querySelector('.column-name-container');
    const rowNameContainer = document.querySelector('.row-name-container');
    const sheet = document.querySelector('#sheet');
    const formulaInput = document.querySelector('#formula-input');
    const selectedCellDisplay = document.querySelector('#selected-cell');
    const addSheetBtn = document.querySelector('.icon-add');
    const sheetBar = document.querySelector('.sheet-bar');
    const iconBold = document.querySelector('.icon-bold');
    const iconItalic = document.querySelector('.icon-italic');
    const iconUnderline = document.querySelector('.icon-underline');

    //  --- App State ---
    let selectedCell = null;
    let sheetData = {};  // { Sheet1: { A1: { value:'', style:{} }, ... }, Sheet2: {} }
    let currentSheet = "Sheet1";

    // --- Load from localStorage ---
    function loadSheetData() {
        const saved = localStorage.getItem("sheetData");
        if (saved) {
            sheetData = JSON.parse(saved);
        } else {
            sheetData = { Sheet1: {} };
        }
    }

    // --- Save to localStorage ---
    function saveSheetData() {
        localStorage.setItem("sheetData", JSON.stringify(sheetData));
    }

    loadSheetData();

    // --- Convert number → Excel-style column (A, B, ..., Z, AA...) ---
    function getColumnName(n) {
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
    }

    // --- 1 Render Column Headers ---
    const colFrag = document.createDocumentFragment();
    for (let i = 1; i <= TOTAL_COLS; i++) {
        const el = document.createElement('div');
        el.className = 'column-name';
        el.id = `colCode-${i}`;
        el.textContent = getColumnName(i);
        colFrag.appendChild(el);
    }
    colNameContainer.appendChild(colFrag);

    // --- 2 Render Row Numbers ---
    const rowFrag = document.createDocumentFragment();
    for (let i = 1; i <= TOTAL_ROWS; i++) {
        const el = document.createElement('div');
        el.className = 'row-name';
        el.textContent = i;
        rowFrag.appendChild(el);
    }
    rowNameContainer.appendChild(rowFrag);

    // --- 3 Render Cells ---
    const cellFrag = document.createDocumentFragment();
    for (let r = 1; r <= TOTAL_ROWS; r++) {
        for (let c = 1; c <= TOTAL_COLS; c++) {
            const cell = document.createElement('div');
            cell.className = 'cell';
            cell.contentEditable = true;

            const colName = getColumnName(c);
            const cellRef = `${colName}${r}`;
            cell.dataset.cell = cellRef;
            cell.id = `r-${r}-c-${c}`;
            cellFrag.appendChild(cell);
        }
    }
    sheet.appendChild(cellFrag);

    // --- 4 Select Cell Logic ---
    sheet.addEventListener('click', (e) => {
        if (!e.target.classList.contains('cell')) return;

        if (selectedCell) selectedCell.classList.remove('selected');
        selectedCell = e.target;
        selectedCell.classList.add('selected');

        const cellRef = selectedCell.dataset.cell;
        selectedCellDisplay.textContent = cellRef;

        const cellData = sheetData[currentSheet][cellRef] || {};
        formulaInput.textContent = cellData.value || '';
    });

    // --- 5 Formula Bar Sync ---
    formulaInput.addEventListener('input', () => {
        if (!selectedCell) return;
        selectedCell.textContent = formulaInput.textContent;

        const ref = selectedCell.dataset.cell;
        if (!sheetData[currentSheet][ref]) sheetData[currentSheet][ref] = { value: '', style: {} };
        sheetData[currentSheet][ref].value = selectedCell.textContent;
        saveSheetData();
    });

    // --- 6 Cell Typing Syncs Formula Bar ---
    sheet.addEventListener('input', (e) => {
        if (e.target.classList.contains('cell') && e.target === selectedCell) {
            formulaInput.textContent = e.target.textContent;

            const ref = e.target.dataset.cell;
            if (!sheetData[currentSheet][ref]) sheetData[currentSheet][ref] = { value: '', style: {} };
            sheetData[currentSheet][ref].value = e.target.textContent;
            saveSheetData();
        }
    });

    // --- 7 Text Style Toggles ---
    function toggleStyle(styleKey, cssProp, value) {
        if (!selectedCell) return;
        const ref = selectedCell.dataset.cell;
        if (!sheetData[currentSheet][ref]) sheetData[currentSheet][ref] = { value: '', style: {} };
        const cellObj = sheetData[currentSheet][ref];

        const current = cellObj.style[styleKey];
        const isActive = current === value;
        if (isActive) {
            cellObj.style[styleKey] = false;
            selectedCell.style[cssProp] = '';
        } else {
            cellObj.style[styleKey] = value;
            selectedCell.style[cssProp] = value;
        }
        saveSheetData();
    }

    iconBold.addEventListener('click', () => toggleStyle('bold', 'fontWeight', 'bold'));
    iconItalic.addEventListener('click', () => toggleStyle('italic', 'fontStyle', 'italic'));
    iconUnderline.addEventListener('click', () => toggleStyle('underline', 'textDecoration', 'underline'));

    // --- 8 Add New Sheet ---
    addSheetBtn.addEventListener('click', () => {
        const sheetCount = Object.keys(sheetData).length + 1;
        const newSheetName = `Sheet${sheetCount}`;
        sheetData[newSheetName] = {};
        currentSheet = newSheetName;
        saveSheetData();

        // Add new tab in UI
        const newTab = document.createElement('div');
        newTab.className = 'menu-item sheet-tab';
        newTab.textContent = newSheetName;
        sheetBar.appendChild(newTab);
    });

    // --- 9️⃣ Render Saved Sheet ---
    function renderSavedSheet() {
        const data = sheetData[currentSheet];
        if (!data) return;

        Object.entries(data).forEach(([cellRef, cellObj]) => {
            const cell = document.querySelector(`[data-cell="${cellRef}"]`);
            if (!cell) return;

            cell.textContent = cellObj.value || '';

            const s = cellObj.style || {};
            cell.style.fontWeight = s.bold ? 'bold' : '';
            cell.style.fontStyle = s.italic ? 'italic' : '';
            cell.style.textDecoration = s.underline ? 'underline' : '';
        });
    }

    renderSavedSheet();

    // --- 🔁 Optional Auto-save every 5s ---
    setInterval(saveSheetData, 5000);




//     menu bar


    const menuItems = document.querySelectorAll('.menu-item');

    menuItems.forEach(item => {
        item.addEventListener('click', () => {
            // remove previous selection
            menuItems.forEach(i => i.classList.remove('selected'));

            // highlight current tab
            item.classList.add('selected');

            // show related toolbar (future)
            console.log(`Switched to ${item.textContent} tab`);
        });
    });

});
