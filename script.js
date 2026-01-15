/**
 * Spreadsheet Component for Dataverse SOV Entity Integration
 * Follows SOLID principles and OOP structure.
 * 
 * DESIGN NOTES:
 * - GridState: Manages the internal representation of data.
 * - DataConnector: Interface/Placeholder for Dataverse API logic.
 * - ClipboardController: Handles Excel copy-paste (TSV parsing).
 * - ActionController: Handles Bulk Upload/Download/Delete.
 */

const CONFIG = {
    COLS_COUNT: 40,
    ROWS_COUNT: 50,
    DEFAULT_COL_WIDTH: 130,
    DEFAULT_ROW_HEIGHT: 35,
    MIN_SIZE: 20,
    ENTITY_NAME: 'SOV_Entity',
    COL_NAMES: [
        'Loc #', 'Bldg #', 'Complex Name', 'Address', 'City', 'State',
        'Zip Code', 'County', 'Building Value', 'Contents Value',
        'Association Income', 'TIV', 'Construction Type',
        'Protection Class', '# of Stories', 'Year Built',
        'Roof Update', 'Roof Type'
    ]
};

/**
 * Observable pattern for state synchronization
 */
class Observable {
    constructor() {
        this.observers = [];
    }
    subscribe(fn) { this.observers.push(fn); }
    notify(data) { this.observers.forEach(observer => observer(data)); }
}

/**
 * DataConnector (PLACEHOLDER)
 * This is where a developer would later add the actual 
 * WebAPI calls for Dataverse.
 */
class DataConnector {
    async createRecord(data) {
        console.log('[Dataverse] Creating record:', data);
        // return await Xrm.WebApi.createRecord(CONFIG.ENTITY_NAME, data);
        return { id: Math.random().toString(36).substr(2, 9) };
    }

    async updateRecord(id, data) {
        console.log(`[Dataverse] Updating record ${id}:`, data);
        // return await Xrm.WebApi.updateRecord(CONFIG.ENTITY_NAME, id, data);
    }

    async deleteRecord(id) {
        console.log(`[Dataverse] Deleting record ${id}`);
        // return await Xrm.WebApi.deleteRecord(CONFIG.ENTITY_NAME, id);
    }

    async bulkDelete(ids) {
        console.log(`[Dataverse] Bulk deleting ${ids.length} records`);
        // Implementation for batch delete
    }

    async bulkUpload(records) {
        console.log(`[Dataverse] Bulk uploading ${records.length} records`);
        // Implementation for batch create
    }
}

/**
 * GridState
 * Manages dimensions and cell data.
 */
class GridState extends Observable {
    constructor() {
        super();
        this.colsCount = CONFIG.COLS_COUNT;
        this.colNames = [...CONFIG.COL_NAMES];
        this.colWidths = new Array(this.colsCount).fill(CONFIG.DEFAULT_COL_WIDTH);

        // data[rowIndex][colIndex] = { value: '', recordId: null }
        this.data = Array.from({ length: CONFIG.ROWS_COUNT }, () =>
            Array.from({ length: this.colsCount }, () => ({ value: '', recordId: null }))
        );
    }

    updateSchema(newColNames) {
        this.colNames = newColNames;
        this.colsCount = newColNames.length;
        this.colWidths = new Array(this.colsCount).fill(CONFIG.DEFAULT_COL_WIDTH);
        // We reset the data when schema changes as per user request (show only x columns from file)
        this.data = Array.from({ length: CONFIG.ROWS_COUNT }, () =>
            Array.from({ length: this.colsCount }, () => ({ value: '', recordId: null }))
        );
        this.notify({ type: 'SCHEMA_CHANGE' });
    }

    updateCell(row, col, value, recordId = null) {
        if (row >= CONFIG.ROWS_COUNT || col >= this.colsCount) return;
        this.data[row][col] = {
            value: value,
            recordId: recordId || this.data[row][col].recordId
        };
        this.notify({ type: 'CELL_UPDATE', row, col, value });
    }

    setColWidth(index, width) {
        if (width < CONFIG.MIN_SIZE) return;
        this.colWidths[index] = width;
        this.notify({ type: 'DIMENSION_CHANGE' });
    }

    clearAll() {
        this.data = this.data.map(row => row.map(() => ({ value: '', recordId: null })));
        this.notify({ type: 'BULK_CLEAR' });
    }
}

/**
 * StyleSystem handles dynamic CSS insertion for resizes.
 */
class StyleSystem {
    constructor(state) {
        this.state = state;
        this.styleElement = document.createElement('style');
        document.head.appendChild(this.styleElement);
        this.updateStyles();
        this.state.subscribe(() => this.updateStyles());
    }

    updateStyles() {
        let css = '';
        this.state.colWidths.forEach((w, i) => {
            css += `.col-${i} { width: ${w}px; min-width: ${w}px; }\n`;
        });
        this.styleElement.textContent = css;
    }
}

/**
 * ClipboardController handles Excel Paste logic.
 */
class ClipboardController {
    constructor(state, connector) {
        this.state = state;
        this.connector = connector;
    }

    async handlePaste(e, startRow, startCol) {
        e.preventDefault();
        const clipboardData = e.clipboardData || window.clipboardData;
        const pastedText = clipboardData.getData('text');

        // Excel uses Tab to separate cells and Newline to separate rows
        const rows = pastedText.split(/\r?\n/).filter(line => line.length > 0);

        for (let rIdx = 0; rIdx < rows.length; rIdx++) {
            const row = startRow + rIdx;
            if (row >= CONFIG.ROWS_COUNT) break;

            const cells = rows[rIdx].split('\t');
            const rowDataToCreate = {};

            for (let cIdx = 0; cIdx < cells.length; cIdx++) {
                const col = startCol + cIdx;
                if (col >= this.state.colsCount) break;

                const value = cells[cIdx];
                this.state.updateCell(row, col, value);
                rowDataToCreate[`col_${col}`] = value;
            }

            // After pasting a full row, we simulate creating it in Dataverse
            const res = await this.connector.createRecord(rowDataToCreate);
            // Update the state with the recordId returned from "Dataverse"
            for (let cIdx = 0; cIdx < cells.length; cIdx++) {
                const col = startCol + cIdx;
                if (col < this.state.colsCount) {
                    this.state.data[row][col].recordId = res.id;
                }
            }
        }
    }
}

/**
 * ActionController handles Bulk Actions (Delete, Upload, Download).
 */
class ActionController {
    constructor(state, connector) {
        this.state = state;
        this.connector = connector;
    }

    async bulkDelete() {
        if (!confirm('Are you sure you want to delete all records?')) return;

        const ids = [];
        this.state.data.forEach(row => {
            row.forEach(cell => {
                if (cell.recordId && !ids.includes(cell.recordId)) ids.push(cell.recordId);
            });
        });

        await this.connector.bulkDelete(ids);
        this.state.clearAll();
    }

    downloadTemplate() {
        const headers = this.state.colNames.join(',');
        const blob = new Blob([headers], { type: 'text/csv' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.setAttribute('hidden', '');
        a.setAttribute('href', url);
        a.setAttribute('download', 'sov_template.csv');
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    }

    async bulkUpload(file) {
        const reader = new FileReader();
        reader.onload = async (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // Get first sheet
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            // Convert to JSON (array of arrays)
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            console.log('Uploading Excel content with SheetJS...');

            if (jsonData.length > 0) {
                const headers = jsonData[0];
                this.state.updateSchema(headers);

                // Skip header if needed, here we assume row 0 is headers
                for (let i = 1; i < jsonData.length && i <= CONFIG.ROWS_COUNT; i++) {
                    const cells = jsonData[i];
                    const rowData = {};

                    cells.forEach((val, j) => {
                        if (j < this.state.colsCount) {
                            this.state.updateCell(i - 1, j, val);
                            rowData[`col_${j}`] = val;
                        }
                    });
                    await this.connector.createRecord(rowData);
                }
            }
        };
        reader.readAsArrayBuffer(file);
    }
}

/**
 * GridRenderer handles building the UI.
 */
class GridRenderer {
    constructor(root, state, clipboard, actions, connector) {
        this.root = root;
        this.state = state;
        this.clipboard = clipboard;
        this.actions = actions;
        this.connector = connector;
        this.cellElements = [];
        this.gridContainer = null;

        // Subscribe to state updates
        this.state.subscribe((event) => {
            if (event.type === 'CELL_UPDATE') {
                if (this.cellElements[event.row] && this.cellElements[event.row][event.col]) {
                    this.cellElements[event.row][event.col].textContent = event.value;
                }
            } else if (event.type === 'BULK_CLEAR') {
                this.cellElements.forEach(row => row.forEach(cell => cell.textContent = ''));
            } else if (event.type === 'SCHEMA_CHANGE') {
                this.renderGrid();
            }
        });
    }

    render() {
        this.root.innerHTML = '';
        this.renderToolbar();

        this.gridContainer = document.createElement('div');
        this.gridContainer.className = 'grid-container';
        this.root.appendChild(this.gridContainer);

        this.renderGrid();
    }

    renderToolbar() {
        const toolbar = document.createElement('div');
        toolbar.className = 'grid-toolbar';

        const btnBulkDelete = this.createBtn('Bulk Delete', 'btn-danger', () => this.actions.bulkDelete());
        const btnDownload = this.createBtn('Download Template', 'btn-primary', () => this.actions.downloadTemplate());

        const uploadInput = document.createElement('input');
        uploadInput.type = 'file';
        uploadInput.id = 'bulk-upload-input';
        uploadInput.style.display = 'none';
        uploadInput.addEventListener('change', (e) => {
            if (e.target.files[0]) {
                this.actions.bulkUpload(e.target.files[0]);
                e.target.value = ''; // Reset input
            }
        });

        const btnUpload = this.createBtn('Upload Excel', 'btn-success', () => uploadInput.click());

        toolbar.append(btnBulkDelete, btnDownload, btnUpload, uploadInput);
        this.root.appendChild(toolbar);
    }

    renderGrid() {
        if (!this.gridContainer) return;
        this.gridContainer.innerHTML = '';
        this.cellElements = [];

        // Header Row
        const headerRow = document.createElement('div');
        headerRow.className = 'grid-row header';

        const corner = document.createElement('div');
        corner.className = 'grid-cell corner-header';
        headerRow.appendChild(corner);

        for (let c = 0; c < this.state.colsCount; c++) {
            const cell = document.createElement('div');
            cell.className = `grid-cell col-header col-${c}`;
            cell.textContent = this.getColName(c);

            const resizer = document.createElement('div');
            resizer.className = 'col-resizer';
            resizer.addEventListener('mousedown', (e) => this.handleResizerMouseDown(e, 'col', c));
            cell.appendChild(resizer);

            headerRow.appendChild(cell);
        }
        this.gridContainer.appendChild(headerRow);

        // Grid Rows
        for (let r = 0; r < CONFIG.ROWS_COUNT; r++) {
            const rowEl = document.createElement('div');
            rowEl.className = `grid-row row-${r}`;

            const rowHeader = document.createElement('div');
            rowHeader.className = 'grid-cell row-header';
            rowHeader.textContent = r + 1;

            rowEl.appendChild(rowHeader);

            this.cellElements[r] = [];

            for (let c = 0; c < this.state.colsCount; c++) {
                const cell = document.createElement('div');
                cell.className = `grid-cell col-${c}`;
                cell.contentEditable = true;
                cell.textContent = this.state.data[r][c].value; // Set initial value if any

                cell.addEventListener('paste', (e) => this.clipboard.handlePaste(e, r, c));
                cell.addEventListener('blur', (e) => this.handleCellEdit(r, c, e.target.textContent));
                cell.addEventListener('keydown', (e) => {
                    if (e.key === 'Delete' || e.key === 'Backspace') {
                        if (cell.textContent === '') this.handleCellDelete(r, c);
                    }
                });

                this.cellElements[r][c] = cell;
                rowEl.appendChild(cell);
            }
            this.gridContainer.appendChild(rowEl);
        }
    }

    async handleCellEdit(r, c, value) {
        const cellData = this.state.data[r][c];
        if (cellData.value === value) return;

        this.state.data[r][c].value = value;

        if (cellData.recordId) {
            await this.connector.updateRecord(cellData.recordId, { [`col_${c}`]: value });
        } else {
            const res = await this.connector.createRecord({ [`col_${c}`]: value });
            // Map all cells in this visible row to the same recordId for simplicity (Excel row behavior)
            for (let i = 0; i < this.state.colsCount; i++) {
                this.state.data[r][i].recordId = res.id;
            }
        }
    }

    async handleCellDelete(r, c) {
        const cellData = this.state.data[r][c];
        if (cellData.recordId) {
            await this.connector.deleteRecord(cellData.recordId);
            // Clear recordId from this row locally
            for (let i = 0; i < this.state.colsCount; i++) {
                this.state.data[r][i].recordId = null;
            }
        }
    }

    handleResizerMouseDown(e, type, index) {
        e.preventDefault();
        const startPos = e.pageX;
        const startSize = this.state.colWidths[index];
        const resizer = e.target;
        resizer.classList.add('resizing');

        const onMouseMove = (moveEvent) => {
            const currentPos = moveEvent.pageX;
            const delta = currentPos - startPos;
            const newSize = Math.max(CONFIG.MIN_SIZE, startSize + delta);

            this.state.setColWidth(index, newSize);
        };

        const onMouseUp = () => {
            resizer.classList.remove('resizing');
            document.removeEventListener('mousemove', onMouseMove);
            document.removeEventListener('mouseup', onMouseUp);
        };

        document.addEventListener('mousemove', onMouseMove);
        document.addEventListener('mouseup', onMouseUp);
    }

    createBtn(text, className, onClick) {
        const btn = document.createElement('button');
        btn.textContent = text;
        btn.className = `btn ${className}`;
        btn.onclick = onClick;
        return btn;
    }

    getColName(index) {
        // Use provided column names first
        if (this.state.colNames[index]) return this.state.colNames[index];

        // Fallback to Excel-style naming
        let name = '';
        let i = index;
        while (i >= 0) {
            name = String.fromCharCode((i % 26) + 65) + name;
            i = Math.floor(i / 26) - 1;
        }
        return name;
    }
}

/**
 * App Bootstrapper
 */
class SpreadsheetApp {
    constructor() {
        this.state = new GridState();
        this.connector = new DataConnector();
        this.style = new StyleSystem(this.state);
        this.clipboard = new ClipboardController(this.state, this.connector);
        this.actions = new ActionController(this.state, this.connector);
        this.renderer = new GridRenderer(document.getElementById('app'), this.state, this.clipboard, this.actions, this.connector);
    }

    init() {
        this.renderer.render();
    }
}

document.addEventListener('DOMContentLoaded', () => {
    const app = new SpreadsheetApp();
    app.init();
});
