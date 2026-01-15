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
    ROWS_COUNT: 200, // Increased to support more data
    DEFAULT_COL_WIDTH: 130,
    DEFAULT_ROW_HEIGHT: 35,
    MIN_SIZE: 20,
    ENTITY_NAME: 'SOV_Entity',
    COL_NAMES: [],
    DEFAULT_PAGE_SIZE: 25,
    PAGE_SIZE_OPTIONS: [10, 25, 50, 100]
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

        this.currentPage = 1;
        this.pageSize = CONFIG.DEFAULT_PAGE_SIZE;
        this.searchQuery = '';

        // data[rowIndex][colIndex] = { value: '', recordId: null }
        this.data = Array.from({ length: CONFIG.ROWS_COUNT }, () =>
            Array.from({ length: this.colsCount }, () => ({ value: '', recordId: null }))
        );
    }

    setPage(page) {
        this.currentPage = page;
        this.notify({ type: 'PAGINATION_CHANGE' });
    }

    setPageSize(size) {
        this.pageSize = size;
        this.currentPage = 1; // Reset to first page when size changes
        this.notify({ type: 'PAGINATION_CHANGE' });
    }

    setSearchQuery(query) {
        this.searchQuery = query;
        this.currentPage = 1; // Reset to first page on search
        this.notify({ type: 'SEARCH_CHANGE' });
    }

    getFilteredData() {
        if (!this.searchQuery) return this.data.map((row, index) => ({ row, index }));

        const lowerQuery = this.searchQuery.toLowerCase();
        return this.data.map((row, index) => ({ row, index }))
            .filter(item => item.row.some(cell =>
                cell.value.toString().toLowerCase().includes(lowerQuery)
            ));
    }

    getPaginatedData() {
        const filtered = this.getFilteredData();
        const start = (this.currentPage - 1) * this.pageSize;
        return filtered.slice(start, start + this.pageSize);
    }

    getTotalPages() {
        const filtered = this.getFilteredData();
        return Math.ceil(filtered.length / this.pageSize) || 1;
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
        const lines = pastedText.split(/\r?\n/).filter(line => line.trim().length > 0);

        if (lines.length === 0) return;

        // Treat the first row as headers and update schema
        const headers = lines[0].split('\t');
        this.state.updateSchema(headers);

        // Remaining lines are treated as data rows starting from index 0
        // (Since updateSchema resets the data)
        for (let rIdx = 1; rIdx < lines.length; rIdx++) {
            const row = rIdx - 1;
            if (row >= CONFIG.ROWS_COUNT) break;

            const cells = lines[rIdx].split('\t');
            const rowDataToCreate = {};

            for (let cIdx = 0; cIdx < cells.length; cIdx++) {
                if (cIdx >= this.state.colsCount) break;

                const value = cells[cIdx];
                this.state.updateCell(row, cIdx, value);
                rowDataToCreate[`col_${cIdx}`] = value;
            }

            // After pasting a full row, we simulate creating it in Dataverse
            const res = await this.connector.createRecord(rowDataToCreate);
            // Update the state with the recordId returned from "Dataverse"
            for (let cIdx = 0; cIdx < cells.length; cIdx++) {
                if (cIdx < this.state.colsCount) {
                    this.state.data[row][cIdx].recordId = res.id;
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
                const paginatedData = this.state.getPaginatedData();
                const rowIndexInPage = paginatedData.findIndex(item => item.index === event.row);
                if (rowIndexInPage !== -1 && this.cellElements[rowIndexInPage] && this.cellElements[rowIndexInPage][event.col]) {
                    this.cellElements[rowIndexInPage][event.col].textContent = event.value;
                }
            } else if (event.type === 'BULK_CLEAR' || event.type === 'SCHEMA_CHANGE' || event.type === 'PAGINATION_CHANGE' || event.type === 'SEARCH_CHANGE') {
                this.renderGrid();
                this.renderPagination();
            }
        });
    }

    render() {
        this.root.innerHTML = '';

        const appHeader = document.createElement('div');
        appHeader.className = 'app-header';

        const headerInfo = document.createElement('div');
        headerInfo.className = 'header-info';
        headerInfo.innerHTML = '<h1>SOV Data Management</h1><p>View and manage your property records</p>';

        const searchBox = document.createElement('div');
        searchBox.className = 'search-box';
        searchBox.innerHTML = `
            <span class="search-icon">
                <svg viewBox="0 0 24 24" width="18" height="18">
                    <path fill="currentColor" d="M15.5 14h-.79l-.28-.27A6.471 6.471 0 0 0 16 9.5 6.5 6.5 0 1 0 9.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z"/>
                </svg>
            </span>
            <input type="text" id="grid-search" placeholder="Search records...">
        `;

        const searchInput = searchBox.querySelector('input');
        searchInput.addEventListener('input', (e) => {
            this.state.setSearchQuery(e.target.value);
        });

        appHeader.append(headerInfo, searchBox);
        this.root.appendChild(appHeader);

        this.renderToolbar();

        this.gridContainer = document.createElement('div');
        this.gridContainer.className = 'grid-container';
        this.root.appendChild(this.gridContainer);

        this.paginationContainer = document.createElement('div');
        this.paginationContainer.className = 'pagination-container';
        this.root.appendChild(this.paginationContainer);

        this.renderGrid();
        this.renderPagination();
    }

    renderToolbar() {
        const toolbar = document.createElement('div');
        toolbar.className = 'grid-toolbar';

        const deleteIcon = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a2 2 0 012-2h4a2 2 0 012 2v2M10 11v6M14 11v6"/></svg>`;
        const downloadIcon = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M7 10l5 5 5-5M12 15V3"/></svg>`;
        const uploadIcon = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M17 8l-5-5-5 5M12 3v12"/></svg>`;

        const btnBulkDelete = this.createBtn('Delete All', 'btn-danger', () => this.actions.bulkDelete(), deleteIcon, 'Permanently remove all records');
        const btnDownload = this.createBtn('Download', 'btn-primary', () => this.actions.downloadTemplate(), downloadIcon, 'Download CSV template with current headers');

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

        const btnUpload = this.createBtn('Upload Excel', 'btn-success', () => uploadInput.click(), uploadIcon, 'Populate grid from Excel or CSV file');

        toolbar.append(btnBulkDelete, btnDownload, btnUpload, uploadInput);
        this.root.appendChild(toolbar);
    }

    renderGrid() {
        if (!this.gridContainer) return;
        this.gridContainer.innerHTML = '';
        this.cellElements = [];

        const paginatedData = this.state.getPaginatedData();

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
        paginatedData.forEach((item, rIdxInPage) => {
            const r = item.index;
            const rowData = item.row;

            const rowEl = document.createElement('div');
            rowEl.className = `grid-row row-${r}`;

            const rowHeader = document.createElement('div');
            rowHeader.className = 'grid-cell row-header';
            rowHeader.textContent = r + 1;

            rowEl.appendChild(rowHeader);

            this.cellElements[rIdxInPage] = [];

            for (let c = 0; c < this.state.colsCount; c++) {
                const cell = document.createElement('div');
                cell.className = `grid-cell col-${c}`;
                cell.contentEditable = true;
                cell.textContent = rowData[c].value;

                cell.addEventListener('paste', (e) => this.clipboard.handlePaste(e, r, c));
                cell.addEventListener('blur', (e) => this.handleCellEdit(r, c, e.target.textContent));
                cell.addEventListener('keydown', (e) => {
                    if (e.key === 'Delete' || e.key === 'Backspace') {
                        if (cell.textContent === '') this.handleCellDelete(r, c);
                    }

                    const moveFocus = (rowOffset, colOffset) => {
                        let targetRowInPage = rIdxInPage + rowOffset;
                        let targetCol = c + colOffset;

                        if (targetRowInPage >= 0 && targetRowInPage < paginatedData.length &&
                            targetCol >= 0 && targetCol < this.state.colsCount) {
                            const target = this.cellElements[targetRowInPage][targetCol];
                            if (target) {
                                e.preventDefault();
                                target.focus();

                                const range = document.createRange();
                                range.selectNodeContents(target);
                                const selection = window.getSelection();
                                selection.removeAllRanges();
                                selection.addRange(range);

                                target.scrollIntoView({ block: 'nearest', inline: 'nearest' });
                            }
                        }
                    };

                    if (e.key === 'ArrowUp') moveFocus(-1, 0);
                    else if (e.key === 'ArrowDown') moveFocus(1, 0);
                    else if (e.key === 'ArrowLeft') {
                        const selection = window.getSelection();
                        if (selection.anchorOffset === 0 || selection.toString() === cell.textContent) moveFocus(0, -1);
                    } else if (e.key === 'ArrowRight') {
                        const selection = window.getSelection();
                        if (selection.anchorOffset === cell.textContent.length || selection.toString() === cell.textContent) moveFocus(0, 1);
                    } else if (e.key === 'Enter') moveFocus(e.shiftKey ? -1 : 1, 0);
                    else if (e.key === 'Tab') moveFocus(0, e.shiftKey ? -1 : 1);
                });

                this.cellElements[rIdxInPage][c] = cell;
                rowEl.appendChild(cell);
            }
            this.gridContainer.appendChild(rowEl);
        });
    }

    renderPagination() {
        if (!this.paginationContainer) return;
        this.paginationContainer.innerHTML = '';

        const totalPages = this.state.getTotalPages();
        const currentPage = this.state.currentPage;

        // Page Size Selector
        const sizeSelector = document.createElement('div');
        sizeSelector.className = 'page-size-selector';
        sizeSelector.innerHTML = `
            <label>Rows per page:</label>
            <select>
                ${CONFIG.PAGE_SIZE_OPTIONS.map(size => `
                    <option value="${size}" ${this.state.pageSize === size ? 'selected' : ''}>${size}</option>
                `).join('')}
            </select>
        `;
        sizeSelector.querySelector('select').addEventListener('change', (e) => {
            this.state.setPageSize(parseInt(e.target.value));
        });

        // Page Controls
        const pageControls = document.createElement('div');
        pageControls.className = 'page-controls';

        const prevBtn = document.createElement('button');
        prevBtn.className = 'btn btn-secondary btn-sm';
        prevBtn.innerHTML = 'Previous';
        prevBtn.disabled = currentPage === 1;
        prevBtn.onclick = () => this.state.setPage(currentPage - 1);

        const pageInfo = document.createElement('span');
        pageInfo.className = 'page-info';
        pageInfo.textContent = `Page ${currentPage} of ${totalPages}`;

        const nextBtn = document.createElement('button');
        nextBtn.className = 'btn btn-secondary btn-sm';
        nextBtn.innerHTML = 'Next';
        nextBtn.disabled = currentPage === totalPages;
        nextBtn.onclick = () => this.state.setPage(currentPage + 1);

        pageControls.append(prevBtn, pageInfo, nextBtn);

        this.paginationContainer.append(sizeSelector, pageControls);
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

    handleSearch(query) {
        this.state.setSearchQuery(query);
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

    createBtn(text, className, onClick, iconSvg, tooltip) {
        const btn = document.createElement('button');
        btn.className = `btn ${className}`;
        if (tooltip) btn.setAttribute('data-tooltip', tooltip);

        if (iconSvg) {
            btn.innerHTML = `${iconSvg} <span>${text}</span>`;
        } else {
            btn.textContent = text;
        }

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
