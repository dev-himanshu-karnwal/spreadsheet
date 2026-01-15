/**
 * Spreadsheet Component for Dataverse SOV Entity Integration
 * Refactored to strictly follow SOLID principles and OOP structure.
 */

const CONFIG = {
    INITIAL_COLS: 10,
    INITIAL_ROWS: 10,
    DEFAULT_COL_WIDTH: 130,
    MIN_SIZE: 20,
    COL_NAMES: [],
    DEFAULT_PAGE_SIZE: 25,
    PAGE_SIZE_OPTIONS: [10, 25, 50, 100]
};

// --- Infrastructure ---

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
 * NotificationSystem handles showing custom popups/toasts.
 * Responsibility: UI Notifications only.
 */
class NotificationSystem {
    constructor() {
        this.container = document.createElement('div');
        this.container.className = 'notification-container';
        document.body.appendChild(this.container);
    }

    show(title, message, type = 'error', duration = 4000) {
        const notification = document.createElement('div');
        notification.className = `notification ${type}`;

        const icon = type === 'error'
            ? `<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>`
            : `<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>`;

        notification.innerHTML = `
            <div class="notification-icon">${icon}</div>
            <div class="notification-content">
                <div class="notification-title">${title}</div>
                <div class="notification-message">${message}</div>
            </div>
            <div class="notification-close">
                <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </div>
        `;

        const closeBtn = notification.querySelector('.notification-close');
        const dismiss = () => {
            notification.classList.add('fade-out');
            setTimeout(() => notification.remove(), 300);
        };

        closeBtn.onclick = dismiss;
        setTimeout(dismiss, duration);

        this.container.appendChild(notification);
    }
}

// --- Data Domain ---

/**
 * DataConnector Interface
 * Responsibility: Low-level API communication.
 */
class IDataConnector {
    async createRecord(data) { throw new Error('Not implemented'); }
    async updateRecord(id, data) { throw new Error('Not implemented'); }
    async deleteRecord(id) { throw new Error('Not implemented'); }
    async bulkDelete(ids) { throw new Error('Not implemented'); }
    async bulkUpload(records) { throw new Error('Not implemented'); }
}

class DataverseConnector extends IDataConnector {
    async createRecord(data) {
        console.log('[Dataverse] Creating record:', data);
        return { id: Math.random().toString(36).substr(2, 9) };
    }

    async updateRecord(id, data) {
        console.log(`[Dataverse] Updating record ${id}:`, data);
    }

    async deleteRecord(id) {
        console.log(`[Dataverse] Deleting record ${id}`);
    }

    async bulkDelete(ids) {
        console.log(`[Dataverse] Bulk deleting ${ids.length} records`);
    }

    async bulkUpload(records) {
        console.log(`[Dataverse] Bulk uploading ${records.length} records`);
    }
}

/**
 * GridState
 * Responsibility: Manage internal representation of data and dimensions.
 */
class GridState extends Observable {
    constructor() {
        super();
        this.colsCount = CONFIG.INITIAL_COLS;
        this.colNames = [...CONFIG.COL_NAMES];
        this.colWidths = new Array(this.colsCount).fill(CONFIG.DEFAULT_COL_WIDTH);

        this.currentPage = 1;
        this.pageSize = CONFIG.DEFAULT_PAGE_SIZE;
        this.searchQuery = '';

        this.data = Array.from({ length: CONFIG.INITIAL_ROWS }, () =>
            Array.from({ length: this.colsCount }, () => ({ value: '', recordId: null }))
        );
    }

    setPage(page) {
        this.currentPage = page;
        this.notify({ type: 'PAGINATION_CHANGE' });
    }

    setPageSize(size) {
        this.pageSize = size;
        this.currentPage = 1;
        this.notify({ type: 'PAGINATION_CHANGE' });
    }

    setSearchQuery(query) {
        this.searchQuery = query;
        this.currentPage = 1;
        this.notify({ type: 'SEARCH_CHANGE' });
    }

    getFilteredData() {
        const allRows = this.data.map((row, index) => ({ row, index }));

        if (!this.searchQuery) return allRows;

        const lowerQuery = this.searchQuery.toLowerCase();
        return allRows.filter(item => item.row.some(cell =>
            (cell.value || '').toString().toLowerCase().includes(lowerQuery)
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

    updateSchema(newColNames, targetRowCount = null) {
        this.colNames = newColNames;
        this.colsCount = newColNames.length;

        const rowCount = targetRowCount !== null ? targetRowCount : this.data.length;

        // Reset/Resize data to match exact dimensions
        this.data = Array.from({ length: rowCount }, () =>
            Array.from({ length: this.colsCount }, () => ({ value: '', recordId: null }))
        );

        // Reset colWidths
        this.colWidths = new Array(this.colsCount).fill(CONFIG.DEFAULT_COL_WIDTH);

        this.notify({ type: 'SCHEMA_CHANGE' });
    }

    _ensureCapacity(row, col) {
        let changed = false;
        // Expand rows
        while (this.data.length <= row) {
            this.data.push(Array.from({ length: this.colsCount }, () => ({ value: '', recordId: null })));
            changed = true;
        }
        // Expand columns
        if (col >= this.colsCount) {
            const newColsCount = col + 1;
            this.data.forEach(rowData => {
                while (rowData.length < newColsCount) {
                    rowData.push({ value: '', recordId: null });
                }
            });
            while (this.colWidths.length < newColsCount) {
                this.colWidths.push(CONFIG.DEFAULT_COL_WIDTH);
            }
            this.colsCount = newColsCount;
            changed = true;
        }
        return changed;
    }

    addRow() {
        const newRow = Array.from({ length: this.colsCount }, () => ({ value: '', recordId: null }));
        this.data.push(newRow);
        this.notify({ type: 'DATA_CHANGE' });
        return this.data.length - 1;
    }

    updateCell(row, col, value, recordId = null, silent = false) {
        const capacityChanged = this._ensureCapacity(row, col);

        const wasActive = row < this.data.length ? this.isActiveRow(row) : false;

        this.data[row][col] = {
            value: value,
            recordId: recordId || (this.data[row][col] ? this.data[row][col].recordId : null)
        };

        if (silent) return;

        const isActive = this.isActiveRow(row);
        if (wasActive !== isActive || capacityChanged) {
            this.notify({ type: 'DATA_CHANGE' });
        } else {
            this.notify({ type: 'CELL_UPDATE', row, col, value });
        }
    }

    isActiveRow(rowIndex) {
        return this.data[rowIndex].some(cell => (cell.value !== null && cell.value !== '') || cell.recordId !== null);
    }

    setColWidth(index, width) {
        if (width < CONFIG.MIN_SIZE) return;
        this.colWidths[index] = width;
        this.notify({ type: 'DIMENSION_CHANGE' });
    }

    clearAll() {
        this.colsCount = CONFIG.INITIAL_COLS;
        this.colNames = [...CONFIG.COL_NAMES];
        this.colWidths = new Array(this.colsCount).fill(CONFIG.DEFAULT_COL_WIDTH);
        this.data = Array.from({ length: CONFIG.INITIAL_ROWS }, () =>
            Array.from({ length: this.colsCount }, () => ({ value: '', recordId: null }))
        );
        this.currentPage = 1;
        this.searchQuery = '';
        this.notify({ type: 'SCHEMA_CHANGE' });
    }

    getColName(index) {
        if (this.colNames && this.colNames[index]) return this.colNames[index];
        let name = '';
        let i = index;
        while (i >= 0) {
            name = String.fromCharCode((i % 26) + 65) + name;
            i = Math.floor(i / 26) - 1;
        }
        return name;
    }
}

// --- Services ---

/**
 * PersistenceService
 * Responsibility: Monitor state changes and sync with DataConnector.
 * Implements Dependency Inversion by depending on IDataConnector.
 */
class PersistenceService {
    constructor(state, connector) {
        this.state = state;
        this.connector = connector;
        this.init();
    }

    init() {
        this.state.subscribe(async (event) => {
            if (event.type === 'CELL_UPDATE') {
                await this.handleCellUpdate(event.row, event.col, event.value);
            }
        });
    }

    async handleCellUpdate(row, col, value) {
        const cellData = this.state.data[row][col];
        if (cellData.recordId) {
            await this.connector.updateRecord(cellData.recordId, { [`col_${col}`]: value });
        } else {
            const res = await this.connector.createRecord({ [`col_${col}`]: value });
            // For simplicity, all cells in this visible row share the same recordId
            for (let i = 0; i < this.state.colsCount; i++) {
                this.state.data[row][i].recordId = res.id;
            }
        }
    }

    async deleteRowRecords(row) {
        const recordId = this.state.data[row][0].recordId;
        if (recordId) {
            await this.connector.deleteRecord(recordId);
            for (let i = 0; i < this.state.colsCount; i++) {
                this.state.data[row][i].recordId = null;
            }
        }
    }
}

/**
 * ImportExportService
 * Responsibility: Handle file operations and clipboard parsing.
 */
class ImportExportService {
    constructor(state, connector, notifications) {
        this.state = state;
        this.connector = connector;
        this.notifications = notifications;
    }

    async handlePaste(pastedText) {
        if (!pastedText || pastedText.trim().length === 0) {
            throw new Error('No data found in clipboard.');
        }

        const lines = pastedText.split(/\r?\n/).filter(line => line.trim().length > 0);
        if (lines.length === 0) throw new Error('The pasted content is empty.');

        const firstLineCells = lines[0].split('\t');
        if (firstLineCells.length < 2 && lines.length < 2) {
            throw new Error('The pasted data format is not supported. Please paste a table from Excel.');
        }

        const headers = firstLineCells.slice(1);
        if (headers.length === 0 || headers.every(h => h.trim() === '')) {
            throw new Error('Headers are missing or invalid.');
        }

        // Set exact dimensions: rows = total lines - 1 (excluding header), cols = headers count
        this.state.updateSchema(headers, lines.length - 1);

        for (let rIdx = 1; rIdx < lines.length; rIdx++) {
            const cells = lines[rIdx].split('\t');
            // We use the relative index (rIdx - 1) into our newly resized empty grid
            const row = rIdx - 1;

            const rowData = {};
            for (let cIdx = 1; cIdx < cells.length; cIdx++) {
                const col = cIdx - 1;
                if (col >= this.state.colsCount) break;
                const value = cells[cIdx];
                this.state.updateCell(row, col, value, null, true);
                rowData[`col_${col}`] = value;
            }

            const res = await this.connector.createRecord(rowData);
            for (let c = 0; c < this.state.colsCount; c++) {
                this.state.data[row][c].recordId = res.id;
            }
        }
        this.state.notify({ type: 'DATA_CHANGE' });
    }

    async handleFileUpload(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = async (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const sheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                    if (jsonData.length === 0) throw new Error('File is empty');

                    const headers = jsonData[0].slice(1);
                    // Match file dimensions exactly
                    this.state.updateSchema(headers, jsonData.length - 1);

                    for (let i = 1; i < jsonData.length; i++) {
                        const cells = jsonData[i];
                        const row = i - 1;

                        const rowData = {};
                        for (let j = 1; j < cells.length; j++) {
                            const colIdx = j - 1;
                            if (colIdx < this.state.colsCount) {
                                const val = cells[j];
                                this.state.updateCell(row, colIdx, val, null, true);
                                rowData[`col_${colIdx}`] = val;
                            }
                        }
                        await this.connector.createRecord(rowData);
                    }
                    this.state.notify({ type: 'DATA_CHANGE' });
                    resolve(jsonData.length - 1);
                } catch (err) { reject(err); }
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    downloadCSV() {
        const rows = [];
        const headerRow = ['Row No'];
        for (let c = 0; c < this.state.colsCount; c++) {
            headerRow.push(this.state.getColName(c));
        }
        rows.push(headerRow);

        const activeRows = this.state.getFilteredData();
        activeRows.forEach(item => {
            const csvRow = [item.index + 1];
            item.row.forEach(cell => csvRow.push(cell.value || ''));
            rows.push(csvRow);
        });

        const worksheet = XLSX.utils.aoa_to_sheet(rows);
        const csvContent = XLSX.utils.sheet_to_csv(worksheet);
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'sov_export.csv';
        a.click();
        URL.revokeObjectURL(url);
    }
}

// --- UI Components ---

/**
 * StyleSystem
 * Responsibility: Dynamic CSS injection.
 */
class StyleSystem {
    constructor(state) {
        this.state = state;
        this.styleElement = document.createElement('style');
        document.head.appendChild(this.styleElement);
        this.init();
    }

    init() {
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
 * Base Component
 */
class Component {
    constructor(elementId, state) {
        this.container = typeof elementId === 'string' ? document.getElementById(elementId) : elementId;
        this.state = state;
    }
    render() { throw new Error('Render not implemented'); }
}

class HeaderComponent extends Component {
    render() {
        this.container.className = 'app-header';
        this.container.innerHTML = `
            <div class="header-info">
                <h1>SOV Data Management</h1>
                <p>View and manage your property records</p>
            </div>
            <div class="search-box">
                <span class="search-icon">
                    <svg viewBox="0 0 24 24" width="18" height="18"><path fill="currentColor" d="M15.5 14h-.79l-.28-.27A6.471 6.471 0 0 0 16 9.5 6.5 6.5 0 1 0 9.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z"/></svg>
                </span>
                <input type="text" id="grid-search" placeholder="Search records..." value="${this.state.searchQuery}">
            </div>
        `;
        this.container.querySelector('input').addEventListener('input', (e) => this.state.setSearchQuery(e.target.value));
    }
}

class ToolbarComponent extends Component {
    constructor(container, state, actions) {
        super(container, state);
        this.actions = actions;
    }

    render() {
        this.container.className = 'grid-toolbar';
        this.container.innerHTML = '';

        const deleteIcon = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a2 2 0 012-2h4a2 2 0 012 2v2M10 11v6M14 11v6"/></svg>`;
        const downloadIcon = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M7 10l5 5 5-5M12 15V3"/></svg>`;
        const uploadIcon = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M17 8l-5-5-5 5M12 3v12"/></svg>`;
        const plusIcon = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 5v14M5 12h14"/></svg>`;

        const btnAdd = this.createBtn('Add Row', 'btn-primary', () => this.actions.onAddRow(), plusIcon);
        const btnDelete = this.createBtn('Delete All', 'btn-danger', () => this.actions.onDeleteAll(), deleteIcon);
        const btnDownload = this.createBtn('Download', 'btn-secondary', () => this.actions.onDownload(), downloadIcon);

        const fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.style.display = 'none';
        fileInput.onchange = (e) => e.target.files[0] && this.actions.onUpload(e.target.files[0]);

        const btnUpload = this.createBtn('Upload Excel', 'btn-success', () => fileInput.click(), uploadIcon);

        this.container.append(btnAdd, btnDelete, btnDownload, btnUpload, fileInput);
    }

    createBtn(text, className, onClick, icon) {
        const btn = document.createElement('button');
        btn.className = `btn ${className}`;
        btn.innerHTML = `${icon} <span>${text}</span>`;
        btn.onclick = onClick;
        return btn;
    }
}

class GridBodyComponent extends Component {
    constructor(container, state, handlers) {
        super(container, state);
        this.handlers = handlers;
        this.cellElements = [];
    }

    render() {
        this.container.className = 'grid-container';
        this.container.innerHTML = '';
        this.cellElements = [];

        const paginatedData = this.state.getPaginatedData();

        // Header
        const headerRow = document.createElement('div');
        headerRow.className = 'grid-row header';
        headerRow.innerHTML = '<div class="grid-cell corner-header"></div>';

        for (let c = 0; c < this.state.colsCount; c++) {
            const cell = document.createElement('div');
            cell.className = `grid-cell col-header col-${c}`;
            cell.textContent = this.state.getColName(c);

            const resizer = document.createElement('div');
            resizer.className = 'col-resizer';
            resizer.onmousedown = (e) => this.handlers.onResize(e, c);
            cell.appendChild(resizer);
            headerRow.appendChild(cell);
        }
        this.container.appendChild(headerRow);

        // Body
        paginatedData.forEach((item, rIdxInPage) => {
            const rowEl = document.createElement('div');
            rowEl.className = `grid-row row-${item.index}`;

            const rowHeader = document.createElement('div');
            rowHeader.className = 'grid-cell row-header';
            rowHeader.textContent = item.index + 1;
            rowEl.appendChild(rowHeader);

            this.cellElements[rIdxInPage] = [];

            item.row.forEach((cellData, cIdx) => {
                const cell = document.createElement('div');
                cell.className = `grid-cell col-${cIdx}`;
                cell.contentEditable = true;
                cell.textContent = cellData.value;

                cell.onpaste = (e) => this.handlers.onPaste(e, item.index, cIdx);
                cell.onblur = (e) => this.handlers.onBlur(item.index, cIdx, e.target.textContent);
                cell.onkeydown = (e) => this.handlers.onKeyDown(e, rIdxInPage, cIdx, this.cellElements);

                rowEl.appendChild(cell);
                this.cellElements[rIdxInPage][cIdx] = cell;
            });
            this.container.appendChild(rowEl);
        });
    }

    updateCellValue(rInPage, c, value) {
        if (this.cellElements[rInPage] && this.cellElements[rInPage][c]) {
            this.cellElements[rInPage][c].textContent = value;
        }
    }
}

class PaginationComponent extends Component {
    render() {
        this.container.className = 'pagination-container';
        const totalPages = this.state.getTotalPages();
        const currentPage = this.state.currentPage;

        this.container.innerHTML = `
            <div class="page-size-selector">
                <label>Rows per page:</label>
                <select id="page-size">
                    ${CONFIG.PAGE_SIZE_OPTIONS.map(s => `<option value="${s}" ${this.state.pageSize === s ? 'selected' : ''}>${s}</option>`).join('')}
                </select>
            </div>
            <div class="page-controls">
                <button class="btn btn-secondary btn-sm" id="prev-page" ${currentPage === 1 ? 'disabled' : ''}>Previous</button>
                <span class="page-info">Page ${currentPage} of ${totalPages}</span>
                <button class="btn btn-secondary btn-sm" id="next-page" ${currentPage === totalPages ? 'disabled' : ''}>Next</button>
            </div>
        `;

        this.container.querySelector('#page-size').onchange = (e) => this.state.setPageSize(parseInt(e.target.value));
        this.container.querySelector('#prev-page').onclick = () => this.state.setPage(currentPage - 1);
        this.container.querySelector('#next-page').onclick = () => this.state.setPage(currentPage + 1);
    }
}

/**
 * MainRenderer
 * Responsibility: Orchestrate UI components.
 */
class MainRenderer {
    constructor(root, state, services, notifications) {
        this.root = root;
        this.state = state;
        this.services = services;
        this.notifications = notifications;

        this.components = {
            header: new HeaderComponent(document.createElement('div'), state),
            toolbar: new ToolbarComponent(document.createElement('div'), state, this.getToolbarActions()),
            grid: new GridBodyComponent(document.createElement('div'), state, this.getGridHandlers()),
            pagination: new PaginationComponent(document.createElement('div'), state)
        };

        this.init();
    }

    init() {
        this.root.append(
            this.components.header.container,
            this.components.toolbar.container,
            this.components.grid.container,
            this.components.pagination.container
        );

        this.state.subscribe((event) => {
            if (event.type === 'CELL_UPDATE') {
                const paginatedData = this.state.getPaginatedData();
                const rInPage = paginatedData.findIndex(item => item.index === event.row);
                if (rInPage !== -1) this.components.grid.updateCellValue(rInPage, event.col, event.value);
            } else {
                this.renderAll();
            }
        });
    }

    renderAll() {
        Object.values(this.components).forEach(c => c.render());
    }

    getToolbarActions() {
        return {
            onAddRow: async () => {
                const newIdx = this.state.addRow();
                const res = await this.services.connector.createRecord({});
                for (let c = 0; c < this.state.colsCount; c++) {
                    this.state.data[newIdx][c].recordId = res.id;
                }

                const filtered = this.state.getFilteredData();
                const idx = filtered.findIndex(item => item.index === newIdx);
                if (idx !== -1) {
                    this.state.setPage(Math.ceil((idx + 1) / this.state.pageSize));
                }
            },
            onDeleteAll: async () => {
                if (!confirm('Clear all?')) return;
                const ids = [...new Set(this.state.data.flatMap(r => r.map(c => c.recordId)).filter(Boolean))];
                await this.services.connector.bulkDelete(ids);
                this.state.clearAll();
                this.notifications.show('Success', 'All records cleared and schema reset', 'success');
            },
            onDownload: () => this.services.importExport.downloadCSV(),
            onUpload: async (file) => {
                try {
                    const count = await this.services.importExport.handleFileUpload(file);
                    this.notifications.show('Success', `Imported ${count} rows`, 'success');
                } catch (e) {
                    this.notifications.show('Error', e.message, 'error');
                }
            }
        };
    }

    getGridHandlers() {
        return {
            onResize: (e, index) => {
                const startX = e.pageX;
                const startW = this.state.colWidths[index];
                const onMove = (me) => this.state.setColWidth(index, Math.max(CONFIG.MIN_SIZE, startW + (me.pageX - startX)));
                const onUp = () => {
                    document.removeEventListener('mousemove', onMove);
                    document.removeEventListener('mouseup', onUp);
                };
                document.addEventListener('mousemove', onMove);
                document.addEventListener('mouseup', onUp);
            },
            onPaste: async (e, r, c) => {
                e.preventDefault();
                try {
                    await this.services.importExport.handlePaste(e.clipboardData.getData('text'));
                    this.notifications.show('Success', 'Imported data from clipboard', 'success');
                } catch (err) {
                    this.notifications.show('Error', err.message, 'error');
                    this.state.clearAll();
                }
            },
            onBlur: (r, c, val) => this.state.updateCell(r, c, val),
            onKeyDown: (e, rInPage, c, elements) => {
                const helpers = {
                    move: (dr, dc) => {
                        const tr = rInPage + dr, tc = c + dc;
                        if (elements[tr] && elements[tr][tc]) {
                            e.preventDefault();
                            const target = elements[tr][tc];
                            target.focus();
                            const range = document.createRange();
                            range.selectNodeContents(target);
                            window.getSelection().removeAllRanges();
                            window.getSelection().addRange(range);
                            target.scrollIntoView({ block: 'nearest', inline: 'nearest' });
                        }
                    }
                };

                if (e.key === 'ArrowUp') helpers.move(-1, 0);
                else if (e.key === 'ArrowDown') helpers.move(1, 0);
                else if (e.key === 'ArrowLeft' && window.getSelection().anchorOffset === 0) helpers.move(0, -1);
                else if (e.key === 'ArrowRight' && window.getSelection().anchorOffset === elements[rInPage][c].textContent.length) helpers.move(0, 1);
                else if (e.key === 'Enter') helpers.move(e.shiftKey ? -1 : 1, 0);
                else if (e.key === 'Tab') helpers.move(0, e.shiftKey ? -1 : 1);
                else if ((e.key === 'Delete' || e.key === 'Backspace') && elements[rInPage][c].textContent === '') {
                    this.services.persistence.deleteRowRecords(this.state.getPaginatedData()[rInPage].index);
                }
            }
        };
    }
}

// --- App Root ---

class SpreadsheetApp {
    constructor() {
        this.notifications = new NotificationSystem();
        this.state = new GridState();
        this.connector = new DataverseConnector();

        this.services = {
            persistence: new PersistenceService(this.state, this.connector),
            importExport: new ImportExportService(this.state, this.connector, this.notifications),
            style: new StyleSystem(this.state),
            connector: this.connector // for convenience
        };

        this.renderer = new MainRenderer(document.getElementById('app'), this.state, this.services, this.notifications);
    }

    init() {
        this.renderer.renderAll();
    }
}

document.addEventListener('DOMContentLoaded', () => {
    new SpreadsheetApp().init();
});
