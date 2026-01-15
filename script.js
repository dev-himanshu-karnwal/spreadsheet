/**
 * Configuration Constants
 * Open for extension via configuration injection if needed.
 */
const CONFIG = {
    COLS_COUNT: 26,
    ROWS_COUNT: 100,
    DEFAULT_COL_WIDTH: 100,
    DEFAULT_ROW_HEIGHT: 30,
    MIN_SIZE: 20
};

/**
 * Observer Pattern Implementation
 * Base class for State management to allow loose coupling.
 */
class Observable {
    constructor() {
        this.observers = [];
    }

    subscribe(fn) {
        this.observers.push(fn);
    }

    notify(data) {
        this.observers.forEach(observer => observer(data));
    }
}

/**
 * GridState
 * RESPONSIBILITY: Single Source of Truth for grid data (dimensions, content).
 * FOLLOWS: Single Responsibility Principle (SRP).
 */
class GridState extends Observable {
    constructor(cols = CONFIG.COLS_COUNT, rows = CONFIG.ROWS_COUNT) {
        super();
        this.colWidths = new Array(cols).fill(CONFIG.DEFAULT_COL_WIDTH);
        this.rowHeights = new Array(rows).fill(CONFIG.DEFAULT_ROW_HEIGHT);
    }

    setColWidth(index, width) {
        if (width < CONFIG.MIN_SIZE) return;
        this.colWidths[index] = width;
        this.notify({ type: 'COL_RESIZE', index, width });
    }

    setRowHeight(index, height) {
        if (height < CONFIG.MIN_SIZE) return;
        this.rowHeights[index] = height;
        this.notify({ type: 'ROW_RESIZE', index, height });
    }

    getColWidth(index) {
        return this.colWidths[index];
    }

    getRowHeight(index) {
        return this.rowHeights[index];
    }
}

/**
 * StyleSystem
 * RESPONSIBILITY: Dynamic CSS transformations.
 * FOLLOWS: Dependency Inversion (depends on inputs, not concrete DOM details of other components).
 */
class StyleSystem {
    constructor(state) {
        this.state = state;
        this.styleElement = document.createElement('style');
        document.head.appendChild(this.styleElement);

        // Initial render
        this.updateAll();

        // React to state changes
        this.state.subscribe((event) => {
            // Optimization: In a real app we might update specific rules, 
            // but for simplicity we re-generate or validly update the sheet.
            // Here we re-generate for correctness guarantee.
            this.updateAll();
        });
    }

    updateAll() {
        let css = '';
        this.state.colWidths.forEach((w, i) => {
            css += `.col-${i} { width: ${w}px; min-width: ${w}px; }\n`;
        });
        this.state.rowHeights.forEach((h, i) => {
            css += `.row-${i} { height: ${h}px; }\n`;
        });
        css += `.header-col { height: ${CONFIG.DEFAULT_ROW_HEIGHT}px; }\n`;
        this.styleElement.textContent = css;
    }
}

/**
 * ResizeController
 * RESPONSIBILITY: Handling user input details for resizing.
 * FOLLOWS: Interface Segregation (Specific tool for resizing interactions).
 */
class ResizeController {
    constructor(state) {
        this.state = state;
        this.resizing = null; // { type, index, startPos, startSize }

        this.handleMouseMove = this.handleMouseMove.bind(this);
        this.handleMouseUp = this.handleMouseUp.bind(this);

        window.addEventListener('mousemove', this.handleMouseMove);
        window.addEventListener('mouseup', this.handleMouseUp);
    }

    startResize(type, index, event) {
        event.preventDefault();
        event.stopPropagation();

        this.resizing = {
            type,
            index,
            startPos: type === 'col' ? event.clientX : event.clientY,
            startSize: type === 'col' ? this.state.getColWidth(index) : this.state.getRowHeight(index)
        };

        document.body.style.cursor = type === 'col' ? 'col-resize' : 'row-resize';
        event.target.classList.add('active');
    }

    handleMouseMove(e) {
        if (!this.resizing) return;

        const { type, index, startPos, startSize } = this.resizing;
        const currentPos = type === 'col' ? e.clientX : e.clientY;
        const delta = currentPos - startPos;
        const newSize = startSize + delta;

        if (type === 'col') {
            this.state.setColWidth(index, newSize);
        } else {
            this.state.setRowHeight(index, newSize);
        }
    }

    handleMouseUp() {
        if (!this.resizing) return;
        this.resizing = null;
        document.body.style.cursor = 'default';
        document.querySelectorAll('.active').forEach(el => el.classList.remove('active'));
    }
}

/**
 * UI Component Base
 * Helper for creating DOM elements.
 */
class Component {
    createEl(tag, className = '', text = '') {
        const el = document.createElement(tag);
        if (className) el.className = className;
        if (text) el.textContent = text;
        return el;
    }
}

/**
 * GridRenderer
 * RESPONSIBILITY: Constructing the Visual Grid.
 * FOLLOWS: Single Responsibility (Rendering only).
 */
class GridRenderer extends Component {
    constructor(rootElement, state, resizeController) {
        super();
        this.root = rootElement;
        this.state = state;
        this.resizeController = resizeController;
    }

    render() {
        // Toolbar
        const toolbar = this.createEl('div', 'toolbar', 'Spreadsheet OOP');
        this.root.appendChild(toolbar);

        // Container
        const container = this.createEl('div', 'spreadsheet-container');
        this.root.appendChild(container);

        // Content
        const content = this.createEl('div', 'grid-content');
        container.appendChild(content);

        this.renderHeaderRow(content);
        this.renderRows(content);
    }

    renderHeaderRow(container) {
        const headerRow = this.createEl('div', 'row');

        // Corner
        const corner = this.createEl('div', 'corner-header');
        headerRow.appendChild(corner);

        // Columns A-Z
        for (let c = 0; c < this.state.colWidths.length; c++) {
            const letter = String.fromCharCode(65 + c);
            const cell = this.createEl('div', `cell header-col col-${c}`, letter);

            const resizer = this.createEl('div', 'resizer-col');
            resizer.addEventListener('mousedown', (e) => this.resizeController.startResize('col', c, e));
            cell.appendChild(resizer);

            headerRow.appendChild(cell);
        }
        container.appendChild(headerRow);
    }

    renderRows(container) {
        for (let r = 0; r < this.state.rowHeights.length; r++) {
            const row = this.createEl('div', `row row-${r}`);

            // Row Header
            const rowHeader = this.createEl('div', 'cell header-row', (r + 1).toString());
            const resizer = this.createEl('div', 'resizer-row');
            resizer.addEventListener('mousedown', (e) => this.resizeController.startResize('row', r, e));
            rowHeader.appendChild(resizer);
            row.appendChild(rowHeader);

            // Cells
            for (let c = 0; c < this.state.colWidths.length; c++) {
                const cell = this.createEl('div', `cell col-${c}`);
                cell.contentEditable = true;
                row.appendChild(cell);
            }
            container.appendChild(row);
        }
    }
}

/**
 * Application Core
 * RESPONSIBILITY: Bootstrapping and wiring dependencies.
 */
class SpreadsheetApp {
    constructor(rootId) {
        this.root = document.getElementById(rootId);
        this.state = new GridState(CONFIG.COLS_COUNT, CONFIG.ROWS_COUNT);
        this.styleSystem = new StyleSystem(this.state);
        this.resizeController = new ResizeController(this.state);
        this.renderer = new GridRenderer(this.root, this.state, this.resizeController);
    }

    init() {
        console.log('App Initializing...');
        this.renderer.render();
    }
}

// Entry Point
document.addEventListener('DOMContentLoaded', () => {
    const app = new SpreadsheetApp('app');
    app.init();
});
