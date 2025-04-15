const XLSX = require("xlsx");

let currentData = [];
let undoStack = [];
let redoStack = [];

function saveStateForUndo() {
    const tableHTML = document.getElementById("excelTable").innerHTML;
    undoStack.push(tableHTML);
    redoStack = []; // clear redo stack on new action
}

document.getElementById("excelFile").addEventListener("change", (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 }); // raw 2D array

        currentData = json;
        renderTable(json);
    };
    reader.readAsArrayBuffer(file);
});

let selectedRowIndex = null;

function openRankModal(currentIndex) {
    selectedRowIndex = currentIndex;
    const modal = document.getElementById('rankModal');
    const input = document.getElementById('rankInput');

    modal.style.display = 'flex';
    input.value = '';
    input.focus();
}

document.getElementById('closeModal').onclick = () => {
    document.getElementById('rankModal').style.display = 'none';
};

document.getElementById('submitRank').onclick = () => {
    const inputVal = parseInt(document.getElementById('rankInput').value);
    if (isNaN(inputVal)) return;

    const totalRows = document.querySelectorAll('.excel-row').length;
    let newIndex = inputVal - 1;

    if (inputVal < 1) newIndex = 0;
    else if (inputVal > totalRows) newIndex = totalRows - 1;

    if (selectedRowIndex !== null) {
        const container = document.getElementById('excelTable');
        const rows = Array.from(container.children);
        const rowToMove = rows[selectedRowIndex];

        rows.splice(selectedRowIndex, 1); // remove
        rows.splice(newIndex, 0, rowToMove); // insert

        container.innerHTML = ''; // clear
        rows.forEach(row => container.appendChild(row)); // re-render
    }

    document.getElementById('rankModal').style.display = 'none';
};


function createRankPopup() {
    let popup = document.getElementById("rankPopup");
    if (!popup) {
        popup = document.createElement("div");
        popup.id = "rankPopup";
        // (add styles...)
        document.body.appendChild(popup);
    }

    let overlay = document.getElementById("popupOverlay");
    if (!overlay) {
        overlay = document.createElement("div");
        overlay.id = "popupOverlay";
        // (add styles...)
        document.body.appendChild(overlay);
    }

    popup.innerHTML = `
        <label for="rankInput">Enter rank:</label>
        <input type="number" id="rankInput" style="width: 100%; padding: 8px; margin-bottom: 15px; box-sizing: border-box;" />
        <div style="display: flex; justify-content: space-between;">
            <button id="rankCancelBtn">Cancel</button>
            <button id="rankOkBtn">OK</button>
        </div>
    `;

    return {
        popup,
        input: popup.querySelector("#rankInput"),
        okBtn: popup.querySelector("#rankOkBtn"),
        cancelBtn: popup.querySelector("#rankCancelBtn"),
        overlay
    };
}


function showRankPopup(rowId, currentRank) {
    const { popup, input, okBtn, cancelBtn, overlay } = rankPopupElements;

    // input.value = 0;
    popup.style.display = "block";
    overlay.style.display = "block";

    setTimeout(() => input.focus(), 100);

    const handleConfirm = () => {
        const targetRank = input.value;
        hideRankPopup();
        moveRowToRank(rowId, targetRank);
        cleanup();
    };

    const cleanup = () => {
        okBtn.removeEventListener("click", handleConfirm);
        input.removeEventListener("keydown", onEnterKey);
    };

    const onEnterKey = (e) => {
        if (e.key === "Enter") {
            e.preventDefault();
            handleConfirm();
        }
    };

    okBtn.addEventListener("click", handleConfirm);
    input.addEventListener("keydown", onEnterKey);
    cancelBtn.addEventListener("click", () => {
        hideRankPopup();
        cleanup();
    });
}

function hideRankPopup() {
    const { popup, overlay } = rankPopupElements;
    popup.style.display = "none";
    overlay.style.display = "none";
}


// Move row to the specified rank
function moveRowToRank(rowId, targetRank) {
    // Validate input
    if (!targetRank || targetRank === "") return;

    targetRankNum = parseInt(targetRank, 10);
    if (isNaN(targetRankNum)) return;

    // Get the row to be moved
    const row = document.querySelector(`tr[data-index="${rowId}"]`);
    if (!row) return;
    saveStateForUndo();
    // Get all rows and validate range
    const tbody = document.querySelector("#excelTable tbody");
    const allRows = Array.from(tbody.querySelectorAll("tr"));

    if (targetRankNum < 1) {
        targetRankNum = 1;
    } else if (targetRankNum > allRows.length) {
        targetRankNum = allRows.length;
    }

    // Get the target position
    const targetIndex = targetRankNum - 1;
    const targetRow = allRows[targetIndex];

    if (targetRow) {
        const currentIndex = parseInt(row.dataset.index) - 1;

        // Move the row
        if (targetIndex < currentIndex) {
            tbody.insertBefore(row, targetRow); // Moving up
        } else if (targetIndex > currentIndex) {
            if (targetRow.nextSibling) {
                tbody.insertBefore(row, targetRow.nextSibling); // Moving down
            } else {
                tbody.appendChild(row); // Moving to the end
            }
        }
    }

    // Update all rank numbers
    updateRanks();
}

// Create context menu
function createContextMenu() {
    // Remove existing menu if any
    const existingMenu = document.getElementById("contextMenu");
    if (existingMenu) {
        existingMenu.remove();
    }

    const menu = document.createElement("div");
    menu.id = "contextMenu";
    menu.style.position = "absolute";
    menu.style.display = "none";
    menu.style.backgroundColor = "white";
    menu.style.border = "1px solid #ccc";
    menu.style.boxShadow = "2px 2px 5px rgba(0,0,0,0.2)";
    menu.style.padding = "5px 0";
    menu.style.zIndex = "1000";

    // Add menu items
    const goToLocationItem = document.createElement("div");
    goToLocationItem.textContent = "Set Rank";
    goToLocationItem.style.padding = "8px 15px";
    goToLocationItem.style.cursor = "pointer";
    goToLocationItem.addEventListener("mouseover", () => {
        goToLocationItem.style.backgroundColor = "#f2f2f2";
    });
    goToLocationItem.addEventListener("mouseout", () => {
        goToLocationItem.style.backgroundColor = "white";
    });
    goToLocationItem.addEventListener("click", () => {
        const rowId = menu.dataset.rowId;
        const row = document.querySelector(`tr[data-index="${rowId}"]`);
        if (row) {
            const currentRank = row.querySelector(".rank-cell").textContent;
            showRankPopup(rowId, currentRank);

        }
        hideContextMenu();
    });

    const deleteRowItem = document.createElement("div");
    deleteRowItem.textContent = "Delete Row";
    deleteRowItem.style.padding = "8px 15px";
    deleteRowItem.style.cursor = "pointer";
    deleteRowItem.addEventListener("mouseover", () => {
        deleteRowItem.style.backgroundColor = "#f2f2f2";
    });
    deleteRowItem.addEventListener("mouseout", () => {
        deleteRowItem.style.backgroundColor = "white";
    });
    deleteRowItem.addEventListener("click", () => {
        const rowId = menu.dataset.rowId;
        deleteRow(rowId);
        hideContextMenu();
    });

    menu.appendChild(goToLocationItem);
    menu.appendChild(deleteRowItem);
    document.body.appendChild(menu);

    // Hide menu when clicking elsewhere
    document.addEventListener("click", (e) => {
        if (e.target.closest("#contextMenu") === null) {
            hideContextMenu();
        }
    });
}

function showContextMenu(x, y, rowId) {
    const menu = document.getElementById("contextMenu");
    if (!menu) {
        createContextMenu();
    }
    const activeMenu = document.getElementById("contextMenu");
    activeMenu.style.display = "block";
    activeMenu.style.left = `${x}px`;
    activeMenu.style.top = `${y}px`;
    activeMenu.dataset.rowId = rowId;
}

function hideContextMenu() {
    const menu = document.getElementById("contextMenu");
    if (menu) {
        menu.style.display = "none";
    }
}

function deleteRow(rowId) {
    // Find the row to delete
    const row = document.querySelector(`tr[data-index="${rowId}"]`);
    if (!row) return;
    saveStateForUndo();
    // Remove the row from the DOM
    row.remove();

    // Update the ranks of remaining rows
    updateRanks();
}

function undoAction() {
    if (undoStack.length === 0) return;

    const current = document.getElementById("excelTable").innerHTML;
    redoStack.push(current);

    const previousState = undoStack.pop();
    document.getElementById("excelTable").innerHTML = previousState;

    updateRanks(); // reassign correct ranks
    reattachRowEventListeners(); // 🔁 rebind events
    enableRowDragDrop();
}

function redoAction() {
    if (redoStack.length === 0) return;

    const current = document.getElementById("excelTable").innerHTML;
    undoStack.push(current);

    const nextState = redoStack.pop();
    document.getElementById("excelTable").innerHTML = nextState;

    updateRanks(); // reassign correct ranks
    reattachRowEventListeners(); // 🔁 rebind events
    enableRowDragDrop();
}

function reattachRowEventListeners() {
    const rows = document.querySelectorAll("#excelTable tbody tr");

    rows.forEach(row => {
        const rowId = row.dataset.index;

        // Right-click for context menu
        row.oncontextmenu = (e) => {
            e.preventDefault();
            showContextMenu(e.pageX, e.pageY, rowId);
        };
    });
}


function updateRanks() {
    const rows = document.querySelectorAll("#excelTable tbody tr");
    rows.forEach((row, index) => {
        const rankCell = row.querySelector(".rank-cell");
        if (rankCell) rankCell.textContent = index + 1;

        // Update the data-index attribute to match the new rank
        row.dataset.index = index + 1;
    });
}

function makeColumnsResizable(table) {
    const ths = table.querySelectorAll("th");

    ths.forEach((th) => {
        const resizer = document.createElement("div");
        resizer.classList.add("resizer");
        th.appendChild(resizer);

        let x = 0;
        let w = 0;

        resizer.addEventListener("mousedown", (e) => {
            x = e.clientX;
            w = th.offsetWidth;

            document.addEventListener("mousemove", onMouseMove);
            document.addEventListener("mouseup", onMouseUp);
        });

        function onMouseMove(e) {
            const dx = e.clientX - x;
            th.style.width = `${w + dx}px`;
        }

        function onMouseUp() {
            document.removeEventListener("mousemove", onMouseMove);
            document.removeEventListener("mouseup", onMouseUp);
        }
    });
}

function renderTable(data) {
    const thead = document.querySelector("#excelTable thead");
    const tbody = document.querySelector("#excelTable tbody");

    thead.innerHTML = "";
    tbody.innerHTML = "";

    if (data.length === 0) return;

    // Render headers
    const headerRow = document.createElement("tr");
    const rankHeader = document.createElement("th");
    rankHeader.textContent = "Rank";
    headerRow.appendChild(rankHeader);

    data[0].forEach((col) => {
        const th = document.createElement("th");
        th.textContent = col;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    // Render data rows
    for (let i = 1; i < data.length; i++) {
        const row = document.createElement("tr");
        row.setAttribute("draggable", true);
        row.dataset.index = i;

        const rankCell = document.createElement("td");
        rankCell.textContent = i;
        rankCell.classList.add("rank-cell");
        row.appendChild(rankCell);

        data[i].forEach((cell) => {
            const td = document.createElement("td");
            td.textContent = cell;
            row.appendChild(td);
        });

        // Add context menu event to row
        row.addEventListener("contextmenu", (e) => {
            e.preventDefault();
            showContextMenu(e.pageX, e.pageY, row.dataset.index);
        });

        tbody.appendChild(row);
    }

    enableRowDragDrop();
    makeColumnsResizable(document.getElementById("excelTable"));
}

function enableRowDragDrop() {
    const tbody = document.querySelector("#excelTable tbody");
    let dragSrc = null;

    tbody.querySelectorAll("tr").forEach((row) => {
        row.addEventListener("dragstart", () => {
            dragSrc = row;
            row.classList.add("dragging");
        });

        row.addEventListener("dragend", () => {
            row.classList.remove("dragging");
            tbody.querySelectorAll("tr").forEach((r) => r.classList.remove("drag-over"));
        });

        row.addEventListener("dragover", (e) => {
            e.preventDefault();
            if (row !== dragSrc) {
                row.classList.add("drag-over");
            }
        });

        row.addEventListener("dragleave", () => {
            row.classList.remove("drag-over");
        });

        row.addEventListener("drop", (e) => {
            e.preventDefault();
            if (dragSrc && dragSrc !== row) {
                const rows = Array.from(tbody.children);
                const srcIndex = rows.indexOf(dragSrc);
                const targetIndex = rows.indexOf(row);
                tbody.insertBefore(dragSrc, srcIndex < targetIndex ? row.nextSibling : row);
                updateRanks();
            }
        });
    });
}

document.getElementById("saveBtn").addEventListener("click", () => {
    const tbody = document.querySelector("#excelTable tbody");
    const rows = Array.from(tbody.rows);

    const newData = [
        currentData[0], // original headers (without "Rank")
        ...rows.map((row) =>
            Array.from(row.cells).slice(1).map((cell) => cell.textContent)
        )
    ];

    const ws = XLSX.utils.aoa_to_sheet(newData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });

    const blob = new Blob([wbout], { type: "application/octet-stream" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "reordered_data.xlsx";
    a.click();
});

// Create the context menu when the page loads
document.addEventListener("keydown", (e) => {
    if (e.ctrlKey && e.key === "z") {
        undoAction();
    }
    if (e.ctrlKey && e.key === "y") {
        redoAction();
    }
});

document.addEventListener("DOMContentLoaded", createContextMenu);
let rankPopupElements = createRankPopup();
hideRankPopup()

