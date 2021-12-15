// INITIALIZATIONS
const app = document.getElementById("app");
let rowSize = 50;
let colSize = 26;
let spreadSheet = [...Array(rowSize)].map((x) => Array(colSize).fill(""));
const columnPopup = [
  { key: "insert3", value: "Insert 1 Column to the Left" },
  { key: "insert4", value: "Insert 1 Column to the Right" },
  { key: "sort1", value: "Sort Sheet A to Z" },
  { key: "sort2", value: "Sort Sheet Z to A" },
];
const rowPopup = [
  { key: "insert1", value: "Insert 1 Row above" },
  { key: "insert2", value: "Insert 1 Row below" },
];
const cellPopup = rowPopup.concat(columnPopup.slice(0, 2));
/**
 * Utility function to add multiple attributes
 *
 * @param {Object} x The element on which attributes are to be added.
 * @param {Object} attrs The attributes to be added to the element.
 */
function setAttributes(el, attrs) {
  for (const key in attrs) {
    el.setAttribute(key, attrs[key]);
  }
}

/**
 * Create initial spreadSheet Layout with default row,col values.
 */
(function createInitialLayout() {
  const fragment = document.createDocumentFragment();
  const newCol = createColumn();
  fragment.appendChild(newCol);
  for (let row = 0; row < rowSize; row++) {
    const newRow = createRow(row);
    fragment.appendChild(newRow);
  }
  app.appendChild(fragment);
})();

/**
 * Create HTML element of entire column of the spreadsheet.
 *
 * @return {Object} HTML element of column.
 */
function createColumn() {
  const colParent = document.createElement("div");
  setAttributes(colParent, { class: "col-parent cell-parent" });
  const colParentFragment = document.createDocumentFragment();
  const colHeader = createRowHeader(-1);
  colParentFragment.appendChild(colHeader);
  for (let col = 0; col < colSize; col++) {
    const cell = createColCell(col);
    colParentFragment.appendChild(cell);
  }
  colParent.appendChild(colParentFragment);
  return colParent;
}

/**
 * Returns HTML element of column of specified index.
 *
 * @param {number} col The column index.
 * @return {Object} column of specified index.
 */
function createColCell(col) {
  //   console.log("create column at index ", col);
  const cell = document.createElement("div");
  setAttributes(cell, { class: "col-header", "data-col-parent": col });
  cell.textContent = indexToAlpha(col);
  return cell;
}

/**
 * Returns Column header value of specified index.
 *
 * @param {number} index The column index.
 * @return {String} header value of specified index.
 */
function indexToAlpha(index) {
  let value = "",
    isMultiChar = index > 25;
  while (index > 25) {
    value += String.fromCharCode((index % 26) + 65);
    index = Math.floor(index / 26);
  }
  index = isMultiChar ? index - 1 : index;
  value += String.fromCharCode(index + 65);
  return value.split("").reverse().join("");
}

/**
 * Returns HTML element of a row at specified index.
 *
 * @param {number} row The row index.
 * @return {Object} HTML element of a whole row.
 */
function createRow(row) {
  //   console.log("Create row at index ", row);
  const rowParent = document.createElement("div");
  setAttributes(rowParent, { class: "row-parent cell-parent" });
  const rowParentFragment = document.createDocumentFragment();
  const rowHeader = createRowHeader(row);
  rowParentFragment.appendChild(rowHeader);
  for (let col = 0; col < colSize; col++) {
    const cell = createRowCell(row, col);
    rowParentFragment.appendChild(cell);
  }
  rowParent.appendChild(rowParentFragment);
  return rowParent;
}

/**
 * Returns HTML element of row header of specified index.
 *
 * @param {number} row The row index.
 * @return {Object} row header of specified index.
 */
function createRowHeader(row) {
  //   console.log("Start row header at index ", row);
  const cell = document.createElement("div");
  setAttributes(cell, { class: "row-header", "data-row-parent": row });
  if (row >= 0) cell.textContent = Number(row) + 1;
  return cell;
}

/**
 * Returns HTML element of row of specified index.
 *
 * @param {number} row The row index.
 * @param {number} col The column index.
 * @return {Object} row of specified index.
 */
function createRowCell(row, col) {
  //   console.log("Start row cell at ", row, " ", col);
  const cell = document.createElement("span");
  setAttributes(cell, {
    contentEditable: "true",
    "data-row": row,
    "data-col": col,
  });
  return cell;
}

/**
 * Create popup element that we need on right click
 *
 * @param {Object} options array of object used to build our HTML element.
 * @param {String} cls custom className which is added to uniquely identify popup.
 */
function createPopup(options, cls) {
  const fragment = document.createDocumentFragment();
  const listParent = document.createElement("ul");
  setAttributes(listParent, { class: "popup " + cls });

  for (let i = 0; i < options.length; i++) {
    const list = document.createElement("li");
    setAttributes(list, { class: options[i].key });
    list.textContent = options[i].value;
    fragment.appendChild(list);
  }
  listParent.appendChild(fragment);
  app.appendChild(listParent);
}
createPopup(columnPopup, "col-pop");
createPopup(rowPopup, "row-pop");
createPopup(cellPopup, "cell-pop");
