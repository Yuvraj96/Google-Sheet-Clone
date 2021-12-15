// INITIALIZATIONS
let currentColHeaderIndex,
  currentRowHeaderIndex,
  currentColInputIndex,
  currentRowInputIndex;
const sumRegex = /^=SUM\([A-Z0-9]+\:[A-Z0-9]+\)$/i;
const sumReplaceRegex = /=SUM\(|\)/gi;
const multiplyRegex = /^=MULTIPLY\([A-Z0-9]+\,[A-Z0-9]+\)$/i;
const multiplyReplaceRegex = /=MULTIPLY\(|\)/gi;

// Update array data when user types in contentEditable element
app.addEventListener("keyup", (e) => {
  const target = e.target;
  if (target.contentEditable) {
    const row = target.dataset.row;
    const col = target.dataset.col;
    spreadSheet[row][col] = target.textContent;
    // console.log("UPDATE SPREADSHEET", spreadSheet);
    if (e.key === "Enter" || e.keyCode === 13) {
      if (sumRegex.test(spreadSheet[row][col])) {
        currentRowInputIndex = row;
        currentColInputIndex = col;
        let sumIndex = spreadSheet[row][col]
          .replace(sumReplaceRegex, "")
          .split(":");
        spreadSheet[row][col] = getSum(sumIndex);
        const cell = document
          .querySelectorAll(".row-parent")
          [row].querySelectorAll('span[contentEditable="true"]')[col];
        cell.textContent = spreadSheet[row][col];
      } else if (multiplyRegex.test(spreadSheet[row][col])) {
        currentRowInputIndex = row;
        currentColInputIndex = col;
        let multiplyIndex = spreadSheet[row][col]
          .replace(multiplyReplaceRegex, "")
          .split(",");
        spreadSheet[row][col] = getMultiply(multiplyIndex);
        const cell = document
          .querySelectorAll(".row-parent")
          [row].querySelectorAll('span[contentEditable="true"]')[col];
        cell.textContent = spreadSheet[row][col];
      }
    }
  }
});

/**
 * Returns sum from applied formula.
 *
 * @param {Object} index Array of cell position.
 * @return {Number|String} sum of cells or error.
 */
function getSum(index) {
  // console.log(index);
  const range = hasValidSumRange(index);
  if (range.valid) {
    const { result } = range;
    // console.log(result);
    let sum = spreadSheet
      .slice(result.rowStart, result.rowEnd + 1)
      .map((i) => i.slice(result.colStart, result.colEnd + 1))
      .flat();
    // console.log(sum);
    return sum.reduce(
      (acc, curr) => (!isNaN(curr) ? acc + Number(curr) : acc),
      0
    );
  }
  return range.result;
}

/**
 * Returns validity of sum formula.
 *
 * @param {Object} index Array of cell position.
 * @return {Object} validity object.
 */
function hasValidSumRange(index) {
  //A:5, 8:AB
  if (!isNaN(index[0]) || !isNaN(index[1]))
    return { valid: false, result: "#ERROR!" };

  const colStartValue = index[0].match(/[A-Z]+/gi);
  const colEndValue = index[1].match(/[A-Z]+/gi);
  // console.log(colStartValue, colEndValue);
  //A1B:U, V:Z9B, 1AB:U, U:1AB
  if (
    colStartValue.length > 1 ||
    colEndValue.length > 1 ||
    /^[0-9]+[A-Z]+$/i.test(index[0]) ||
    /^[0-9]+[A-Z]+$/i.test(index[1]) ||
    /^[0-9]+[A-Z0-9]+$/i.test(index[0]) ||
    /^[0-9]+[A-Z0-9]+$/i.test(index[1])
  )
    return { valid: false, result: "#NAME!" };

  let rowStart = index[0].match(/\d/g);
  //A0:B
  if (rowStart) {
    rowStart = rowStart.join("");
    if (Number(rowStart) === 0) return { valid: false, result: "#ERROR!" };
    rowStart = Number(rowStart) <= rowSize ? Number(rowStart) - 1 : rowSize - 1;
  } else {
    rowStart = rowSize - 1;
  }

  let rowEnd = index[1].match(/\d/g);
  //A:B0
  if (rowEnd) {
    rowEnd = rowEnd.join("");
    if (Number(rowEnd) === 0) return { valid: false, result: "#ERROR!" };
    rowEnd = Number(rowEnd) <= rowSize ? Number(rowEnd) - 1 : rowSize - 1;
  } else {
    rowEnd = rowSize - 1;
  }
  if (rowStart > rowEnd) [rowStart, rowEnd] = [rowEnd, rowStart];
  else if (rowStart === rowSize - 1 && rowStart === rowEnd) rowStart = 0;

  let colStart = alphaToIndex(colStartValue[0]);
  colStart = colStart <= colSize ? colStart : colSize - 1;
  let colEnd = alphaToIndex(colEndValue[0]);
  colEnd = colEnd <= colSize ? colEnd : colSize - 1;
  if (colStart > colEnd) [colStart, colEnd] = [colEnd, colStart];
  if (
    currentRowInputIndex >= rowStart &&
    currentRowInputIndex <= rowEnd &&
    currentColInputIndex >= colStart &&
    currentColInputIndex <= colEnd
  )
    return { valid: false, result: "#REF!" };
  return {
    valid: true,
    result: {
      rowStart: rowStart,
      rowEnd: rowEnd,
      colStart: colStart,
      colEnd: colEnd,
    },
  };
}

/**
 * Returns multiplication from applied formula.
 *
 * @param {Object} index Array of cell position.
 * @return {Number|String} multiplication of cells or error.
 */
function getMultiply(index) {
  // console.log(index);
  const range = hasValidRangeMultiply(index);
  if (range.valid) {
    const { result } = range;
    // console.log(result);
    let multiply = result.num1 * result.num2;
    // console.log(multiply);
    return multiply;
  }
  return range.result;
}

/**
 * Returns validity of multiply formula.
 *
 * @param {Object} index Array of cell position.
 * @return {Object} validity object.
 */
function hasValidRangeMultiply(index) {
  const colStartValue = index[0].match(/[A-Z]+/gi);
  const colEndValue = index[1].match(/[A-Z]+/gi);
  // console.log(colStartValue, colEndValue);
  //A1B:U, V:Z9B, 1AB:U, U:1AB, A:B
  if (
    (colStartValue && colStartValue.length > 1) ||
    (colEndValue && colEndValue.length > 1) ||
    /^[0-9]+[A-Z]+$/i.test(index[0]) ||
    /^[0-9]+[A-Z]+$/i.test(index[1]) ||
    /^[A-Z]+$/i.test(index[0]) ||
    /^[A-Z]+$/i.test(index[1])
  )
    return { valid: false, result: "#NAME!" };

  let num1, num2;
  if (!isNaN(index[0])) num1 = Number(index[0]);
  else {
    let rowStart = index[0].match(/\d/g);
    //A0:B
    rowStart = rowStart.join("");
    if (Number(rowStart) === 0) return { valid: false, result: "#ERROR!" };

    // console.log("rowStart", rowStart);
    if (Number(rowStart) <= rowSize) {
      rowStart = Number(rowStart) - 1;
      let colStart = alphaToIndex(colStartValue[0]);
      if (colStart > colSize) num1 = 0;
      else {
        if (
          Number(currentRowInputIndex) === rowStart &&
          Number(currentColInputIndex) === colStart
        )
          return { valid: false, result: "#REF!" };
        num1 = spreadSheet[rowStart][colStart];
      }
      if (isNaN(num1)) return { valid: false, result: "#VALUE!" };
    } else {
      num1 = 0;
    }
  }

  if (!isNaN(index[1])) num2 = Number(index[1]);
  else {
    let rowEnd = index[1].match(/\d/g);
    //A:B0
    rowEnd = rowEnd.join("");
    if (Number(rowEnd) === 0) return { valid: false, result: "#ERROR!" };

    // console.log("rowEnd", rowEnd);
    if (Number(rowEnd) <= rowSize) {
      rowEnd = Number(rowEnd) - 1;
      let colEnd = alphaToIndex(colEndValue[0]);
      if (colEnd > colSize) num2 = 0;
      else {
        if (
          Number(currentRowInputIndex) === rowEnd &&
          Number(currentColInputIndex) === colEnd
        )
          return { valid: false, result: "#REF!" };
        num2 = spreadSheet[rowEnd][colEnd];
      }
      if (isNaN(num2)) return { valid: false, result: "#VALUE!" };
    } else {
      num2 = 0;
    }
  }

  return {
    valid: true,
    result: {
      num1: num1,
      num2: num2,
    },
  };
}

/**
 * Returns index of column header value.
 *
 * @param {String} value The column header value.
 * @return {Number} index of column header value.
 */
function alphaToIndex(value) {
  value = value.toUpperCase();
  let index = value.charCodeAt(0) - 65;
  for (let i = 1; i < value.length; i++) {
    index += 26;
    index += value.charCodeAt(i) - 65;
  }
  // console.log("index ", index);
  return index;
}

// Detect right click on row/column headers
app.addEventListener(
  "contextmenu",
  (e) => {
    e.preventDefault();
    const target = e.target;
    const clickedRowHeader = target.classList.contains("row-header");
    const clickedColHeader = target.classList.contains("col-header");
    const clickedCells = target.hasAttribute("contentEditable");
    if (clickedRowHeader || clickedColHeader || clickedCells) {
      let popup;
      if (clickedRowHeader) {
        popup = document.querySelector(".row-pop");
        currentRowHeaderIndex = target.dataset.rowParent;
      } else if (clickedColHeader) {
        popup = document.querySelector(".col-pop");
        currentColHeaderIndex = target.dataset.colParent;
      } else {
        popup = document.querySelector(".cell-pop");
        currentRowHeaderIndex = target.dataset.row;
        currentColHeaderIndex = target.dataset.col;
      }
      closePopup();
      setPopupPosition(e, popup);
      popup.classList.add("show");
    }
    return false;
  },
  false
);

// Detect left click after dropdown opens and execute functionalities
app.addEventListener("click", (e) => {
  const target = e.target;
  const popupOption = target.parentNode.classList.contains("popup");
  if (popupOption) {
    if (target.classList.contains("insert1")) insertRow(true);
    else if (target.classList.contains("insert2")) insertRow();
    else if (target.classList.contains("insert3")) insertCol(true);
    else if (target.classList.contains("insert4")) insertCol();
    else if (target.classList.contains("sort1")) sortColumn(true);
    else if (target.classList.contains("sort2")) sortColumn();
  }
  closePopup();
});

/**
 * Close the popup that are opened
 */
function closePopup() {
  const popups = document.querySelectorAll(".popup.show");
  for (let i = 0; i < popups.length; i++) {
    popups[i].classList.remove("show");
  }
}

/**
 * Set position of popup based on user's mouse position
 *
 * @param {Object} e The event object.
 * @param {Object} element The popup element to be modified.
 */
function setPopupPosition(e, element) {
  const x = e.clientX,
    y = e.clientY;
  element.style.top = y + 5 + "px";
  element.style.left = x + 5 + "px";
}

/**
 * Sort the multidimensional array in ascending/descending order
 *
 * @param {Boolean} asc Checks weather column needs to be sorted ascending.
 */
function sortColumn(asc = false) {
  // console.log("sortedIndex ", currentColHeaderIndex);
  const isAsc = asc ? 1 : -1;
  spreadSheet.sort((a, b) => {
    a = a[currentColHeaderIndex];
    b = b[currentColHeaderIndex];
    if (a === b) return 0;
    else if (a === "") return 1;
    else if (b === "") return -1;
    else return a < b ? -isAsc : isAsc;
  });
  // console.log("AFTER SORT");
  updateUI();
}

/**
 * Insert a column left/right a specified column
 *
 * @param {Boolean} left Checks weather column needs to be added on left of the specified column.
 */
function insertCol(left = false) {
  const insertIndex = left
    ? currentColHeaderIndex
    : Number(currentColHeaderIndex) + 1;
  // console.log("insertColIndex ", insertIndex);

  const getColumn = document.querySelector(".col-parent");
  const newColHeader = createColCell(colSize);
  getColumn.append(newColHeader);
  const getAllRows = document.querySelectorAll(".row-parent");
  for (let row = 0; row < rowSize; row++) {
    const newCol = createRowCell(row, colSize);
    getAllRows[row].append(newCol);
  }

  colSize++;
  spreadSheet.forEach((row) => {
    row.splice(insertIndex, 0, "");
  });
  updateUI();
}

/**
 * Insert a row above/below a specified row
 *
 * @param {Boolean} top Checks weather row needs to be added on top of the specified row.
 */
function insertRow(top = false) {
  const insertIndex = top
    ? currentRowHeaderIndex
    : Number(currentRowHeaderIndex) + 1;
  // console.log("insertRowIndex ", insertIndex);

  const newRow = createRow(rowSize);
  app.appendChild(newRow);

  rowSize++;
  spreadSheet.splice(insertIndex, 0, Array(colSize).fill(""));
  updateUI();
}

/**
 * Update UI to incorporate new changes(Add/Delete rows and columns, sorting)
 */
function updateUI() {
  const allCells = document.querySelectorAll(
    '.row-parent span[contentEditable="true"]'
  );
  allCells.forEach((cell) => {
    const row = cell.dataset.row;
    const col = cell.dataset.col;
    cell.textContent = spreadSheet[row][col];
  });
  // console.log(spreadSheet);
}
