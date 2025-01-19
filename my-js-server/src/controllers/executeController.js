// src/controllers/executeController.js
const executeService = require('../services/executeService');
const { PythonShell } = require('python-shell');
const path = require('path');
const util = require('util');
const { exec } = require('child_process');



// Convert exec into a promise-based function for async/await
const execPromise = util.promisify(exec);
// Define the path to the C executable
const exePath = path.join('C:', 'Users', 'abarb', 'source', 'repos', 'mediaSorter', 'x64', 'Debug', 'mediaSorter.exe');


const allColumnTypes = {
  CONDITIONS: 1,
  ATTRIBUTES: 2
};

const msgType = {
  ERROR: 0,
  WARNING: 1,
  RELATIVES: 3
};

const Actions = {
  Select: 1,
  NewVal: 2,
  NewBgCol: 3
}


/**
 * Converts a given 0-based integer to its corresponding Excel column tag.
 * @param {number} num - The integer to be converted to a column tag.
 * @returns {string} The corresponding Excel column tag.
 */
function getColumnTag(num) {
  let columnTag = '';
  while (num >= 0) {
      columnTag = String.fromCharCode((num % 26) + 65) + columnTag;
      num = Math.floor(num / 26) - 1;
  }
  return columnTag;
}

function addListBoxList(params) {
  let values = params.values;
  let i = params.i;
  let j = params.j;
  let k = params.k;
  let val = params.val;
  let columnTitle = params.columnTitle;
  let message = params.message;
  let suggs = params.suggs;
  let msgTypeColors = params.msgTypeColors;
  let values0Glob = params.values0Glob;
  let valuesGlob = params.valuesGlob;
  let prevLine = params.prevLine;
  let lenAgg = params.lenAgg;
  let attributes = params.attributes;
  let sorting = params.sorting;
  let sheetCodeName = params.sheetCodeName;
  let editRow = params.editRow;
  let editCol = params.editCol;
  let resolved = params.resolved;
  let oldVersionsMaxNb = params.oldVersionsMaxNb;
  let columnTypes = params.columnTypes;
  let data = params.data;
  let allMediaRows = params.allMediaRows;
  let attrOfAttr = params.attrOfAttr;
  let nameSt = params.nameSt;
  let nameInd = params.nameInd;
  let mediaSt = params.mediaSt;
  let mediaInd = params.mediaInd;
  let headerColors = params.headerColors;
  let attNames = params.attNames;
  let actionsHistory = params.actionsHistory;
  let indActHist = params.indActHist;
  let handleChanges = params.handleChanges;
  let rowIdMap = params.rowIdMap;
  let response = params.response;
  if(response.length === 0) {
    response.push({});
  }
  if(response[0]["listBoxList"] === undefined) {
    response[0]["listBoxList"] = [];
    for (const key in msgType) {
      response[0]["listBoxList"].push([]);
    }
  }
}

/**
 *  Update the links to show at the menu each time a cell is selected
 * 
 */
function handleSelectLinks(params) {
  // let values = params.values;
  let i = params.i;
  let j = params.j;
  let k = params.k;
  let val = params.val;
  let columnTitle = params.columnTitle;
  let message = params.message;
  let suggs = params.suggs;
  let msgTypeColors = params.msgTypeColors;
  let values0Glob = params.values0Glob;
  let valuesGlob = params.valuesGlob;
  let prevLine = params.prevLine;
  let lenAgg = params.lenAgg;
  let attributes = params.attributes;
  let sorting = params.sorting;
  let sheetCodeName = params.sheetCodeName;
  let editRow = params.editRow;
  let editCol = params.editCol;
  let resolved = params.resolved;
  let oldVersionsMaxNb = params.oldVersionsMaxNb;
  let columnTypes = params.columnTypes;
  let data = params.data;
  let allMediaRows = params.allMediaRows;
  let attrOfAttr = params.attrOfAttr;
  let nameSt = params.nameSt;
  let nameInd = params.nameInd;
  let mediaSt = params.mediaSt;
  let mediaInd = params.mediaInd;
  let headerColors = params.headerColors;
  let attNames = params.attNames;
  let actionsHistory = params.actionsHistory;
  let indActHist = params.indActHist;
  let handleChanges = params.handleChanges;
  let rowIdMap = params.rowIdMap;
  let response = params.response;
  // only if no errors in the sheet
  if(!resolved[sheetCodeName]) {
    return -1;
  }

  let row = editRow[sheetCodeName];
  let column = editCol[sheetCodeName];
  let values = valuesGlob[sheetCodeName];
  let nbLineBef = values.length;
  if(nbLineBef == 0) {
    return -1;
  }
  let colNumb = values[0].length;
  if(row>=nbLineBef || column >= colNumb) {
    return -1;
  }
  
  if(column==0) {
    if(perRef[row] > -1) {
      for(let g = 0; g < 3; g++) {
        if(periods[perRef[row]][g] < nbLineBef && periods[perRef[row]][g]!=row) {
          displayReference(values, periods[perRef[row]][g]);
        }
      }
    }
  } else {
    for(let k = 0; k < values[row][column].length; k++) {
      if(column < 4) {
        for(let r = 1; r < values.length; r++) {
          for(let f = 0; f < values[r][0].length; f++) {
            if(row!=r && searching(values, r, 0, f)) {
              if(perRef[r] > -1) {
                for(let g = 0; g < 3; g++) {
                  if(periods[perRef[r]][g]!=-1) {
                    displayReference(values, r);
                  }
                }
              } else {
                displayReference(values, r);
              }
            }
          }
        }
      } else if (column < colNumb - 1) {
        nbMediaInCat = attributes[data[row][2]].length;
        for(let r = 1; r < values.length; r++) {
          for(let d = 4; d < colNumb - 1; d++) {
            for(let f = 0; f < values[r][d].length; f++) {
              if(row!=r && searching(values, r, d, f)) {
                displayReference(values, r);
              }
            }
          }
        }
      }
    }
  }
}

function addToDic(dic, key, value) {
  if(dic[key] === undefined) {
    dic[key] = [];
  }
  dic[key].push(value);
}

function addToSupDic(dic, key1, key2, value) {
  addToDic(dic, key1, {});
  addToDic(dic[key1], key2, value);
}

async function handleChange(params, updates, sheetCodeName) {
  // let values = params.values;
  let i = params.i;
  let j = params.j;
  let k = params.k;
  let val = params.val;
  let columnTitle = params.columnTitle;
  let message = params.message;
  let suggs = params.suggs;
  let msgTypeColors = params.msgTypeColors;
  let values0Glob = params.values0Glob;
  let valuesGlob = params.valuesGlob;
  let prevLine = params.prevLine;
  let lenAgg = params.lenAgg;
  let attributes = params.attributes;
  let sorting = params.sorting;
  let sheetCodeName = params.sheetCodeName;
  let editRow = params.editRow;
  let editCol = params.editCol;
  let resolved = params.resolved;
  let oldVersionsMaxNb = params.oldVersionsMaxNb;
  let columnTypes = params.columnTypes;
  let data = params.data;
  let allMediaRows = params.allMediaRows;
  let attrOfAttr = params.attrOfAttr;
  let nameSt = params.nameSt;
  let nameInd = params.nameInd;
  let mediaSt = params.mediaSt;
  let mediaInd = params.mediaInd;
  let headerColors = params.headerColors;
  let attNames = params.attNames;
  let actionsHistory = params.actionsHistory;
  let indActHist = params.indActHist;
  let handleChanges = params.handleChanges;
  let rowIdMap = params.rowIdMap;
  let response = params.response;
  
  let values = valuesGlob[sheetCodeName];
  let values0 = values0Glob[sheetCodeName];
  for(let i = 0; i < updates.length; i++) {
    var row = updates[i][0] - 1;  // Adjusting for zero-based index
    var column = updates[i][1] - 1;  // Adjusting for zero-based index
    var value = updates[i][2];
    if(value) {
      // add rows if needed
      while(values.length <= row) {
        values.push([]);
        values0.push([]);
      }

      // add columns if needed
      if(values[0].length <= column) {
        while(headerColors.length <= column) {
          headerColors.push(null);
          columnTypes.push(null);
        }
        for (let i = 0; i < values.length; i++) {
          while (values[i].length <= column) {
            values[i].push([]);
            values0[i].push(null);
          }
        }
      }

      let splitValue = [];
      splitValue = value
          .split(";")
          .map(v => v.trim().toLowerCase())
          .filter(e => e !== "");
      values[row][column] = splitValue;
      values0[row][column] = value;
    } else {
      // the user has deleted a cell

      values[row][column] = [];
      values0[row][column] = value;
      let nbLineBef = values.length;
      let colNumb;
      if(nbLineBef > 0) {
        colNumb = values[0].length;
      } else {
        colNumb = 0;
      }

      // remove empty lines at the end
      const nbLineBefBef = nbLineBef;
      for (let i = nbLineBef - 1; i > 0; i--) {
        if (values[i].some(v => v.length != 0)) {
          break;
        }
        nbLineBef--;
      }
      values.splice(nbLineBef, nbLineBefBef - nbLineBef);
      values0.splice(nbLineBef, nbLineBefBef - nbLineBef);

      if(column === colNumb - 1) {
        let newColNumb = colNumb;
        let stop = false;
        for(let j = colNumb - 1; j > -1; j--) {
          for(let i = 0; i < nbLineBef; i++) {
            if(values[i][j].length !== 0) {
              stop = true;
              break;
            }
          }
          if(stop) {
            break;
          }
          newColNumb--;
        }
        if(newColNumb != colNumb) {
          for(let i = 0; i < nbLineBef; i++) {
            values[i].splice(newColNumb, colNumb - newColNumb);
            values0[i].splice(newColNumb, colNumb - newColNumb);
          }
        }
      }
    }
  }
  

  // Label the columns "names" and "media"
  headerNames = {};
  let namesColor;
  resolved[sheetCodeName] = false;
  nameInd = -1;
  mediaInd = -1;
  for(let j = 0; j < values[0].length; j++) {
    for(let k = 0; k < values[0][j].length; k++) {
      let name = values[0][j][k];
      let alreadyInd = [];
      if(name == nameSt) {
        if(nameInd != -1) {
          alreadyInd = [nameSt, nameInd];
        }
        nameInd = j;
        namesColor = headerColors[j];
      } else if(name == mediaSt) {
        if(mediaInd != -1) {
          alreadyInd = [mediaSt, mediaInd];
        }
        mediaInd = j;
      }
      if(alreadyInd.length) {
        inconsist_removing_elem(values, 0, j, k, `Column ${getColumnTag(j)} labelled as "${alreadyInd[0]}" already at column ${getColumnTag(alreadyInd[1])}.`);
        return;
      }
    }
  }
  if(nameInd == -1) {
    inconsist(0, values[0].length, `Missing column "${nameSt}".`, [[nameSt]]);
    return;
  }
  if(headerColors[nameInd] === 16777215) {
    inconsist(0, nameInd, `Column ${getColumnTag(nameInd)} labelled as "${nameSt}" must have a background color.`, [['']]);
    return;
  }
  if(nameInd == mediaInd) {
    inconsist(0, j, `Column ${getColumnTag(nameInd)} labelled as "${nameSt}" and "${mediaSt}".`, [['']]);
    return;
  }
  if(mediaInd != -1 && namesColor !== headerColors[mediaInd]) {
    inconsist(0, mediaInd, `Column ${getColumnTag(mediaInd)} must have the same background color as the attributes.`, [headerColors[nameInd]], Actions.NewBgCol);
    return;
  }
  columnTypes = [];
  attNames = {};
  let condColor;
  for (let j = 0; j < headerColors.length; j++) {
    if (headerColors[j] === namesColor) {
      columnTypes.push(allColumnTypes.ATTRIBUTES);
      attNames[j] = [];
    } else if (headerColors[j] === null) {
      columnTypes.push(null);
    } else if (condColor === undefined || headerColors[j] === condColor) {
      condColor = headerColors[j];
      columnTypes.push(allColumnTypes.CONDITIONS);
    } else {
      inconsist(0, j, `Third background color at column ${j + 1} (must be only two, for the conditions and the attributes)`, [condColor], true);
      return;
    }
  }



  await check(sheetCodeName);
  // oldValues.splice(indOldValues, oldValues.length - 1 - indOldValues);
  // oldValues.push(values);
  // if(oldVersionsMaxNb < oldValues.length) {
  //   oldValues.shift();
  // }
}

function getAddress(i, j) {
  return getColumnTag(j) + ":" + (i + 1);
}

function checkBrack(params) {
  let values = params.values;
  let i = params.i;
  let j = params.j;
  let k = params.k;
  let val = params.val;
  let columnTitle = params.columnTitle;
  let message = params.message;
  let suggs = params.suggs;
  let msgTypeColors = params.msgTypeColors;
  let values0Glob = params.values0Glob;
  let valuesGlob = params.valuesGlob;
  let prevLine = params.prevLine;
  let lenAgg = params.lenAgg;
  let attributes = params.attributes;
  let sorting = params.sorting;
  let sheetCodeName = params.sheetCodeName;
  let editRow = params.editRow;
  let editCol = params.editCol;
  let resolved = params.resolved;
  let oldVersionsMaxNb = params.oldVersionsMaxNb;
  let columnTypes = params.columnTypes;
  let data = params.data;
  let allMediaRows = params.allMediaRows;
  let attrOfAttr = params.attrOfAttr;
  let nameSt = params.nameSt;
  let nameInd = params.nameInd;
  let mediaSt = params.mediaSt;
  let mediaInd = params.mediaInd;
  let headerColors = params.headerColors;
  let attNames = params.attNames;
  let actionsHistory = params.actionsHistory;
  let indActHist = params.indActHist;
  let handleChanges = params.handleChanges;
  let rowIdMap = params.rowIdMap;
  let response = params.response;
  val = val.toLowerCase();
  let isAtt = -2;
  if(columnTitle !== undefined) {
    let columnInd = -1;
    for(let n = 0; n < colNumb; n++) {
      if(columnTypes[n] == allColumnTypes.ATTRIBUTES && values[0][n].includes(columnTitle)) {
        columnInd = n;
        break;
      }
    }
    if(columnInd == -1) {
      inconsist_removing_elem(values, i, j, k, `No attribute column named "${columnTitle}" at ${getAddress(i, j)}.`)
      return -1;
    }
    for(let m = 0; m < attNames[columnInd].length; m++) {
      if(attNames[columnInd][m] == val) {
        isAtt = m;
        break;
      }
    }
    if(isAtt == -2) {
      inconsist_replacing_elem(values, i, j, k, `The column named "${columnTitle}" at ${getAddress(i, j)} doesn't have the attribute "${val}".`)
      return -1;
    }
  } else {
    for(let key in attNames) {
      for(let m = 0; m < attNames[key].length; m++) {
        if(attNames[key][m] == val) {
          if(isAtt != -2) {
            inconsist_replacing_elem(values, i, nameInd, k, findNewName(values, val, false), `The attribute at ${getAddress(i, nameInd)} exists in multiple columns.`);
            return -1;
          }
          isAtt = m;
          break;
        }
      }
    }
  }
  return isAtt;
}

function initializeData(data, nbLineBef) {
  for (let i = 1; i < nbLineBef; i++) {
    data.push({attributes: new Set(), posteriors: [], ulteriors: [], nbPost: 0, minDist: 0, maxDist: Infinity});
  }
  data = data.filter(e => e !== -1);
}

function getAttributes(params) {
  let values = params.values;
  let i = params.i;
  let j = params.j;
  let k = params.k;
  let val = params.val;
  let columnTitle = params.columnTitle;
  let message = params.message;
  let suggs = params.suggs;
  let msgTypeColors = params.msgTypeColors;
  let values0Glob = params.values0Glob;
  let valuesGlob = params.valuesGlob;
  let prevLine = params.prevLine;
  let lenAgg = params.lenAgg;
  let attributes = params.attributes;
  let sorting = params.sorting;
  let sheetCodeName = params.sheetCodeName;
  let editRow = params.editRow;
  let editCol = params.editCol;
  let resolved = params.resolved;
  let oldVersionsMaxNb = params.oldVersionsMaxNb;
  let columnTypes = params.columnTypes;
  let data = params.data;
  let allMediaRows = params.allMediaRows;
  let attrOfAttr = params.attrOfAttr;
  let nameSt = params.nameSt;
  let nameInd = params.nameInd;
  let mediaSt = params.mediaSt;
  let mediaInd = params.mediaInd;
  let headerColors = params.headerColors;
  let attNames = params.attNames;
  let actionsHistory = params.actionsHistory;
  let indActHist = params.indActHist;
  let handleChanges = params.handleChanges;
  let rowIdMap = params.rowIdMap;
  let response = params.response;

  let colNumb = params.colNumb;
  // attributes
  let accAttr = {};
  let acc = 0;
  for (let j = 0; j < colNumb; j++) {
    let colTitle = values0Glob[sheetCodeName][0][j];
    if(colTitle !== null){
      try{
        colTitle = colTitle.split(";");
      } catch (error) {
        return;
      }
    } else {
      colTitle = [];
    }
    for(let k = 0; k < colTitle.length; k++) {
      if(colTitle[k].match(/^[A-Z]+$/)) {
        inconsist_replacing_elem(values, 0, j, k, findNewName(values, colTitle[k]), `Column name at ${getAddress(0, j)} is only upper case letters.`);
        return -1;
      }
    }
    if (columnTypes[j] == allColumnTypes.ATTRIBUTES && j!=nameInd) {
      accAttr[j] = acc;
      for (let i = 1; i < nbLineBef; i++) {
        for (let k = 0; k < values[i][j].length; k++) {
          let val = values[i][j][k];
          let attInd = attNames[j].length;
          for(let r = 0; r < attInd; r++) {
            if(attNames[j][r] == val) {
              attInd = r;
              break;
            }
          }
          if(attInd == attNames[j].length) {
            attNames[j].push(val);
            attrOfAttr.push(new Set());
            attributes.push([]);
            attInd += accAttr[j];
            acc+=1;
          }
          if(!data[i]) {
            continue;
          }
          data[i].attributes.add(attInd);
        }
      }
    }
  }
}

function getNames(params) {
  let values = params.values;
  let i = params.i;
  let j = params.j;
  let k = params.k;
  let val = params.val;
  let columnTitle = params.columnTitle;
  let message = params.message;
  let suggs = params.suggs;
  let msgTypeColors = params.msgTypeColors;
  let values0Glob = params.values0Glob;
  let valuesGlob = params.valuesGlob;
  let prevLine = params.prevLine;
  let lenAgg = params.lenAgg;
  let attributes = params.attributes;
  let sorting = params.sorting;
  let sheetCodeName = params.sheetCodeName;
  let editRow = params.editRow;
  let editCol = params.editCol;
  let resolved = params.resolved;
  let oldVersionsMaxNb = params.oldVersionsMaxNb;
  let columnTypes = params.columnTypes;
  let data = params.data;
  let allMediaRows = params.allMediaRows;
  let attrOfAttr = params.attrOfAttr;
  let nameSt = params.nameSt;
  let nameInd = params.nameInd;
  let mediaSt = params.mediaSt;
  let mediaInd = params.mediaInd;
  let headerColors = params.headerColors;
  let attNames = params.attNames;
  let actionsHistory = params.actionsHistory;
  let indActHist = params.indActHist;
  let handleChanges = params.handleChanges;
  let rowIdMap = params.rowIdMap;
  let response = params.response;
  
  let colNumb = params.colNumb;
  
  allMediaRows = [];
  for(let i = 1; i < nbLineBef; i++) {
    let valuesI = values[i][nameInd];
    for(let k = 0; k < valuesI.length; k++) {
      let val = valuesI[k];
      const regexName = /(?<val>[^\[\]]+)(?:\s*\[\s*\"(?<columnTitle>[^\[\]]+)\s*\])?/g;
      let match = regexName.exec(val);
      let refInd_val;
      if(match !== null) {
        let params = {};
        params.values = values;
        params.i = i;
        params.j = j;
        params.k = k;
        params.val = match.groups.val;
        params.columnTitle = match.groups.columnTitle;
        refInd_val = checkBrack(params);
        if(refInd_val == -1) {
          return -1;
        }
        valuesI[k] = val + (match.groups.columnTitle===undefined?"":'[' + match.groups.columnTitle + ']');
      }
      if(match === null || refInd_val === -2) {
        // check if the name is already used
        for (let r = 1; r <= i; r++) {
          for (let f = 0; f < values[r][nameInd].length; f++) {
            if ((r < i || f < k) && valuesI[k]==values[r][nameInd][k]) {
              inconsist_replacing_elem(values, i, nameInd, k, findNewName(values, valuesI[k]), `name "${valuesI[k]}" at ${getAddress(i, nameInd)} already at ${getAddress(r, nameInd)}`);
              return -1;
            }
          }
        }
      } else {
        attrOfAttr[refInd_val] = data[i].attributes;
        if(data[i].isAtt !== undefined && data[i].isAtt !== refInd_val) {
          inconsist_removing_elem(values, i, nameInd, k, `attribute "${val}" at ${getAddress(i, nameInd)} is already linked to another name.`);
          return -1;
        }
        data[i].isAtt = refInd_val;
      }
    }
    if(data[i].isAtt === undefined) {
      allMediaRows.push(i);
    }
  }
}

function setAllAttr(params) {
  let values = params.values;
  let i = params.i;
  let j = params.j;
  let k = params.k;
  let val = params.val;
  let columnTitle = params.columnTitle;
  let message = params.message;
  let suggs = params.suggs;
  let msgTypeColors = params.msgTypeColors;
  let values0Glob = params.values0Glob;
  let valuesGlob = params.valuesGlob;
  let prevLine = params.prevLine;
  let lenAgg = params.lenAgg;
  let attributes = params.attributes;
  let sorting = params.sorting;
  let sheetCodeName = params.sheetCodeName;
  let editRow = params.editRow;
  let editCol = params.editCol;
  let resolved = params.resolved;
  let oldVersionsMaxNb = params.oldVersionsMaxNb;
  let columnTypes = params.columnTypes;
  let data = params.data;
  let allMediaRows = params.allMediaRows;
  let attrOfAttr = params.attrOfAttr;
  let nameSt = params.nameSt;
  let nameInd = params.nameInd;
  let mediaSt = params.mediaSt;
  let mediaInd = params.mediaInd;
  let headerColors = params.headerColors;
  let attNames = params.attNames;
  let actionsHistory = params.actionsHistory;
  let indActHist = params.indActHist;
  let handleChanges = params.handleChanges;
  let rowIdMap = params.rowIdMap;
  let response = params.response;
  
  let colNumb = params.colNumb;
  
  for(const i of allMediaRows) {
    let visited = new Set();
    for(let j of data[i].attributes) {
      const queue = [j]; // Initialize the queue with the start index
      
      // While the queue is not empty
      while (queue.length > 0) {
        const currentIndex = queue.shift();  // Get the next set index to explore
        
        // If we haven't visited this set yet
        if (!visited.has(currentIndex)) {
          // Mark this set as visited
          visited.add(currentIndex);
          
          // Add all the sets that this set points to into the queue
          for (let pointedIndex of attrOfAttr[currentIndex]) {
            if (!visited.has(pointedIndex)) {
              queue.push(pointedIndex);
            }
          }
        }
      }
    }
    data[i].attributes = visited;
  }
}

function addFormula(params, values, i, formula) {
  let values = params.values;
  let i = params.i;
  let j = params.j;
  let k = params.k;
  let val = params.val;
  let columnTitle = params.columnTitle;
  let message = params.message;
  let suggs = params.suggs;
  let msgTypeColors = params.msgTypeColors;
  let values0Glob = params.values0Glob;
  let valuesGlob = params.valuesGlob;
  let prevLine = params.prevLine;
  let lenAgg = params.lenAgg;
  let attributes = params.attributes;
  let sorting = params.sorting;
  let sheetCodeName = params.sheetCodeName;
  let editRow = params.editRow;
  let editCol = params.editCol;
  let resolved = params.resolved;
  let oldVersionsMaxNb = params.oldVersionsMaxNb;
  let columnTypes = params.columnTypes;
  let data = params.data;
  let allMediaRows = params.allMediaRows;
  let attrOfAttr = params.attrOfAttr;
  let nameSt = params.nameSt;
  let nameInd = params.nameInd;
  let mediaSt = params.mediaSt;
  let mediaInd = params.mediaInd;
  let headerColors = params.headerColors;
  let attNames = params.attNames;
  let actionsHistory = params.actionsHistory;
  let indActHist = params.indActHist;
  let handleChanges = params.handleChanges;
  let rowIdMap = params.rowIdMap;
  let response = params.response;
  
  let colNumb = params.colNumb;

  for (let j = 0; j < colNumb; j++) {
    if (columnTypes[j] == allColumnTypes.CONDITIONS) {
      let colPattern = values[0][j];
      if(colPattern.length) {
        colPattern = colPattern[0].split("?!");
      }
      if (values[i][j].length + (colPattern.length?1:-1) > colPattern.length) {
        inconsist_removing_elem(values, i, j, k, `Too much arguments at ${getAddress(i, j)}.`);
        return -1;
      }
      if(values[i][j].length) {
        let subForumla = values[i][j][0];
        if(colPattern.length) {
          subForumla = colPattern[0] + colPattern.slice(1).map((str, index) => values[i][j][index] + str).join('');
        }
        formula.push(subForumla);
      }
    }
  }
}

async function getConditions(params, values, sheetCodeName) {
  let values = params.values;
  let i = params.i;
  let j = params.j;
  let k = params.k;
  let val = params.val;
  let columnTitle = params.columnTitle;
  let message = params.message;
  let suggs = params.suggs;
  let msgTypeColors = params.msgTypeColors;
  let values0Glob = params.values0Glob;
  let valuesGlob = params.valuesGlob;
  let prevLine = params.prevLine;
  let lenAgg = params.lenAgg;
  let attributes = params.attributes;
  let sorting = params.sorting;
  let sheetCodeName = params.sheetCodeName;
  let editRow = params.editRow;
  let editCol = params.editCol;
  let resolved = params.resolved;
  let oldVersionsMaxNb = params.oldVersionsMaxNb;
  let columnTypes = params.columnTypes;
  let data = params.data;
  let allMediaRows = params.allMediaRows;
  let attrOfAttr = params.attrOfAttr;
  let nameSt = params.nameSt;
  let nameInd = params.nameInd;
  let mediaSt = params.mediaSt;
  let mediaInd = params.mediaInd;
  let headerColors = params.headerColors;
  let attNames = params.attNames;
  let actionsHistory = params.actionsHistory;
  let indActHist = params.indActHist;
  let handleChanges = params.handleChanges;
  let rowIdMap = params.rowIdMap;
  let response = params.response;
  
  let colNumb = params.colNumb;

  let optConds = false;
  for(const i of allMediaRows) {
    let formula = [];
    addFormula(values, i, formula);
    for (let a of data[i].attributes) {
      addFormula(values, a, formula);
    }
    formula = formula.join(" && ");
    let replacedWithWords;
    let reverseReplacements;
    try{
      ({replacedWithWords, reverseReplacements} = replaceWithSameWord(formula));
    } catch(error) {
      inconsist(i, 0, error.message, []);
      return;
    }
    
    let direct_terms = [];
    if(replacedWithWords!=='') {
      try {

        direct_terms = await runPythonScript("test_direct", replacedWithWords)
      } catch (err) {
        inconsist(i, 0, `incorrect condition : ${err.message}.`, []);
        return -1;
      }
    }
    if(!optConds) {
      const maxWord = direct_terms.reduce((max, item) => (item.value > max ? item.value : max), -Infinity);
      if(direct_terms.length <= maxWord) {
        optConds = true;
      }
    }
    
    // push in reverseReplacements groups for all the rows that have the aimed attributes
    for(let word in reverseReplacements) {
      if (reverseReplacements.hasOwnProperty(word)) {
        let groups = reverseReplacements[word][0];
        if(groups.precedent === undefined) {
          continue;
        }
        if(groups.attrib_ref !== undefined) {
          let params = {};
          params.values = values;
          params.i = i;
          params.j = 0;
          params.k = 0;
          params.val = groups.attrib_ref;
          params.columnTitle = groups.attrib_ref_col;
          let refInd_val = checkBrack(params);
          if(refInd_val == -1) {
            return -1;
          }
          if(refInd_val == -2) {
            inconsist(i, 0, `attribute "${val}" at row ${i + 1} not found.`, []);
            return -1;
          }
          groups.attrib_ref = refInd_val;
        }
        let params = {};
        params.values = values;
        params.i = i;
        params.j = 0;
        params.k = 0;
        params.val = groups.precedent;
        params.columnTitle = groups.precedent_col;
        let refInd_val = checkBrack(params);
        if(refInd_val == -1) {
          return -1;
        }
        groups.precedent_is_att = refInd_val !== -2;
        if(groups.precedent_is_att) {
          groups.precedent = refInd_val;
        } else {
          let found = false;
          for(let m = 1; m < values.length; m++) {
            for(let f = 0; f < values[m][nameInd].length; f++) {
              if(values[m][nameInd][f] == groups.precedent) {
                if(data[m].isAtt !== undefined) {
                  groups.precedent_is_att = true;
                  groups.precedent = data[m].isAtt;
                } else {
                  groups.precedent = m;
                }
                found = true;
                break;
              }
            }
            if(found) {
              break;
            }
          }
          if(!found) {
            inconsist(i, 0, `precedent "${groups.precedent}" at row ${i + 1} not found.`, []);
            return -1;
          }
        }
        if(groups.precedent_is_att) {
          for(let k = 0; k < attributes[groups.precedent].length; k++) {
            let rowHasAtt = attributes[groups.precedent][k];
            const copyElement = {
              ...groups
            };
            copyElement.precedent = rowHasAtt;
            copyElement.precedent_is_att = false;
            reverseReplacements[word].push(copyElement);
          }
        }
      }
    }
    
    for(let j = 0; j < direct_terms.length; j++) {
      for(let k = 0; k < reverseReplacements[direct_terms[j]].length; k++) {
        let groups = reverseReplacements[direct_terms[j]][k];
        if(groups.precedent !== undefined) {
          if(groups.precedent_is_att) {
            continue;
          }
          if(groups.max_d !== undefined && groups.max_d < 1) {
            groups.min_d = -groups.max_d;
            if(groups.min_d === undefined) {
              groups.max_d = undefined;
            } else {
              groups.max_d = -groups.min_d;
            }
            replacedWithWords = replacedWithWords.replace(direct_terms[j] + " ", " true ");
            data[groups.precedent].posteriors.push({precedent: i, min_d: groups.min_d, max_d: groups.max_d});
          } else if(groups.min_d !== undefined && groups.min_d > -1) {
            data[i].posteriors.push({precedent: groups.precedent, min_d: groups.min_d, max_d: groups.max_d});
          }
        } else if(groups.position !== undefined) {
          if(data[i].position !== undefined && data[i].position !== groups.position) {
            inconsist(i, j, `position "${groups.position}" at ${getAddress(i, j)} is not unique.`, []);
            return -1;
          }
          data[i].position = groups.position;
        }
      }
    }
    data[i].reverseReplacements = reverseReplacements;
    data[i].replacedWithWords = replacedWithWords;
  }

  let incoherence = checkGraph(data, allMediaRows.length);
  if(incoherence !== null) {
    inconsist(incoherence.row, 0, incoherence.result, []);
    return -1;
  }

  let newInds = [];
  let a = 0;
  for (let i = 1; i < nbLineBef; i++) {
    newInds.push(a);
    if(!data[i].isAtt) {
      a++;
    }
  }
  var lowests = [];
  for(const i of allMediaRows) {
    data[i].nbPost = 0;
    for (let j = 0; j < data[i].posteriors.length; j++) {
      if(data[i].posteriors[j].min_d > 0) {
        data[data[i].posteriors[j].precedent].ulteriors.push(newInds[i]);
        data[i].nbPost++;
      }
    }
    data[i].highest = allMediaRows.length - data[i].minDist;
    lowests.push(data[i].minDist);
    data[i].posteriors = undefined;
    data[i].minDist = undefined;
    data[i].maxDist = undefined;

    
    // simplify the formulas
    let reverseReplacements = data[i].reverseReplacements;
    let replacedWithWords = data[i].replacedWithWords;
    let simplified = [];
    data[i].conditions = [];
    if(replacedWithWords !== '') {
      try {
        // Await the result from the Python script
        simplified = await runPythonScript("simplify_expression", replacedWithWords);
      } catch (err) {
        inconsist(i, 0, `incorrect condition : ${err.message}.`, []);
        return -1;
      }

      let formulaList = transformLogicalFormula(getTrimmedResult(simplified));

      // change the 
      data[i].conditions = transformTerms(formulaList, exampleTransform, reverseReplacements, newInds);
    }
  }
  
  checkWithC(sheetCodeName);
  response.push({ "listBoxList": [] });
  resolved[sheetCodeName] = true;
}

async function checkWithC(params, sheetCodeName) {
  let values = params.values;
  let i = params.i;
  let j = params.j;
  let k = params.k;
  let val = params.val;
  let columnTitle = params.columnTitle;
  let message = params.message;
  let suggs = params.suggs;
  let msgTypeColors = params.msgTypeColors;
  let values0Glob = params.values0Glob;
  let valuesGlob = params.valuesGlob;
  let prevLine = params.prevLine;
  let lenAgg = params.lenAgg;
  let attributes = params.attributes;
  let sorting = params.sorting;
  let sheetCodeName = params.sheetCodeName;
  let editRow = params.editRow;
  let editCol = params.editCol;
  let resolved = params.resolved;
  let oldVersionsMaxNb = params.oldVersionsMaxNb;
  let columnTypes = params.columnTypes;
  let data = params.data;
  let allMediaRows = params.allMediaRows;
  let attrOfAttr = params.attrOfAttr;
  let nameSt = params.nameSt;
  let nameInd = params.nameInd;
  let mediaSt = params.mediaSt;
  let mediaInd = params.mediaInd;
  let headerColors = params.headerColors;
  let attNames = params.attNames;
  let actionsHistory = params.actionsHistory;
  let indActHist = params.indActHist;
  let handleChanges = params.handleChanges;
  let rowIdMap = params.rowIdMap;
  let response = params.response;
  
  let colNumb = params.colNumb;
  if(0) {
    resolved[sheetCodeName] = false;
  }
}

async function check(params, sheetCodeName) {
  // let values = params.values;
  let i = params.i;
  let j = params.j;
  let k = params.k;
  let val = params.val;
  let columnTitle = params.columnTitle;
  let message = params.message;
  let suggs = params.suggs;
  let msgTypeColors = params.msgTypeColors;
  let values0Glob = params.values0Glob;
  let valuesGlob = params.valuesGlob;
  let prevLine = params.prevLine;
  let lenAgg = params.lenAgg;
  let attributes = params.attributes;
  let sorting = params.sorting;
  let sheetCodeName = params.sheetCodeName;
  let editRow = params.editRow;
  let editCol = params.editCol;
  let resolved = params.resolved;
  let oldVersionsMaxNb = params.oldVersionsMaxNb;
  let columnTypes = params.columnTypes;
  let data = params.data;
  let allMediaRows = params.allMediaRows;
  let attrOfAttr = params.attrOfAttr;
  let nameSt = params.nameSt;
  let nameInd = params.nameInd;
  let mediaSt = params.mediaSt;
  let mediaInd = params.mediaInd;
  let headerColors = params.headerColors;
  let attNames = params.attNames;
  let actionsHistory = params.actionsHistory;
  let indActHist = params.indActHist;
  let handleChanges = params.handleChanges;
  let rowIdMap = params.rowIdMap;
  let response = params.response;
  
  let colNumb = params.colNumb;
  resolved[sheetCodeName] = false;
  
  data = [-1];
  let values = valuesGlob[sheetCodeName];
  if(initializeData(data, nbLineBef) == -1) {
    return -1;
  }
  stop = false;
  attributes = [];

  if(getAttributes(values, sheetCodeName) == -1) {
    return -1;
  }

  if(getNames(values) == -1) {
    return -1;
  }

  if(setAllAttr() == -1) {
    return -1;
  }

  await getConditions(values, sheetCodeName);
}

function searching(params, values, r, d, f) {
  let values = params.values;
  let i = params.i;
  let j = params.j;
  let k = params.k;
  let val = params.val;
  let columnTitle = params.columnTitle;
  let message = params.message;
  let suggs = params.suggs;
  let msgTypeColors = params.msgTypeColors;
  let values0Glob = params.values0Glob;
  let valuesGlob = params.valuesGlob;
  let prevLine = params.prevLine;
  let lenAgg = params.lenAgg;
  let attributes = params.attributes;
  let sorting = params.sorting;
  let sheetCodeName = params.sheetCodeName;
  let editRow = params.editRow;
  let editCol = params.editCol;
  let resolved = params.resolved;
  let oldVersionsMaxNb = params.oldVersionsMaxNb;
  let columnTypes = params.columnTypes;
  let data = params.data;
  let allMediaRows = params.allMediaRows;
  let attrOfAttr = params.attrOfAttr;
  let nameSt = params.nameSt;
  let nameInd = params.nameInd;
  let mediaSt = params.mediaSt;
  let mediaInd = params.mediaInd;
  let headerColors = params.headerColors;
  let attNames = params.attNames;
  let actionsHistory = params.actionsHistory;
  let indActHist = params.indActHist;
  let handleChanges = params.handleChanges;
  let rowIdMap = params.rowIdMap;
  let response = params.response;
  
  let colNumb = params.colNumb;

  elem2 = values[r][d][f];
  val2 = elem2.toLowerCase();
  let stop2 = true;
  if (val2 != val) {
    stop2 = false;
    for (let q = 1; q < 3; q++) {
      if (values[r][d][f].startsWith(perNot[q])) {
        console.log(`startsWith : ${values[r][d][f]} ${perNot[q]}`)
        elem2 = values[r][d][f].slice(perNot[q].length).trim();
        val2 = elem2.toLowerCase();
        if (val2 == val) {
          stat2 = q;
          return true;
        }
      }
    }
  }
  stat2 = 0;
  console.log("return stop2")
  return stop2;
}

function inconsist_removing_elem(values, i, j, k, message) {
  inconsist(i, j, message, [values[i][j].filter((_, m) => m !== k)]);
}

function inconsist_replacing_elem(values, i, j, k, newElement, message) {
  inconsist(i, j, message, [values[i][j].map((item, m) => (m === k ? newElement : item))]);
}

function midToWhite(msgTypeColors, theMsgType = msgType.ERROR) {
  let r = msgTypeColors[theMsgType].r % 256;
  let g = Math.floor(msgTypeColors[theMsgType].g / 256) % 256;
  let b = Math.floor(msgTypeColors[theMsgType].b / 65536) % 256;
  let newR = Math.floor((r + 255) / 2);
  let newG = Math.floor((g + 255) / 2);
  let newB = Math.floor((b + 255) / 2);
  return {r:newR, g:newG, b:newB};
}

function displayReference(params, row) {
  // values, row
  let values = params.values;
  let i = params.i;
  let j = params.j;
  let k = params.k;
  let val = params.val;
  let columnTitle = params.columnTitle;
  let message = params.message;
  let suggs = params.suggs;
  let msgTypeColors = params.msgTypeColors;
  let values0Glob = params.values0Glob;
  let valuesGlob = params.valuesGlob;
  let prevLine = params.prevLine;
  let lenAgg = params.lenAgg;
  let attributes = params.attributes;
  let sorting = params.sorting;
  let sheetCodeName = params.sheetCodeName;
  let editRow = params.editRow;
  let editCol = params.editCol;
  let resolved = params.resolved;
  let oldVersionsMaxNb = params.oldVersionsMaxNb;
  let columnTypes = params.columnTypes;
  let data = params.data;
  let allMediaRows = params.allMediaRows;
  let attrOfAttr = params.attrOfAttr;
  let nameSt = params.nameSt;
  let nameInd = params.nameInd;
  let mediaSt = params.mediaSt;
  let mediaInd = params.mediaInd;
  let headerColors = params.headerColors;
  let attNames = params.attNames;
  let actionsHistory = params.actionsHistory;
  let indActHist = params.indActHist;
  let handleChanges = params.handleChanges;
  let rowIdMap = params.rowIdMap;
  let response = params.response;
  let colNumb = params.colNumb;

  let params = {};
  params.values = values;
  params.i = i;
  params.j = j;
  params.k = k;
  params.val = val;
  params.columnTitle = columnTitle;
  params.message = message;
  params.suggs = suggs;
  params.msgTypeColors = msgTypeColors;
  params.values0Glob = values0Glob;
  params.valuesGlob = valuesGlob;
  params.prevLine = prevLine;
  params.lenAgg = lenAgg;
  params.attributes = attributes;
  params.sorting = sorting;
  params.sheetCodeName = sheetCodeName;
  params.editRow = editRow;
  params.editCol = editCol;
  params.resolved = resolved;
  params.oldVersionsMaxNb = oldVersionsMaxNb;
  params.columnTypes = columnTypes;
  params.data = data;
  params.allMediaRows = allMediaRows;
  params.attrOfAttr = attrOfAttr;
  params.nameSt = nameSt;
  params.nameInd = nameInd;
  params.mediaSt = mediaSt;
  params.mediaInd = mediaInd;
  params.headerColors = headerColors;
  params.attNames = attNames;
  params.actionsHistory = actionsHistory;
  params.indActHist = indActHist;
  params.handleChanges = handleChanges;
  params.rowIdMap = rowIdMap;
  params.response = response;
  params.colNumb = colNumb;
  
  addListBoxList(params);
  response[0]["listBoxList"][msgType.RELATIVES].push({color: msgTypeColors[msgType.RELATIVES], msg: "; ".join(values[r][0]), actions: [{action: Actions.Select, address: [row, 0]}]});
}

function inconsist(params, theAction = Actions.NewVal) {
  let values = params.values;
  let i = params.i;
  let j = params.j;
  let k = params.k;
  let val = params.val;
  let columnTitle = params.columnTitle;
  let message = params.message;
  let suggs = params.suggs;
  let msgTypeColors = params.msgTypeColors;
  let values0Glob = params.values0Glob;
  let valuesGlob = params.valuesGlob;
  let prevLine = params.prevLine;
  let lenAgg = params.lenAgg;
  let attributes = params.attributes;
  let sorting = params.sorting;
  let sheetCodeName = params.sheetCodeName;
  let editRow = params.editRow;
  let editCol = params.editCol;
  let resolved = params.resolved;
  let oldVersionsMaxNb = params.oldVersionsMaxNb;
  let columnTypes = params.columnTypes;
  let data = params.data;
  let allMediaRows = params.allMediaRows;
  let attrOfAttr = params.attrOfAttr;
  let nameSt = params.nameSt;
  let nameInd = params.nameInd;
  let mediaSt = params.mediaSt;
  let mediaInd = params.mediaInd;
  let headerColors = params.headerColors;
  let attNames = params.attNames;
  let actionsHistory = params.actionsHistory;
  let indActHist = params.indActHist;
  let handleChanges = params.handleChanges;
  let rowIdMap = params.rowIdMap;
  let response = params.response;
  let colNumb = params.colNumb;

  let params = {};
  params.values = values;
  params.i = i;
  params.j = j;
  params.k = k;
  params.val = val;
  params.columnTitle = columnTitle;
  params.message = message;
  params.suggs = suggs;
  params.msgTypeColors = msgTypeColors;
  params.values0Glob = values0Glob;
  params.valuesGlob = valuesGlob;
  params.prevLine = prevLine;
  params.lenAgg = lenAgg;
  params.attributes = attributes;
  params.sorting = sorting;
  params.sheetCodeName = sheetCodeName;
  params.editRow = editRow;
  params.editCol = editCol;
  params.resolved = resolved;
  params.oldVersionsMaxNb = oldVersionsMaxNb;
  params.columnTypes = columnTypes;
  params.data = data;
  params.allMediaRows = allMediaRows;
  params.attrOfAttr = attrOfAttr;
  params.nameSt = nameSt;
  params.nameInd = nameInd;
  params.mediaSt = mediaSt;
  params.mediaInd = mediaInd;
  params.headerColors = headerColors;
  params.attNames = attNames;
  params.actionsHistory = actionsHistory;
  params.indActHist = indActHist;
  params.handleChanges = handleChanges;
  params.rowIdMap = rowIdMap;
  params.response = response;
  params.colNumb = colNumb;
  
  addListBoxList(params);
  response[0]["listBoxList"][msgType.ERROR].push({color: msgTypeColors[msgType.ERROR], msg: message, actions: [{action: Actions.Select, address: [i, j]}]});
  suggs.forEach(sugg => {
    let theNewVal = sugg;
    if (theAction == Actions.NewVal) {
      theNewVal = sugg.join("; ");
    }
    response[0]["listBoxList"][msgType.ERROR].push({color: midToWhite(msgTypeColors), msg: theNewVal, actions: [{action: Actions.Select, address: 0}, {action: theAction, address: 0, newVal: theNewVal}]});
  });
}

function findNewName(params, name, checked = true) {
  let values = params.values;
  let i = params.i;
  let j = params.j;
  let k = params.k;
  let val = params.val;
  let columnTitle = params.columnTitle;
  let message = params.message;
  let suggs = params.suggs;
  let msgTypeColors = params.msgTypeColors;
  let values0Glob = params.values0Glob;
  let valuesGlob = params.valuesGlob;
  let prevLine = params.prevLine;
  let lenAgg = params.lenAgg;
  let attributes = params.attributes;
  let sorting = params.sorting;
  let sheetCodeName = params.sheetCodeName;
  let editRow = params.editRow;
  let editCol = params.editCol;
  let resolved = params.resolved;
  let oldVersionsMaxNb = params.oldVersionsMaxNb;
  let columnTypes = params.columnTypes;
  let data = params.data;
  let allMediaRows = params.allMediaRows;
  let attrOfAttr = params.attrOfAttr;
  let nameSt = params.nameSt;
  let nameInd = params.nameInd;
  let mediaSt = params.mediaSt;
  let mediaInd = params.mediaInd;
  let headerColors = params.headerColors;
  let attNames = params.attNames;
  let actionsHistory = params.actionsHistory;
  let indActHist = params.indActHist;
  let handleChanges = params.handleChanges;
  let rowIdMap = params.rowIdMap;
  let response = params.response;
  
  let colNumb = params.colNumb;
  let alreadyExist = true;
  while (alreadyExist) {
    if(checked) {
      let theMatch = name.match(/(.*? -)(\d)$/);
      if (theMatch) {
        name = theMatch[1] + (+theMatch[2] + 1);
      } else {
        name += " -2";
      }
    }
    let name2 = name;
    name = name.trim();
    checked = true;
    alreadyExist = false;
    for (let r = 1; r < nbLineBef; r++) {
      for(let s = 0; s < colNumb; s++) {
        if(columnTypes[s] != allColumnTypes.ATTRIBUTES) {
          continue;
        }
        for (let f = 0; f < values[r][s].length; f++) {
          if (name === values[r][s][f]) {
            alreadyExist = true;
            break;
          }
        }
        if(alreadyExist) {
          break;
        }
      }
      if(alreadyExist) {
        break;
      }
    }
    name = name2;
  }
  return name;
}

function correct(params) {
  let values = params.values;
  let i = params.i;
  let j = params.j;
  let k = params.k;
  let val = params.val;
  let columnTitle = params.columnTitle;
  let message = params.message;
  let suggs = params.suggs;
  let msgTypeColors = params.msgTypeColors;
  let values0Glob = params.values0Glob;
  let valuesGlob = params.valuesGlob;
  let prevLine = params.prevLine;
  let lenAgg = params.lenAgg;
  let attributes = params.attributes;
  let sorting = params.sorting;
  let sheetCodeName = params.sheetCodeName;
  let editRow = params.editRow;
  let editCol = params.editCol;
  let resolved = params.resolved;
  let oldVersionsMaxNb = params.oldVersionsMaxNb;
  let columnTypes = params.columnTypes;
  let data = params.data;
  let allMediaRows = params.allMediaRows;
  let attrOfAttr = params.attrOfAttr;
  let nameSt = params.nameSt;
  let nameInd = params.nameInd;
  let mediaSt = params.mediaSt;
  let mediaInd = params.mediaInd;
  let headerColors = params.headerColors;
  let attNames = params.attNames;
  let actionsHistory = params.actionsHistory;
  let indActHist = params.indActHist;
  let handleChanges = params.handleChanges;
  let rowIdMap = params.rowIdMap;
  let response = params.response;
  
  let colNumb = params.colNumb;
  for (let i = 1; i < nbLineBef; i++) {
    for (let j = 0; j < colNumb; j++) {
      let value = values[i][j].join("; ");
      if (values0Glob[sheetCodeName][i][j] != value && values0Glob[sheetCodeName][i][j]!=null && value != "") {
        values0Glob[sheetCodeName][i][j] = value;
        response.push({ "chgValue": [i + 1, j + 1, value] });
      }
    }
  }
}

function dataGeneratorSub(params) {
  let values = params.values;
  let i = params.i;
  let j = params.j;
  let k = params.k;
  let val = params.val;
  let columnTitle = params.columnTitle;
  let message = params.message;
  let suggs = params.suggs;
  let msgTypeColors = params.msgTypeColors;
  let values0Glob = params.values0Glob;
  let valuesGlob = params.valuesGlob;
  let prevLine = params.prevLine;
  let lenAgg = params.lenAgg;
  let attributes = params.attributes;
  let sorting = params.sorting;
  let sheetCodeName = params.sheetCodeName;
  let editRow = params.editRow;
  let editCol = params.editCol;
  let resolved = params.resolved;
  let oldVersionsMaxNb = params.oldVersionsMaxNb;
  let columnTypes = params.columnTypes;
  let data = params.data;
  let allMediaRows = params.allMediaRows;
  let attrOfAttr = params.attrOfAttr;
  let nameSt = params.nameSt;
  let nameInd = params.nameInd;
  let mediaSt = params.mediaSt;
  let mediaInd = params.mediaInd;
  let headerColors = params.headerColors;
  let attNames = params.attNames;
  let actionsHistory = params.actionsHistory;
  let indActHist = params.indActHist;
  let handleChanges = params.handleChanges;
  let rowIdMap = params.rowIdMap;
  let response = params.response;
  
  let colNumb = params.colNumb;
  if (!resolved[sheetCodeName]) {
    return -1;
  }
  correct(values, sheetCodeName);
  
  // Example data to pass to the C program
  let dataJson = [
    {
      ulteriors: [],
      conditions: [],
      nbPost: 0,
      highest: -1
    },
    {
      ulteriors: [0],
      conditions: [],
      nbPost: 0,
      highest: -1
    },
    {
      ulteriors: [0],
      conditions: [],
      nbPost: 0,
      highest: -1
    },
    {
      ulteriors: [1, 2],
      conditions: [],
      nbPost: 0,
      highest: -1
    }
  ];
  dataJson = [];
  for(let i = 0; i < lenAgg; i++) {
    dataJson.push({
      ulteriors: data[i].ulteriors,
      conditions: data[i].conditions,
      nbPost: data[i].nbPost,
      highest: data[i].highest
    });
  }


  
  const programPath = "C:/Users/abarb/Documents/health/news_underground/mediaSorter/programs/c_prog/mediaSorter/x64/Release/mediaSorter.exe";

  let dataFolder = `C:/Users/abarb/Documents/health/news_underground/mediaSorter/programs/data/${sheetCodeName}`;
  
  // check if the folder exists, if not create it
  if (!fs.existsSync(dataFolder)) {
    fs.mkdirSync(dataFolder, { recursive: true });
  }

  
  
  // Convert the object to JSON and write it to a file
  fs.writeFile(dataFolder + `/data.json`, JSON.stringify(data, null, 2), (err) => {
    if (err) throw err;
    console.log('Data has been written to data.json');
  });

  exec(`${programPath} ${sheetCodeName}`, (error, stdout, stderr) => {
    if (error) {
      console.error(`Error executing program: ${error}`);
      return;
    }
    if (stderr) {
      console.error(`Error output: ${stderr}`);
      return;
    }
    // Output from the C program will be in stdout
    console.log(`Program output: ${stdout}`);
  });
}

async function renameSymbol(params, oldValue, newValue) {
  let values = params.values;
  let i = params.i;
  let j = params.j;
  let k = params.k;
  let val = params.val;
  let columnTitle = params.columnTitle;
  let message = params.message;
  let suggs = params.suggs;
  let msgTypeColors = params.msgTypeColors;
  let values0Glob = params.values0Glob;
  let valuesGlob = params.valuesGlob;
  let prevLine = params.prevLine;
  let lenAgg = params.lenAgg;
  let attributes = params.attributes;
  let sorting = params.sorting;
  let sheetCodeName = params.sheetCodeName;
  let editRow = params.editRow;
  let editCol = params.editCol;
  let resolved = params.resolved;
  let oldVersionsMaxNb = params.oldVersionsMaxNb;
  let columnTypes = params.columnTypes;
  let data = params.data;
  let allMediaRows = params.allMediaRows;
  let attrOfAttr = params.attrOfAttr;
  let nameSt = params.nameSt;
  let nameInd = params.nameInd;
  let mediaSt = params.mediaSt;
  let mediaInd = params.mediaInd;
  let headerColors = params.headerColors;
  let attNames = params.attNames;
  let actionsHistory = params.actionsHistory;
  let indActHist = params.indActHist;
  let handleChanges = params.handleChanges;
  let rowIdMap = params.rowIdMap;
  let response = params.response;
  
  let colNumb = params.colNumb;
  if(!resolved[sheetCodeName]) {
    return -1;
  }
  let nbRen = 0;
  let oldNames = oldValue.split(";");
  let newNames = newValue.split(";");
  console.log(`oldName : ${oldNames}`);
  for (let t = 0; t < oldNames.length; t++) {
    val = oldNames[t].strip().toLowerCase();
    console.log(`val : ${val}`);
    if (newNames.length <= t) {
      elem = "";
    } else {
      elem = newNames[t];
    }
    console.log(elem);
    for (let i = 1; i < nbLineBef; i++) {
      for (let j = 0; j < colNumb; j++) {
        for (let k = 0; k < values[i][j].length; k++) {
          if (searching(values, i, j, k)) {
            console.log("elem");
            if (elem) {
              if (stat2 > -1) {
                values[i][j][k] = perNot[stat2] + elem;
              } else {
                values[i][j][k] = elem;
              }
            } else {
              values[i][j].splice(k, 1);
            }
            nbRen++;
          }
        }
      }
    }
  }
  correct(values, sheetCodeName);
  response.push({ "Renamings": [`renamed (${nbRen})`] });
  await check(sheetCodeName);
}

function handleoldNameInputClick(params) {
  let values = params.values;
  let i = params.i;
  let j = params.j;
  let k = params.k;
  let val = params.val;
  let columnTitle = params.columnTitle;
  let message = params.message;
  let suggs = params.suggs;
  let msgTypeColors = params.msgTypeColors;
  let values0Glob = params.values0Glob;
  let valuesGlob = params.valuesGlob;
  let prevLine = params.prevLine;
  let lenAgg = params.lenAgg;
  let attributes = params.attributes;
  let sorting = params.sorting;
  let sheetCodeName = params.sheetCodeName;
  let editRow = params.editRow;
  let editCol = params.editCol;
  let resolved = params.resolved;
  let oldVersionsMaxNb = params.oldVersionsMaxNb;
  let columnTypes = params.columnTypes;
  let data = params.data;
  let allMediaRows = params.allMediaRows;
  let attrOfAttr = params.attrOfAttr;
  let nameSt = params.nameSt;
  let nameInd = params.nameInd;
  let mediaSt = params.mediaSt;
  let mediaInd = params.mediaInd;
  let headerColors = params.headerColors;
  let attNames = params.attNames;
  let actionsHistory = params.actionsHistory;
  let indActHist = params.indActHist;
  let handleChanges = params.handleChanges;
  let rowIdMap = params.rowIdMap;
  let response = params.response;
  let colNumb = params.colNumb;

  let rowId = range.rowIndex; // Directly use rowIndex property
  let totalRows = range.rowCount; // Get total rows in the selection
  for (let i = 0; i < totalRows; i++) {
    // Add each row index to rowNumbers

    rowId = Math.min(rowId, range.rowIndex);
  }
  response.push({ "oldNameInput": [values[rowId][0].join("; ")] });
  response.push({ "newNameInput": [values[rowId][0].join("; ")] });
}

function getTrimmedResult(str) {
  return str.split(/(\s*&&\s*|\s*\|\|\s*|\s*!\s*|\s*\(\s*|\s*\)\s*)/).filter(Boolean);
}

function replaceWithSameWord(str) {
  const replacements = {}; // Object to store replacements
  const reverseReplacements = {}; // Object to store reverse mapping
  let wordCounter = 0; // Counter to generate new words

  const trimmedResult  = getTrimmedResult(str);
  const operators = new Set(["&&", "||", "!", "(", ")"]);
  
  // const nameRow = "[^\\[\\]&|!]+?";
  // let pattern = `^\s*(?:a\s*(?!_\s*[^\d\s:]*$)(?<min_d>:?\d+)?\s*_?\s*(?<max_d>-?\d+)?\s*(?:\s*"\s*(?<attrib_ref>${nameRow})\s*"\s*(?:\[\s*"\s*(?<attrib_ref_col>${nameRow})\s*"\s*\])?)?\s*"\s*(?<precedent>${nameRow})\s*"\s*(?:\[\s*"\s*(?<precedent_col>${nameRow})\s*"\s*\])?(?<isAny>\(any\))?)|(?<position>p\s*(?!_\s*[^\d\s:]*$)(?<pos_min_d>-?\d+)?\s*_?\s*(?<pos_max_d>-?\d+)?)\s*$`;
  // let regex = new RegExp(pattern);

  const nameRow = "[^\\[\\]&|!]+?";
  let pattern = `^\\s*(?:a\\s*(?!_\\s*[^\\d\\s:]*$)(?<min_d>:?\\d+)?\\s*_?\\s*(?<max_d>-?\\d+)?\\s*(?:\\s*"\\s*(?<attrib_ref>${nameRow})\\s*"\\s*(?:\\[\\s*"\\s*(?<attrib_ref_col>${nameRow})\\s*"\\s*\\])?)?\\s*"\\s*(?<precedent>${nameRow})\\s*"\\s*(?:\\[\\s*"\\s*(?<precedent_col>${nameRow})\\s*"\\s*\\])?(?<isAny>\\(any\\))?)|(?<position>p\\s*(?!_\\s*[^\\d\\s:]*$)(?<pos_min_d>-?\\d+)?\\s*_?\\s*(?<pos_max_d>-?\\d+)?)\\s*$`;
  let regex = new RegExp(pattern);


  let replacedWithWords = trimmedResult.map(item => {
    if (!operators.has(item)) {
      const match = str.match(regex);
      if(!match) {
        throw Error("incorrect condition.");
      }
      const groups = match.groups; // The last element in args array is the groups object
      groups.min_d = groups.pos_min_d;
      groups.max_d = groups.pos_max_d;
      delete groups.pos_min_d;
      delete groups.pos_max_d;

      if (!replacements[match[0]]) {
        const newWord = `word${wordCounter++}`;
        replacements[match[0]] = newWord;

        // Store the reverse mapping along with the capturing groups
        if(groups.max_d === undefined && groups.min_d === undefined) {
          if(groups.position !== undefined) {
            throw Error("incorrect condition : p must be followed by a number.");
          }
          groups.min_d = 1;
        } else if (groups.max_d !== undefined && groups.min_d !== undefined) {
          if (groups.max_d < groups.min_d) {
            throw Error("incorrect condition : the maximum distance is lesser than the minimum distance.");
          }
          if (groups.max_d === groups.min_d && groups.max_d === 0) {
            throw Error("incorrect condition : the maximum distance and the minimum distance are 0.");
          }
        }
        if (groups.position === undefined) {
          if (groups.min_d === 0) {
            groups.min_d = 1;
          } else if (groups.max_d === 0) {
            groups.max_d = -1;
          }
        }
        reverseReplacements[newWord] = [{}];
        reverseReplacements[newWord][0] = {
          ...groups
        };
      }
      return replacements[match[0]] + " ";
    } else {
      return item;
    }
  }).join('');

  return { replacedWithWords, reverseReplacements };
}

function runPythonScript(pythonFile, arg1) {
  return new Promise((resolve, reject) => {
    const options = {
      mode: 'json',
      pythonOptions: ['-u'], // Unbuffered stdout
      scriptPath: './', // Directory where the script is located
      args: [arg1], // Passing argument directly to the Python script
    };

    const pyshell = new PythonShell("C:\\Users\\abarb\\Documents\\health\\news_underground\\mediaSorter\\programs\\excel_prog\\mediaSorter\\" + pythonFile + ".py", options);

    // Receive output from the Python script
    pyshell.on('message', (message) => {
      resolve(message); // Resolve the promise with the message (result) from Python
    });

    pyshell.end((err) => {
      if (err) reject(err); // Reject the promise if there's an error
    });
  });
}

exports.executeSomething = async (req, res) => {
  try {
    let values0Glob = {};
    let valuesGlob = {};
    let prevLine = {};
    let lenAgg;
    let attributes;
    let sorting = {};
    let sheetCodeName;
    let editRow = {};
    let editCol = {};
    let resolved = {};
    let oldVersionsMaxNb = 20;
    let columnTypes;
    let data;
    let allMediaRows;
    let attrOfAttr = [];
    let nameSt = "names";
    let nameInd;
    let mediaSt = "media";
    let mediaInd = -1;
    let headerColors;
    let attNames;
    let actionsHistory = {};
    let indActHist = 0;
    let handleChanges = {};
    let rowIdMap = {};
    let response;
  
  
    const timestamp = req.body.timestamp;
    var msgTypeColors = [{ r: 255, g: 0, b: 0 }, { r: 255, g: 137, b: 0 }, { r: 0, g: 145, b: 255 }];
    
    // Write the timestamp to the received.txt file
    fs.appendFile("C:\\Users\\abarb\\Documents\\health\\news_underground\\mediaSorter\\programs\\excel_prog\\mediaSorter\\received.txt", timestamp + '\n', (err) => {
      if (err) {
          console.error('Error writing to file:', err);
          // Handle the error appropriately, e.g., by sending a response with an error message
      } else {
          console.log('Timestamp successfully written to file.');
          // You can also send a success response back to the client here
      }
    });
    
    response = [];
    let body = req.body;
    const funcName = body.functionName;
    let values = valuesGlob[sheetCodeName];
    switch (funcName) {
      case "handleChange":
        await handleChange(body.changes, sheetCodeName);
        break;
      case "selectionChange":
        editRow[sheetCodeName] = body.editRow;
        editCol[sheetCodeName] = body.editCol;
        
        let params = {};
        params.values = values;
        params.i = i;
        params.j = j;
        params.k = k;
        params.val = val;
        params.columnTitle = columnTitle;
        params.message = message;
        params.suggs = suggs;
        params.msgTypeColors = msgTypeColors;
        params.values0Glob = values0Glob;
        params.valuesGlob = valuesGlob;
        params.prevLine = prevLine;
        params.lenAgg = lenAgg;
        params.attributes = attributes;
        params.sorting = sorting;
        params.sheetCodeName = sheetCodeName;
        params.editRow = editRow;
        params.editCol = editCol;
        params.resolved = resolved;
        params.oldVersionsMaxNb = oldVersionsMaxNb;
        params.columnTypes = columnTypes;
        params.data = data;
        params.allMediaRows = allMediaRows;
        params.attrOfAttr = attrOfAttr;
        params.nameSt = nameSt;
        params.nameInd = nameInd;
        params.mediaSt = mediaSt;
        params.mediaInd = mediaInd;
        params.headerColors = headerColors;
        params.attNames = attNames;
        params.actionsHistory = actionsHistory;
        params.indActHist = indActHist;
        params.handleChanges = handleChanges;
        params.rowIdMap = rowIdMap;
        params.response = response;
        params.colNumb = colNumb;

        handleSelectLinks(params);
        break;
      case 'chgSheet':
        sheetCodeName = body.sheetCodeName;
        headerColors = body.headerColors;
        await handleChange([], sheetCodeName);
        break;
      case "dataGeneratorSub":
        dataGeneratorSub(values, sheetCodeName);
        break;
      case "stop sorting":
        if(sorting[sheetCodeName]) {
          // TODO: call the C program to stop sorting
        }
        break;
      case "better sorting":
        response.push({"sort": body});
        break;
      case "C stops sorting":
        response.push({ "stop sorting": [] });
        sorting[sheetCodeName] = false;
        break;
      case "renameSymbol":
        renameSymbol(values, sheetCodeName, body.oldValue, body.newValue);
        break;
      case "handleoldNameInputClick":
        handleoldNameInputClick(values);
        break;
  
      // TODO
      case "ctrlZ":
        // ctrlZ(sheetCodeName);
        break;
      case "ctrlY":
        // ctrlY(sheetCodeName);
        break;
  
      case "newSheet":
        sheetCodeName = body.sheetCodeName;
        values0Glob[sheetCodeName] = body.values;
        valuesGlob[sheetCodeName] = values0Glob[sheetCodeName].map(r => r.map(c => c===null?[]:c.toString().split(';').map(k => k.trim().toLowerCase()).filter(k => k)));
        sorting[body.sheetCodeName] = false;
        prevLine[body.sheetCodeName] = 0;
        break;
      case "show":
        if(!resolved[sheetCodeName]) {
          return -1;
        }
        // Modifie cette commande pour inclure l'AppUserModelID ou le chemin du fichier exécutable
        const command = `powershell -Command "Start-Process explorer.exe shell:AppsFolder\\a6714fbe-7044-42de-b8ab-099055a0b3b2_fc2wt02jznpqm!App" ${Math.max(0, body.orderedVideos)}`;
        
        exec(command, (error, stdout, stderr) => {
            if (error) {
                console.error(`Erreur lors de l'exécution de l'application: ${error.message}`);
                res.status(500).send(`Erreur: ${error.message}`);
                return;
            }
            if (stderr) {
                console.error(`Erreur standard: ${stderr}`);
                res.status(500).send(`Erreur standard: ${stderr}`);
                return;
            }
            console.log(`Résultat: ${stdout}`);
        });
        break;
      case "better sorting":
        response.push({"sort": body});
        break;
      case "C stops sorting":
        response.push({"sort": body});
        break;
    }
    res.json(response);
  } catch (error) {
    console.error('Error in executeSomething:', error);
    return res.status(500).json({ error: 'Internal server error' });
  }
};
