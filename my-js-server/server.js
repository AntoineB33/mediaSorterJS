const express = require('express');
const fs = require('fs');
const acorn = require('acorn');
const walk = require('acorn-walk');
const { exec } = require('child_process');
const app = express();
const port = 3000;
// Middleware to parse JSON bodies
app.use(express.json());

var msgTypeColors = [{ r: 255, g: 0, b: 0 }, { r: 255, g: 137, b: 0 }, { r: 0, g: 145, b: 255 }];
var updateRegularity = 1;

var values0 = {};
var values = [];
var oldValues = [];
var indOldValues = 0;
var stop = false;
var prevLine = {};
var nbLineBef;
var fileId = "1EoCDqXsL0tqAW6M7qVQT_N7L_NsBirMJ";
var lenAgg;
var attributes;
var dataAgg;
var progChg = 0;
var sorting = {};
var colors = [];
var fonts = [];
var place;
var steps;
var visited;
var notInitialized = true;
var updating = false;
var handleRunning = false;
var checkRunning = false;
var sheetCodeName;
var selectedCells = {};
var resolved = {};
var oldVersionsMaxNb = 20;
var columnTypes;
var child;
var data;
var headerNames;
var nameInd;
var mediaInd;
var headerColors;
var attNames;

// [node.js]
var response;


const allColumnTypes = {
  CONDITIONS: 0,
  ATTRIBUTES: 1
};

const allColumnNames = {
  NAMES: "names",
  MEDIA: "media"
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


class Node {
    constructor(value) {
        this.value = value;
        this.left = null;
        this.right = null;
    }
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

/**
 *  Update the links to show at the menu each time a cell is selected
 * 
 */
function handleSelectLinks() {
  // only if no errors in the sheet
  if(!resolved[sheetCodeName]) {
    return -1;
  }
  const match = selectedCells[sheetCodeName].match(/\$?([A-Z]+)\$?(\d+)/);
  const columnLetters = match[1];
  const row = +match[2] - 1;
  
  // Convert column letters to a number
  let column = 0;
  for (let i = 0; i < columnLetters.length; i++) {
    column = column * 26 + (columnLetters.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  column -= 1
  
  if(row>=nbLineBef || column >= colNumb) {
    return -1;
  }
  
  if(column==0) {
    if(perRef[row] > -1) {
      for(let g = 0; g < 3; g++) {
        if(periods[perRef[row]][g] < nbLineBef && periods[perRef[row]][g]!=row) {
          displayReference(periods[perRef[row]][g]);
        }
      }
    }
  } else {
    for(let k = 0; k < values[row][column].length; k++) {
      if(column < 4) {
        for(let r = 1; r < values.length; r++) {
          for(let f = 0; f < values[r][0].length; f++) {
            if(row!=r && searching(r, 0, f)) {
              if(perRef[r] > -1) {
                for(let g = 0; g < 3; g++) {
                  if(periods[perRef[r]][g]!=-1) {
                    displayReference(r);
                  }
                }
              } else {
                displayReference(r);
              }
            }
          }
        }
      } else if (column < colNumb - 1) {
        nbMediaInCat = attributes[data[row][2]]
        for(let r = 1; r < values.length; r++) {
          for(let d = 4; d < colNumb - 1; d++) {
            for(let f = 0; f < values[r][d].length; f++) {
              if(row!=r && searching(r, d, f)) {
                displayReference(r);
              }
            }
          }
        }
      }
    }
  }
}

function handleChange(updates) {
  var splitValue = [];
  updates.forEach(function(update) {
    var row = update[0] - 1;  // Adjusting for zero-based index
    var column = update[1] - 1;  // Adjusting for zero-based index
    var value = update[2];
    values0[sheetCodeName][row][column] = value;
    if(value) {
      while(headerColors.length <= j) {
        headerColors.push(null);
        columnTypes.push(null);
      }
      splitValue = (row == 0 || columnTypes[j] == allColumnTypes.ATTRIBUTES?
            value.split(";"):
            [value])
          .map(v => v.trim().toLowerCase())
          .filter(e => e !== "");
      if(row == 0) {
        for(let k = 0; k < splitValue.length; k++) {
          if(values[0][j][k].match(/^[A-Z]+$/)) {
            inconsist(0, j, `Column name at ${getAddress(0, column)} is only upper case letters.`, [findNewName(values[0][column][k])]);
            return;
          }
        }
      }
    }
    // Expand the matrix rows if needed
    while (values.length <= row) {
      values.push([]);
    }
    // Expand the matrix columns if needed
    for (let i = 0; i < values.length; i++) {
      while (values[i].length <= column) {
        values[i].push([]);
      }
    }
    values[row][column] = splitValue;
  });

  nbLineBef = values.length;
  
  const nbLineBefBef = nbLineBef;
  for (let i = nbLineBef - 1; i > 0; i--) {
    if (values[i].some(v => v.length != 0)) {
      break;
    }
    nbLineBef--;
  }
  

  // Label the columns "names" and "media"
  headerNames = {};
  let namesColor;
  let alreadyLabelled;
  let missingNamesNb = Object.values(allColumnNames).length;
  resolved[sheetCodeName] = false;
  for(let j = 0; j < values[0].length; j++) {
    alreadyLabelled = "";
    for(let k = 0; k < values[0][j].length; k++) {
      let name = values[0][j][k];
      if(Object.values(allColumnNames).find(item => item === name)) {
        if(alreadyLabelled) {
          inconsist_removing_elem(0, j, k, `Column ${getColumnTag(j)} labelled as "${alreadyLabelled}" and ${name}.`);
          return;
        }
        if(headerNames[name] !== undefined) {
          inconsist_removing_elem(0, j, k, `Column ${getColumnTag(j)} labelled as "${name}" already at column ${getColumnTag(headerNames[name])}.`);
          return;
        }
        if(namesColor !== undefined && namesColor !== headerColors[j]) {
          inconsist(0, j, `Column ${getColumnTag(j)} must have the same background color as the attributes.`, [namesColor], true);
          return;
        }
        if(name == allColumnNames.NAMES) {
          nameInd = j;
        } else if(name == allColumnNames.MEDIA) {
          mediaInd = j;
        }
        missingNamesNb--;
        namesColor = headerColors[j];
        alreadyLabelled = name;
        headerNames[name] = j;
      }
    }
  }
  if(missingNamesNb) {
    let suggs = Object.values(allColumnNames).filter((name) => headerNames[name] === undefined);
    let lastCol = values[0].length;
    let actions = [{action: Actions.Select, address: [0, lastCol]}];
    suggs.forEach(sugg => {
      actions.push({action: Actions.NewVal, address: [0, lastCol], newVal: sugg});
      lastCol+=1;
    });
    response[0]["listBoxList"][msgType.ERROR].push({color: msgTypeColors[msgType.ERROR], msg: `Missing columns: ${suggs.join(", ")}.`, actions: actions});
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

  values.splice(nbLineBef, nbLineBefBef - nbLineBef);
  check();
  oldValues.splice(indOldValues, oldValues.length - 1 - indOldValues);
  oldValues.push(values);
  if(oldVersionsMaxNb < oldValues.length) {
    oldValues.shift();
  }
}

function getAddress(i, j) {
  return getColumnTag(j) + ":" + (i + 1);
}

function checkQuotes(i, val) {
  let countQuot = (val.match(new RegExp(`"`, 'g')) || []).length;
  if(countQuot != 2 || columnTitle[0] != '"' || columnTitle[columnTitle.length - 1] != '"') {
    inconsist_removing_elem(i, nameInd, 0, `The condition at ${getAddress(i, nameInd)} includes square brackets, thus should be in the format ""attribute"["column title"]".`)
    return -1;
  }
}

function checkBrack(i, k, val, withQuotes) {
  let countOpenBr = (val.match(new RegExp(`\\[`, 'g')) || []).length;
  let countCloseBr = (val.match(new RegExp(`\\]`, 'g')) || []).length;
  let refInd = -1;
  if(countOpenBr || countCloseBr) {
    let ind = val.indexOf('[');
    if(val[-1] != ']' || countOpenBr != 1 || countCloseBr != 1 || ind == 0) {
      val = val.replace(new RegExp(`[\\[\\]]`, 'g'), replacementChar);
      inconsist_removing_elem(i, nameInd, k, `The name at ${getAddress(i, nameInd)} includes square brackets, thus should be in the format "attribute[column title]".`)
      return -1;
    }
    let att_col_NAME = values[i][nameInd][k].split('[');
    let columnTitle = att_col_NAME[1].slice(0, -1).trim();
    if(withQuotes) {
      if(checkQuotes(i, columnTitle) == -1) {
        return -1;
      }
      columnTitle = columnTitle.slice(1, -1);
    }
    let columnInd = -1;
    for(let j = 0; j < colNumb; j++) {
      if(columnTypes[j] == allColumnTypes.ATTRIBUTES && values[0][j].includes(columnTitle)) {
        columnInd = j;
        break;
      }
    }
    if(columnInd == -1) {
      inconsist_removing_elem(i, nameInd, k, `No column named "${columnTitle}" at ${getAddress(i, nameInd)} has attribute.`)
      return -1;
    }
    att_col_NAME[0] = att_col_NAME[0].trim();
    if(withQuotes) {
      if(checkQuotes(i, att_col_NAME[0]) == -1) {
        return -1;
      }
      att_col_NAME[0] = att_col_NAME[0].slice(1, -1);
    }
    for(let m = 0; m < attNames[columnInd].length; m++) {
      if(attNames[columnInd][m] == att_col_NAME[0]) {
        refInd = m;
        break;
      }
    }
    if(refInd == -1) {
      inconsist_replacing_elem(i, nameInd, k, `The column named "${columnTitle}" at ${getAddress(i, nameInd)} doesn't have the attribute "${att_col_NAME[0]}".`)
      return -1;
    }
    let quotes = withQuotes?'"':'';
    return [refInd, quotes + att_col_NAME[0] + quotes + (att_col_NAME[0].includes(" ")?" ":"") + "[" + quotes + columnTitle + quotes + "]"];
  } else {
    for(let key in attNames) {
      for(let m = 0; m < attNames[key].length; m++) {
        if(attNames[key][m] == att_col_NAME[0]) {
          if(refInd != -1) {
            inconsist_replacing_elem(i, nameInd, k, findNewName(val, false), `The attribute at ${getAddress(i, nameInd)} exists in multiple columns.`);
            return -1;
          }
          refInd = m;
          break;
        }
      }
    }
    return [refInd, val];
  }
}

function check() {
  resolved[sheetCodeName] = false;
  let found;
  
  data = [undefined];
  var r;

  for (let i = 1; i < nbLineBef; i++) {
    data.push({
      isNotFollowing: [true],
      precedents: [],
      posteriors: [],
      follower: [],
      follower_to_attrib: [],
      follower_with_delay: [],
      followed_with_delay: [],
      attributes:[],
      other_conditions:[],
    });
    let cellList = values[i][nameInd];
    // check if a non-empty row has no name
    if(!cellList.length) {
      for (let j = 0; j < values[i].length; j++) {
        if (values[i][j].length !== 0) {
          inconsist(i, nameInd, `missing name at ${getAddress(i, nameInd)}`, [[findNewName("")]]);
          return;
        }
      }
      data[i] = undefined;
      continue;
    }
  }
  stop = false;
  attributes = [];
  let acc = 0;
  let accAttr = {};
  for (let j = 0; j < colNumb; j++) {
    accAttr[j] = acc;
    for (let i = 1; i < nbLineBef; i++) {
      for (let k = 1; k < values[i][j].length; i++) {
        let val = values[i][j][k];
        if (columnTypes[j] == allColumnTypes.CONDITIONS) {
          let tree = parseExpression(val);
          let necessarySubformulas = findNecessarySubformulas(tree);
          for(let r = 0; r < necessarySubformulas.length; r++) {
            let countQuotes = (val.match(new RegExp(`"`, 'g')) || []).length;
            if(necessarySubformulas[r][0] == 'a') {
              if(necessarySubformulas[r][necessarySubformulas[r].length - 1] != '"' || countQuotes != 2) {
                inconsist_removing_elem(i, j, k, `The condition at ${getAddress(i, j)} should be in the format a"name_of_the_row" .`);
                return;
              }
              let oper = necessarySubformulas[r].split('"')[0].slice(1).split("[");
              let cat = -1;
              let att = oper[1];
              if(oper.length == 2 && )
              if(oper.length == 1) {
                oper = oper[0];
              } else if(oper.length == 2) {
                att = att.slice(0, -1);
                for(let key in attNames) {
                  for(let m = 0; m < attNames[key].length; m++) {
                    if(attNames[key][m] == att) {
                      if(cat != -1) {
                        inconsist_removing_elem(i, nameInd, k, `The attribute "${att}" at ${getAddress(i, nameInd)} exists in multiple columns.`);
                        return;
                      }
                      cat = accAttr[key] + m;
                      break;
                    }
                  }
                }
                if(cat == -1) {
                  inconsist_removing_elem(i, nameInd, k, `The attribute "${att}" at ${getAddress(i, nameInd)} doesn't exist.`);
                  return;
                }
              } else if(oper.length == 3) {
                let colName = oper[2].slice(0, -2);
                let columnInd = -1;
                for(let v = 0; v < values[0].length; v++) {
                  for(let w = 0; w < values[0][v].length; w++) {
                    if(values[0][v][w].includes(colName)) {
                      columnInd = v;
                      break;
                    }
                  }
                }
                if(columnInd == -1) {
                  inconsist_removing_elem(i, nameInd, k, `The column "${colName}" at ${getAddress(i, nameInd)} doesn't exist.`);
                  return;
                }
                for(let m = 0; m < attNames[columnInd].length; m++) {
                  if(attNames[columnInd][m] == att) {
                    cat = accAttr[columnInd] + m;
                    break;
                  }
                }
                if(cat==-1) {
                  inconsist_removing_elem(i, nameInd, k, `The attribute "${att}" at ${getAddress(i, nameInd)} doesn't exist in the column "${colName}".`);
                  return;
                }
              }
              // get the name_of_the_row in a"name_of_the_row"
              let name = necessarySubformulas[r].slice(2, -1);
              let found = false;
              for(let f = 1; f < nbLineBef; f++) {
                if(values[f][nameInd].includes(name)) {
                  data[i].precedents.push(f);
                  data[f].posteriors.push([i, oper, cat]);
                  found = true;
                  break;
                }
              }
              if(!found) {
                inconsist_removing_elem(i, j, k, `The name at ${getAddress(i, j)} is not found.`);
                return;
              }
            }
          }
        } else if (columnTypes[j] == allColumnTypes.ATTRIBUTES) {
          let attInd = attNames[j].length;
          for(let r = 0; r < attInd; r++) {
            if(attNames[j][r] == val) {
              attInd = r;
              break;
            }
          }
          if(attInd == attNames[j].length) {
            attNames[j].push(val);
            attributes.push(0);
            attInd += accAttr[j];
            acc+=1;
          }
          data[i].attributes.push(attInd);
          attributes[attInd].push(1);
        }
      }
    }
  }
  for(let i = 1; i < nbLineBef; i++) {
    let valuesI = values[i][nameInd];
    for(let k = 0; k < valuesI.length; k++) {
      let val = valuesI[k];
      let refInd_val = checkBrack(i, k, val, false);
      if(refInd_val == -1) {
        return;
      }
      valuesI[k] = refInd_val[1];
      if(refInd_val[0] == -1) {
        // check if the name is already used
        for (let k = 0; k < cellList.length; k++) {
          for (let r = 1; r <= i; r++) {
            for (let f = 0; f < values[r][nameInd].length; f++) {
              if ((r < i || f < k) && cellList[k]==values[r][nameInd][k]) {
                inconsist_replacing_elem(i, nameInd, k, findNewName(cellList[k]), `name "${cellList[k]} at ${getAddress(i, j)} already at ${getAddress(r, nameInd)}`);
                return;
              }
            }
          }
        }
      } else {
        if(values[i][mediaInd]) {
          inconsist_removing_elem(i, nameInd, k, `there is a medium at "${cellList[k]} although the row references an attribute`);
          return;
        }
        for(let r = 0; r < values.length; r++) {
          for(let f = 0; f < data[r].attributes.length; f++) {
            if(data[r].attributes[f] == accAttr[columnInd] + m) {
              for(let cond in data[i]) {
                data[r][cond].concat(data[i][cond]);
              }
            }
          }
        }
      }
      data[i] = undefined;
    }
  }













  for (let i = 1; i < nbLineBef; i++) {
    for (let j = 1; j < 4; j++) {
      for (let k = 0; k < values[i][j].length; k++) {

        //if [category]name is in the follow column
        let isCatF = 0;
        if (j == 1 && elem.startsWith("[")) {
          let endPar = elem.indexOf("]");
          if (endPar != -1) {
            let cat = val.slice(1, endPar);
            var catId = 0;
            for (let n = 0; n < colNumb - 4; n++) {
              for (let r = 0; r < attNames[n].length; r++) {
                if (attNames[n][r] == cat) {
                  stop = true;
                  break;
                }
                catId++;
              }
              if (stop) {
                break;
              }
            }
            if (!stop) {
              suggRef(i, j, k);
              return -1;
            }
            stop = false;
            val = val.slice(endPar + 1);
            isCatF = 1;
          }
        }

        //search for the line of the name
        for (var r = 1; r < values.length; r++) {
          for (let f = 0; f < values[r][0].length; f++) {
            if (searching(r, 0, f)) {
              stop = true;
              break;
            }
          }
          if (stop) {
            break;
          }
        }
        if (!stop) {
          suggRef(i, j, k);
          return -1;
        }
        stop = false;
        if (stat > 0) {
          if (found(r, i, j, k) == -1) {
            return -1;
          }
          if(periods) {
            console.log(`2periods : ${periods[0][0]}`)
          }
        }
        if (isCatF) {
          if (data[r][4].some((v) => v[1] == catId)) {
            suggRef(i, j, k);
            return -1;
          }
          data[i][1].push(r);
          data[r][4].push([i, catId]);
          continue;
        }
        perIntCont[i][j - 1].push([r, stat]);
        values[i][j][k] = perNot[stat] + elem2;
      }
    }
  }
  console.log(" hhhhhhhhh ")
  for (let i = 0; i < periods.length; i++) {
    for (let q = 1; q < 3; q++) {
      if (periods[i][q] != -1) {
        continue;
      }
      periods[i][q] = values.length;
      values.push([]);
      perRef.push(i);
      for (let i = 0; i < colNumb; i++) {
        values[values.length - 1].push([]);
      }
      for (let m = 0; m < periods[i][3].length; m++) {
        values[values.length - 1][0].push(perNot[q] + periods[i][3][m]);
      }
      data.push([[], [], [], -1, []]);
    }
    for (let q = 1; q < 3; q++) {
      values[periods[i][q]][0] = periods[i][3]
        .map((e) => perNot[q] + e)
        .concat(values[periods[i][q]][0].filter((e) => !e.startsWith(perNot[q])));
    }

    //fill start and end with the info on the declaration line
    let z = periods[i][1];
    let s = periods[i][2];
    data[s][1].push(z);
    let c = periods[i][0];
    if (c == -1) {
      c = values.length;
      periods[i][0] = c;
      values.push([]);
      for (let i = 0; i < colNumb; i++) {
        values[c].push([]);
      }
      for (let m = 0; m < periods[i][3].length; m++) {
        values[c][0].push(periods[i][3][m]);
      }
      data.push([[], [], [], z, [], []]);
      data[z][0].push();
      continue;
    }
    data[z][0] = false;
    data[c][3] = z;
  }
  for (let i = 1; i < nbLineBef; i++) {
    for (let j = 0; j < perInt[i].length; j++) {
      if (perRef[i] > -1) {
        data[periods[perRef[i]][0]][1].push(periods[perInt[i][j]][1]);
        data[periods[perInt[i][j]][2]][1].push(periods[perRef[i]][2]);
      } else {
        data[i][1].push(periods[perInt[i][j]][1]);
        data[periods[perInt[i][j]][2]][1].push(i);
      }
    }
  }
  for (let i = 1; i < nbLineBef; i++) {
    for (let j = 0; j < 3; j++) {
      for (let k = 0; k < perIntCont[i][j].length; k++) {
        let u = perIntCont[i][j][k];
        let r = u[0];
        let e = r;
        let iR = r;
        if (perRef[r] > -1) {
          var y = periods[perRef[r]];
          r = y[1];
          e = y[2];
          iR = e;
        }
        if (u[1] == 1) {
          e = r;
          iR = y[0];
        } else if (u[1] == 2) {
          r = e;
        }
        perIntCont[i][j][k] = [r, e, u[1], iR];
        console.log(`perIntCont : ${i} ${j} ${k} ${r} ${e} ${u[1]} ${iR}`)
      }
    }
  }
  console.log(" qqqqqqqqqqqqq ")
  for (let i = 1; i < nbLineBef; i++) {
    if (perIntCont[i][0].length > 1) {
      let interm = [];
      let isIn = [];
      for (let j = 0; j < perIntCont[i][0].length; j++) {
        isIn.push(0);
      }
      for (let p = 0; p < isIn.length - 1; p++) {
        if (!isIn[p]) {
          for (let j = 0; j < isIn.length; j++) {
            console.log(`data[perIntCont[i][0][j][3]] : ${data[perIntCont[i][0][j][3]]}`)
            for (let k = 0; k < data[perIntCont[i][0][j][3]][1].length; k++) {
              for (let m = 0; m < isIn.length; m++) {
                if (j != m && k == perIntCont[i][0][m][3]) {
                  if (!isIn[m]) {
                    stop = true;
                    break;
                  }
                }
              }
              if (stop) {
                break;
              }
            }
            if (stop) {
              interm.push(perIntCont[i][0][j]);
              isIn[j] = 1;
              break;
            }
          }
          if (!stop) {
            for (let j = 0; j < isIn.length; j++) {
              if (!isIn[j]) {
                interm.push(perIntCont[i][0][j]);
                isIn[j] = 1;
              }
            }
          }
        }
      }
      console.log(`interm.length : ${interm.length}`)
      for (let j = 0; j < isIn.length; j++) {
        console.log(`j : ${j}`)
        console.log(`interm[j] : ${interm[j]}`)
        perIntCont[i][0][isIn.length - 1 - j] = interm[j].slice();
      }
    }
  }
  console.log(" aaaaaaaaaaa ")
  for (let i = 1; i < nbLineBef; i++) {
    for (let j = 0; j < 3; j++) {
      for (let k = 0; k < perIntCont[i][j].length; k++) {
        let u = perIntCont[i][j][k];
        let r = u[0];
        let e = u[1];
        let ie = i;
        let iR = i;
        if (perRef[i] > -1) {
          let p = periods[perRef[i]]
          if (i == p[1]) {
            iR = p[0];
          } else if (i == p[0]) {
            ie = p[2];
          }
        }
        if (j == 0) {
          while (data[e][3] != -1) {
            e = data[e][3];
          }
          let firstRed = 1;
          while (!data[iR][0]) {
            console.log(`iR : ${iR}`);
            if (perRef[e] > -1) {
              //set follow rule to followed lines
              var ki = k - 1;
              if (firstRed) {
                firstRed = 0;
              } else {
                ki = perIntCont[iR][0].length - 1;
              }
              if (ki < 0) {
                iR = periods[perRef[iR]][0];
              } else {
                iR = perIntCont[iR][0][ki][3];
              }
            } else {
              values[iR][1].splice(1, values[iR][1].length - 1);
              console.log("sugg3");
              sugg(iR, 1);
              return -1;
            }
          }
          data[e][3] = iR;
          data[iR][0] = false;
        } else if (j == 1) {
          data[iR][1].push(e);
        } else {
          data[r][1].push(ie);
        }
      }
    }
  }

  // --search inconcistencies--

  let inLoop = [false];
  let follower = [0];
  for (let i = 1; i < data.length; i++) {
    if (data[i] === 0) {
      inLoop.push(false);
    } else {
      inLoop.push(true);
    }
    follower.push(0);
  }
  //group lines by followings
  lenAgg = 0;
  dataAgg = [];
  for (let i = 1; i < data.length; i++) {
    try {
      if (!data[i][0] || values[i][0].length == 0) {
        1;
      }
    } catch (error) {
      1;
    }
    if (!data[i][0] || values[i][0].length == 0) {
      continue;
    }
    dataAgg.push({
      lines: [],
      before: new Set(),
      after: new Set(),
      shortBefore: [],
      shortAfter: [],
      attr: [],
      catFollowed: [],
      catFollows: [],
      mediaNb: 0
    });
    let j = i;
    let lenLines = 0;
    let mediaNb = 0;
    while (j !== -1) {
      dataAgg[lenAgg].lines.push(j);
      for (let k = 0; k < data[j][2].length; k++) {
        for (var h = 0; h < dataAgg[lenAgg].attr.length; h++) {
          if (dataAgg[lenAgg].attr[h][0] == data[j][2][k]) {
            dataAgg[lenAgg].attr[h][2] = mediaNb + 1;
            dataAgg[lenAgg].attr[h][3]++;
            dataAgg[lenAgg].attr[h].push(mediaNb);
            stop = true;
            break;
          }
        }
        if (!stop) {
          h = dataAgg[lenAgg].attr.length;
          dataAgg[lenAgg].attr.push([data[j][2][k], mediaNb, mediaNb + 1, 1, mediaNb]);
        }
        if (dataAgg[lenAgg].attr[h][3] == attributes[data[j][2][k]]) {
          let sum = 0;
          for (let m = 4; m < dataAgg[lenAgg].attr[h].length; m++) {
            sum += dataAgg[lenAgg].attr[h][m];
          }
          attributes[data[j][2][k]] = -Math.floor(sum / (dataAgg[lenAgg].attr[h].length - 4));
        }
        stop = false;
      }
      if (!inLoop[j]) {
        let j0 = follower[j];
        let j2 = dataAgg[lenAgg].lines[lenLines];
        inconsist(
          values[j0][0][0],
          0,
          "Lines " +
          values[j0][0][0] +
          " (" +
          (j0 + 1) +
          ")" +
          " and " +
          values[j2][0][0] +
          " (" +
          (j2 + 1) +
          ")" +
          " both follow line " +
          j,
          []
        );
        return -1;
      }
      inLoop[j] = false;
      follower[j] = dataAgg[lenAgg].lines[lenLines];
      lenLines++;
      if (j < nbLineBef && perRef[j] == -1) {
        mediaNb++;
      }
      j = data[j][3];
    }
    dataAgg[lenAgg].mediaNb = mediaNb;
    lenAgg++;
  }
  //check if there is a loop that couldn't have been entered
  for (let i = 1; i < data.length; i++) {
    if (values[i][0].length == 0 || !inLoop[i]) {
      continue;
    }
    let j = i;
    let allLines = "";
    while (1) {
      let index = j + 1;
      if (j >= nbLineBef) {
        index = "out";
      }
      allLines += values[j][0][0] + "(" + index + ") >> ";
      if (j == i) {
        break;
      }
      j = data[j][3];
    }
    inconsist(values[i][0][0], 0, allLines, []);
    return -1;
  }
  console.log("oooooooooooooooooo")
  //add the precedents to dataAgg
  console.log(`steps`);
  place = [];
  steps = [];
  visited = [];
  let inPath = [];
  for (let i = 0; i < lenAgg; i++) {
    place.push([]);
    steps.push(0);
    inPath.push(0);
    visited.push(0);
    for (let k = 0; k < dataAgg[i].lines.length; k++) {
      let f = dataAgg[i].lines[k];
      for (let h = 0; h < data[f][1].length; h++) {
        for (let m = 0; m < lenAgg; m++) {
          let ind = dataAgg[m].lines.indexOf(data[f][1][h]);
          if (ind != -1) {
            if (i == m) {
              if (k <= ind) {
                let allLines = "";
                for (let v = k; v < ind + 1; v++) {
                  let vi = dataAgg[i].lines[v];
                  let index = vi + 1;
                  if (vi >= nbLineBef) {
                    index = "out";
                  }
                  allLines += "\n>> " + values[vi][0][0] + "(" + index + ")";
                }
                let vi = dataAgg[i].lines[k];
                let index = vi + 1;
                if (vi >= nbLineBef) {
                  index = "out";
                }
                inconsist(
                  dataAgg[i].lines[k],
                  0,
                  "   " + allLines.slice(4) + "\n > " + values[vi][0][0] + "(" + index + ")",
                  []
                );
                return -1;
              }
            } else {
              dataAgg[i].after.add(m);
              dataAgg[m].before.add(i);
              if (dataAgg[i].lines[0] == 155) {
                console.log(`add after ${dataAgg[m].lines[0]}`);
              }
            }
            break;
          }
        }
      }
      for (let h = 0; h < data[f][4].length; h++) {
        let cat = data[f][4][h];
        let ri = cat[0];
        cat = cat[1];
        for (let m = 0; m < lenAgg; m++) {
          let ind = dataAgg[m].lines.indexOf(ri);
          if (ind != -1) {
            if (i == m) {
              for (let r = ind + 1; r < k; r++) {
                if (data[dataAgg[i].lines[r]][2].includes(cat)) {
                  let allLines = "";
                  for (let v = int; v < r + 1; v++) {
                    let vi = dataAgg[i].lines[v];
                    let index = vi + 1;
                    if (vi >= nbLineBef) {
                      index = "out";
                    }
                    allLines += "\n >> (" + index + ")" + values[vi][0][0];
                  }
                  inconsist(dataAgg[i].lines[int], 0, "[" + cat + "]\n    " + allLines.slice(5), []);
                  return -1;
                }
              }
            }
            dataAgg[m].catFollows.push(cat);
            dataAgg[i].catFollowed.push([cat, m]);
            break;
          }
        }
      }
    }
  }
  console.log("zzzzzzzzzzzzzzzzzzzzzzzzzzz")
  for (let i = 0; i < lenAgg; i++) {
    if (dataAgg[i].mediaNb == 0) {
      dataAgg[i].mediaNb =
        dataAgg[i].mediaNb > 0 || dataAgg[i].catFollows.size || dataAgg[i].catFollowed.size ? -1 : 0;
    }
  }

  let mi;
  let arrow;
  let inc;
  let stack = [];

  // return;
  for (let objectId = 0; objectId < lenAgg; objectId++) {
    if (dataAgg[objectId].after.size == 0) {
      if (!visited[objectId]) {
        let path = [];
        stack.push(objectId);
        while (stack.length > 0) {
          let back = 0;
          let hasUnvisitedNeighbor = 0;
          const current = stack[stack.length - 1];
          if (path.length != 0 && path[path.length - 1] == current) {
            back = 1;
          } else {
            path.push(current);
            inPath[current] = 1;
            visited[current] = 1;
            const neighbors = Array.from(dataAgg[current].before);
            for (const neighbor of neighbors) {
              if (inPath[neighbor]) {
                let allLines = "";
                let last = -1;
                path.splice(0, path.indexOf(neighbor));
                path.push(neighbor);
                path.push(path[Math.min(1, path.length - 1)]);
                for (let i = 0; i < path.length - 1; i++) {
                  for (let ji = 0; ji < dataAgg[path[i + 1]].lines.length; ji++) {
                    let j = dataAgg[path[i + 1]].lines[ji];
                    for (const k of data[j][1]) {
                      for (let m = 0; m < dataAgg[path[i]].lines.length; m++) {
                        let mo = dataAgg[path[i]].lines[m];
                        if (mo == k && (path.length > 3 || j < m)) {
                          let sub = "";
                          let toAdd = [mo, j];
                          if (last != -1) {
                            toAdd = [j];
                            if (last < m) {
                              arrow = ">>";
                              inc = 1;
                            } else {
                              arrow = "<<";
                              inc = -1;
                            }
                            for (let n = last + 1; n < m + 1; n += inc) {
                              let no = dataAgg[path[i]].lines[n];
                              mi = no + 1;
                              if (mi > nbLineBef) {
                                mi = "out";
                              }
                              sub += "\n" + arrow + " (" + mi + ") " + values[no][0].join("; ");
                            }
                            allLines += sub;
                          }
                          if (i < path.length - 2) {
                            last = ji;
                            for (const a of toAdd) {
                              mi = a + 1;
                              if (a > nbLineBef) {
                                mi = "out";
                              }
                              allLines += "\n>  (" + mi + ") " + values[a][0].join("; ");
                            }
                          }
                          stop = 1;
                          break;
                        }
                      }
                      if (stop) {
                        break;
                      }
                    }
                    if (stop) {
                      stop = 0;
                      break;
                    }
                  }
                }
                inconsist(dataAgg[path[1]].lines[0], 0, "   " + allLines.slice(3), []);
                return -1;
              }
              if (steps[neighbor]) {
                // if there is a link that can be removed
                let lvl = place[neighbor][0];
                dataAgg[neighbor].after.delete(lvl);
                dataAgg[lvl].before.delete(neighbor);
                if (dataAgg[lvl].mediaNb != 0) {
                  dataAgg[neighbor].shortAfter.splice(dataAgg[neighbor].shortAfter.indexOf(lvl), 1);
                } else {
                  dataAgg[lvl].shortAfter.forEach((v) => {
                    dataAgg[neighbor].shortAfter.splice(dataAgg[neighbor].shortAfter.indexOf(v), 1);
                  });
                }
                if (dataAgg[neighbor].mediaNb != 0) {
                  dataAgg[lvl].shortBefore.splice(dataAgg[lvl].shortBefore.indexOf(neighbor), 1);
                } else {
                  dataAgg[neighbor].shortBefore.forEach((v) => {
                    dataAgg[lvl].shortBefore.splice(dataAgg[lvl].shortBefore.indexOf(v), 1);
                    if (dataAgg[lvl].mediaNb != 0) {
                      dataAgg[v].shortAfter.splice(dataAgg[v].shortAfter.indexOf(lvl), 1);
                    } else {
                      dataAgg[lvl].shortAfter.forEach(v2 => {
                        dataAgg[v].shortAfter.splice(dataAgg[v].shortAfter.indexOf(v2), 1);
                      });
                    }
                  });
                }
                if (place[neighbor][1] < stack.length && stack[place[neighbor][1]] == neighbor) {
                  stack.splice(place[neighbor][1], 1);
                }
              } else {
                steps[neighbor] = 1;
              }
              if (visited[neighbor]) {
                if (dataAgg[neighbor].mediaNb != 0) {
                  dataAgg[current].shortBefore.push(neighbor);
                } else {
                  dataAgg[neighbor].shortBefore.forEach((v) => {
                    dataAgg[current].shortBefore.push(v);
                    if (dataAgg[current].mediaNb != 0) {
                      dataAgg[v].shortAfter.push(current);
                    } else {
                      dataAgg[current].shortAfter.forEach((v2) => {
                        dataAgg[v].shortAfter.push(v2);
                      });
                    }
                  });
                }
              } else {
                stack.push(neighbor);
                hasUnvisitedNeighbor = 1;
              }
              place[neighbor] = [current, stack.length];
              if (dataAgg[current].mediaNb != 0) {
                dataAgg[neighbor].shortAfter.push(current);
              } else {
                dataAgg[current].shortAfter.forEach((v) => {
                  dataAgg[neighbor].shortAfter.push(v);
                });
              }
            }
          }
          if (back || !hasUnvisitedNeighbor) {
            inPath[path[path.length - 1]] = 0;
            path.pop();
            if (path.length != 0) {
              if (dataAgg[current].mediaNb != 0) {
                dataAgg[path[path.length - 1]].shortBefore.push(current);
              } else {
                dataAgg[current].shortBefore.forEach((v) => {
                  dataAgg[path[path.length - 1]].shortBefore.push(v);
                });
              }
            }
            for (const neighbor of Array.from(dataAgg[current].before)) {
              steps[neighbor] = 0;
            }
            stack.pop();
          }
        }
      }
    }
  }

  console.log(`color loading...`);

  // Loop through each row in the range
  let maxL = Math.max(nbLineBef, prevLine[sheetCodeName]);
  let prevLineBef = prevLine[sheetCodeName];
  prevLine[sheetCodeName] = 1;

  let color, font;
  for(let j = colors.length; j<=maxL; j++) {
    colors.push("");
    fonts.push("");
  }
  for (let i = 0; i < maxL; i++) {
    if (!i) {
      color = "000000";
      font = "ffffff";
    } else {
      font = "000000";
      color = "";
      if (i < nbLineBef) {
        if (perRef[i] > -1) {
          color = "d97373";
          if (periods[perRef[i]][0] == i) {
            color = "ffa600";
          } else if (periods[perRef[i]][1] == i) {
            color = "93ff00";
          }
        } else if (perRef[i] == -2) {
          color = "73d9cb";
        }
      }
    }
    if(color || font!="000000") {
      prevLine[sheetCodeName] = i + 1;
    }
    if(colors[i]!=color || fonts[i]!=font || updateRegularity) {
      response.push({ "range_updateRegularity": [i + 1, colNumb-1] });
      if(colors[i]!=color || updateRegularity) {
        if(color) {
          response.push({ "color_updateRegularity": [color] });
        } else {
          response.push({ "clear_updateRegularity": [] });
        }
      }
      if(fonts[i]!=font || updateRegularity) {
        response.push({ "font_updateRegularity": [font] });
      }
      colors[i]=color
      fonts[i]=font
    }
  }
  iL = Math.min(prevLine[sheetCodeName], prevLineBef);
  iU = Math.max(prevLine[sheetCodeName], prevLineBef);
  for(let i = iL; i<iU; i++) {
    response.push({ "styleBorders": [i, 0, colNumb, iL == prevLineBef] });
  }
  response.push({ "checkSuccess": [] });
  resolved[sheetCodeName] = true;
  handleSelectLinks();
}

function suggRef(i, j, k) {
  let suggWords = [];
  let suggWords2 = [];
  for (let r = 1; r < nbLineBef; r++) {
    if (perRef[r] > -1 && periods[perRef[r]][0] != r) {
      continue;
    }
    for (let f = 0; f < values[r][0].length; f++) {
      let val2 = values[r][0][f];
      if (val2.includes(val)) {
        suggWords.push(val2);
      } else if (val.includes(val2)) {
        suggWords2.push(val2);
      }
    }
  }
  suggWords.sort(function (a, b) {
    return a.indexOf(val) - b.indexOf(val);
  });
  suggWords2.sort(function (a, b) {
    return val.indexOf(a) - val.indexOf(b);
  });
  suggWords = suggWords.concat(suggWords2).map((v) =>
    values[i][j]
      .slice(0, k)
      .concat([perNot[stat] + v])
      .concat(values[i][j].slice(k + 1))
      .join("; ")
  );
  if (suggWords.length == 0) {
    values[i][j].splice(k, 1);
    console.log("sugg4");
    sugg(i, j);
  } else {
    console.log("suggRef");
    suggSet(i, j, suggWords);
  }
}

function found(r, i, j, k) {
  if (perRef[r] == -1) {
    values[i][j].splice(k, 1);
    sugg(i, j);
    return -1;
  }
  if (perRef[r] == -2) {
    perRef[r] = periods.length;
    periods.push([
      -1,
      -1,
      -1,
      values[r][0].filter((e) => e.startsWith(perNot[stat2])).map((e) => e.slice(perNot[stat2].length))
    ]);
    periods[perRef[r]][stat2] = r;
  }
  if (j == 0) {
    periods[perRef[r]][stat] = i;
    periods[perRef[r]][3] = periods[perRef[r]][3].concat(
      values[i][0]
        .filter((e) => e.startsWith(perNot[stat]))
        .map((e) => e.slice(perNot[stat].length))
        .filter((e) => !periods[perRef[r]][3].map((e) => e.toLowerCase()).includes(e.toLowerCase()))
    );
    perRef[i] = perRef[r];
  }
  if (i != r) {
    values[i][j][k] = perNot[stat] + elem2;
  }
  return 1;
}

function searching(r, d, f) {
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

function inconsist_removing_elem(i, j, k, message) {
  inconsist(i, j, message, [values[i][j].filter((_, m) => m !== k)]);
}

function inconsist_replacing_elem(i, j, k, newElement, message) {
  inconsist(i, j, message, [values[i][j][k].map((item, m) => (m === k ? newElement : item))]);
}

function sugg(i, j) {
  suggSet(i, j, [values[i][j]]);
}

function suggSet(i, j, sugg) {
  let cellAddress = `${getColumnTag(j)}${i + 1}`;
  inconsist(i, j, "Error at " + cellAddress, sugg);
}

function midToWhite(theMsgType = msgType.ERROR) {
  let r = msgTypeColors[theMsgType].r % 256;
  let g = Math.floor(msgTypeColors[theMsgType].g / 256) % 256;
  let b = Math.floor(msgTypeColors[theMsgType].b / 65536) % 256;
  let newR = Math.floor((r + 255) / 2);
  let newG = Math.floor((g + 255) / 2);
  let newB = Math.floor((b + 255) / 2);
  return {r:newR, g:newG, b:newB};
}

function displayReference(row) {
  response[0]["listBoxList"][msgType.RELATIVES].push({color: msgTypeColors[msgType.RELATIVES], msg: "; ".join(values[r][0]), actions: [{action: Actions.Select, address: [row, 0]}]});
}

function inconsist(i, j, message, suggs, theAction = Actions.NewVal, updateColor = false) {
  response[0]["listBoxList"][msgType.ERROR].push({color: msgTypeColors[msgType.ERROR], msg: message, actions: [{action: Actions.Select, address: [i, j]}]});
  suggs.forEach(sugg => {
    let theNewVal = sugg;
    if (theAction == Actions.NewVal) {
      theNewVal = sugg.join("; ");
    }
    response[0]["listBoxList"][msgType.ERROR].push({color: midToWhite(), msg: theNewVal, actions: [{action: Actions.Select, address: 0}, {action: theAction, address: 0, newVal: theNewVal}]});
  });
}

function findNewName(name, checked = true) {
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
    checked = true;
    alreadyExist = false;
    for (let r = 1; r < nbLineBef; r++) {
      for (let f = 0; f < values[r][nameInd].length; f++) {
        if (name == values[i][nameInd][k]) {
          alreadyExist = true;
          break;
        }
      }
      if(alreadyExist) {
        break;
      }
    }
  }
  return name;
}

function correct() {
  for (let i = 1; i < nbLineBef; i++) {
    for (let j = 0; j < colNumb; j++) {
      if (!Array.isArray(values[i][j])) {
        console.log("NOT");
      }
      let value = values[i][j].join("; ");
      if (values0[i][j] != value) {
        values0[i][j] = value;
        response.push({ "chgValue": [i + 1, j + 1, value] });
      }
    }
  }
}

function customSort(a, b) {
  if (a.mediaNb != 0 && b.mediaNb == 0) {
    return -1; // 'a' comes before 'b'
  } else if (a.mediaNb == 0 && b.mediaNb != 0) {
    return 1; // 'b' comes before 'a'
  } else {
    return 0; // maintain original order
  }
}

function dataGeneratorSub() {
  if (!resolved[sheetCodeName]) {
    return;
  }

  correct();


  for (let i = 0; i < lenAgg; i++) {
    steps[i] = 0;
    visited[i] = 0;
  }
  for (let objectId = 0; objectId < lenAgg; objectId++) {
    if (dataAgg[objectId].shortAfter.size == 0 && dataAgg[objectId].mediaNb != 0) {
      if (!visited[objectId]) {
        let path = [];
        stack.push(objectId);
        while (stack.length > 0) {
          let back = 0;
          let hasUnvisitedNeighbor = 0;
          const current = stack[stack.length - 1];
          if (path.length != 0 && path[path.length - 1] == current) {
            back = 1;
          } else {
            path.push(current);
            visited[current] = 1;
            const neighbors = Array.from(dataAgg[current].shortBefore);
            for (const neighbor of neighbors) {
              if (steps[neighbor]) {
                // if there is a link that can be removed
                let lvl = place[neighbor][0];
                dataAgg[neighbor].shortAfter.delete(lvl);
                dataAgg[lvl].shortBefore.delete(neighbor);
                if (place[neighbor][1] < stack.length && stack[place[neighbor][1]] == neighbor) {
                  stack.splice(place[neighbor][1], 1);
                }
              } else {
                steps[neighbor] = 1;
              }
              if (visited[neighbor]) {
                dataAgg[current].shortBefore.push(neighbor);
              } else {
                stack.push(neighbor);
                hasUnvisitedNeighbor = 1;
              }
              place[neighbor] = [current, stack.length];
              dataAgg[neighbor].shortAfter.push(current);
            }
          }
          if (back || !hasUnvisitedNeighbor) {
            path.pop();
            if (path.length != 0) {
              dataAgg[path[path.length - 1]].shortBefore.push(current);
            }
            for (const neighbor of Array.from(dataAgg[current].shortBefore)) {
              steps[neighbor] = 0;
            }
            stack.pop();
          }
        }
      }
    }
  }

  let dataAggCop = dataAgg.slice();
  dataAgg.sort(customSort);
  let sortedPositions = dataAggCop.map(function (obj) {
    return dataAgg.indexOf(obj);
  });
  for (let i in dataAgg) {
    if (dataAgg[i].lines[0] == 11) {
      console.log(`ccl ${Array.from(dataAgg[i].before).join(",")}`);
      console.log(
        `ccl ${Array.from(dataAgg[i].before)
          .map((val) => dataAgg[val].lines.join(","))
          .join(" ")}`
      );
    }
    dataAgg[i].before = new Set(Array.from(dataAgg[i].before).map((element) => sortedPositions[element]));
    dataAgg[i].shortBefore = new Set(dataAgg[i].shortBefore.map((element) => sortedPositions[element]));
    dataAgg[i].catFollowed = dataAgg[i].catFollowed.map((element) => {
      console.log(`CATTTT ${element[1]} ${sortedPositions[element[1]]}`);
      return [element[0], sortedPositions[element[1]]];
    });
  }
  let result = attributes.join(",") + '\n' +
    dataAgg
      .map((row) =>
        [
          row.lines.join(";"),
          new Set(row.shortAfter).size + ";" + row.after.size + ";" + row.mediaNb,
          Array.from(row.shortBefore).join(";"),
          Array.from(row.before).join(";"),
          row.attr.map((row) => row.slice(0, 3).join(",")).join(";"),
          row.catFollows.join(";"),
          row.catFollowed.map((r) => r.join(",")).join(";")
        ].join("\t")
      ).join("\n") + '\n' +
    values
      .map((row) =>
        row
          .slice(0, colNumb)
          .map(cell => cell.join("; "))
          // .map(function (cell) {
          //   if (Array.isArray(cell)) {
          //     return cell.join("; ");
          //   } else {
          //     return cell;
          //   }
          // })
          .join("\t")
      ).join("\n");
  response.push({ "sort": [result] });


}

function renameSymbol(oldValue, newValue) {
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
          if (searching(i, j, k)) {
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
  correct();
  response.push({ "Renamings": [`renamed (${nbRen})`] });
  check();
}

function handleoldNameInputClick() {

  let rowId = range.rowIndex; // Directly use rowIndex property
  let totalRows = range.rowCount; // Get total rows in the selection
  for (let i = 0; i < totalRows; i++) {
    // Add each row index to rowNumbers

    rowId = Math.min(rowId, range.rowIndex);
  }
  response.push({ "oldNameInput": [values[rowId][0].join("; ")] });
  response.push({ "newNameInput": [values[rowId][0].join("; ")] });
}

function ctrlZ() {
  if (indOldValues > 0) {
    indOldValues--;
    values = oldValues[indOldValues];
    correct();
    check();
  }
}

function ctrlY() {
  if (indOldValues < oldValues.length - 1) {
    indOldValues++;
    values = oldValues[indOldValues];
    correct();
    check();
  }
}

function addLineToFile(filePath, line) {
  // Add a newline character at the end of the line
  const lineWithNewline = line + '\n';
  
  // Append the line to the file asynchronously
  fs.appendFile(filePath, lineWithNewline, (err) => {
      if (err) {
          console.error('Error writing to file:', err);
      } else {
          console.log('Line added to file successfully.');
      }
  });
}

// Endpoint to call JavaScript functions
app.post('/execute', (req, res) => {
  const timestamp = req.body.timestamp;
  
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
  
  response = [{"listBoxList": []}];
  for (const key in msgType) {
    response[0]["listBoxList"].push([]);
  }
  let body = req.body;
  const funcName = body.functionName;
  if (funcName === 'handleChange') {
    handleChange(body.changes);
  } else if (funcName === 'selectionChange') {
    selectedCells[sheetCodeName] = body.selection;
    handleSelectLinks();
  } else if (funcName === 'chgSheet') {
    sheetCodeName = body.sheetCodeName;
    headerColors = body.headerColors;
    values = values0[sheetCodeName].map(r => r.map(c => c===null?[]:c.split(';').map(k => k.trim().toLowerCase()).filter(k => k)));
    handleChange([]);
  } else if (funcName == "dataGeneratorSub") {
    dataGeneratorSub();
  } else if (funcName == "renameSymbol") {
    renameSymbol(body.oldValue, body.newValue);
  } else if (funcName == "handleoldNameInputClick") {
    handleoldNameInputClick();
  } else if (funcName == "ctrlZ") {
    ctrlZ();
  } else if (funcName == "ctrlY") {
    ctrlY();
  } else if (funcName == "newSheet") {
    values0[body.sheetCodeName] = body.values;
    sorting[body.sheetCodeName] = false;
    prevLine[body.sheetCodeName] = 0;
  }
  let empty = true;
  for (const msgT in response[0]["listBoxList"]) {
    if(msgT.length) {
      empty = false;
      break;
    }
  }
  if(empty) {
    delete response[0]["listBoxList"];
  }
  res.json(response);
});

function parseExpression(expr) {
    // Replace '&&' with 'and' and '||' with 'or' to match JS syntax
    expr = expr.replace(/&&/g, '&&').replace(/\|\|/g, '||');
    const ast = acorn.parseExpressionAt(expr, 0);
    return buildTree(ast);
}

function buildTree(node) {
    if (node.type === 'LogicalExpression') { // Handles && and ||
        const op = node.operator === '&&' ? '&&' : '||';
        const root = new Node(op);
        root.left = buildTree(node.left);
        root.right = buildTree(node.right);
        return root;
    } else if (node.type === 'BinaryExpression') { // Handles comparisons like A == 1
        return new Node(`${node.left.name} == ${node.right.value}`);
    } else if (node.type === 'Identifier') {
        return new Node(node.name);
    } else if (node.type === 'Literal') {
        return new Node(String(node.value));
    } else {
        throw new Error("Unsupported AST node.");
    }
}

function findNecessarySubformulas(node) {
    if (node.value === '&&') {
        const left = findNecessarySubformulas(node.left);
        const right = findNecessarySubformulas(node.right);
        return left.concat(right);
    } else if (node.value === '||') {
        const left = findNecessarySubformulas(node.left);
        const right = findNecessarySubformulas(node.right);
        const common = left.filter(value => right.includes(value));
        return common;
    } else {
        return [node.value];
    }
}


app.post('/test', (req, res) => {



  // Example usage:
  const formula = "(A == 1) && (B == 2 || (C == 3) && (D == 4))";
  // const formula = "(A == 1) && ((C == 3) && (B == 2) || (C == 3) && (D == 4))"
  // const formula = "(A == 1) && (B == 2) || (C == 3) && (D == 4)"
  // const formula = "(A == 1) && (B == 2 || (C == 3) && (D == 4))"
  // const formula = "(A == 1) && (C == 3) || (C == 3) && (D == 4)"
  const tree = parseExpression(formula);
  const necessarySubformulas = findNecessarySubformulas(tree);

  console.log("Necessary Subformulas:", necessarySubformulas);

  
  res.json("response");
});

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});