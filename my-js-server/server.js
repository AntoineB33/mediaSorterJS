const express = require('express');
const { exec } = require('child_process');

const app = express();
const port = 3000;


// Middleware to parse JSON bodies
app.use(express.json());


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
          response.append({ "addItem": [periods[perRef[row]][g]] });
        }
      }
    }
  } else {
    for(let k = 0; k < values[row][column].length; k++) {
      if(isTip(row, column, k) == -1) {
        return -1;
      }
      if(column < 4) {
        for(let r = 1; r < values.length; r++) {
          for(let f = 0; f < values[r][0].length; f++) {
            if(row!=r && searching(r, 0, f)) {
              if(perRef[r] > -1) {
                for(let g = 0; g < 3; g++) {
                  if(periods[perRef[r]][g]!=-1) {
                    response.append({ "addItem": [r] });
                  }
                }
              } else {
                response.append({ "addItem": [r] });
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
                response.append({ "addItem": [r] });
              }
            }
          }
        }
      }
    }
  }
}

function handleChange(updates) {
  updates.forEach(function(update) {
    var row = update[0] - 1;  // Adjusting for zero-based index
    var column = update[1] - 1;  // Adjusting for zero-based index
    var value = update[2];
    values0[sheetCodeName][row][column] = value;
    var splitValue = value;
    if(value) {
      splitValue = value
          .split(";")
          .map(function (v, i) {
            return v.trim().toLowerCase();
          })
          .filter(function (e) {
            return e !== "";
          });
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
  check();
  values.splice(nbLineBef, nbLineBefBef - nbLineBef);
  oldValues.splice(indOldValues, oldValues.length - 1 - indOldValues);
  oldValues.push(values);
  if(oldVersionsMaxNb < oldValues.length) {
    oldValues.shift();
  }
}

function getAddress(i, j) {
  return getColumnTag(j) + ":" + (i + 1);
}

function check() {
  resolved[sheetCodeName] = false;
  
  data = [0];
  var r;

  for (let i = 1; i < nbLineBef; i++) {
    data.push({"follows":[], "comesAfter":[]});
    data[-1][allColumnTypes.ATTRIBUTES] = [];
    data[-1][allColumnTypes.CONDITIONS] = [];

    let cellList = values[i][nameInd]
    if(!cellList.length) {
      for (let j = 0; j < values[i].length; j++) {
        if (values[i][j].length !== 0) {
          inconsist(i, nameInd, `missing name at ${getAddress(i, nameInd)}`, [[findNewName("")]]);
          return;
        }
      }
    }
    for (let k = 0; k < cellList.length; k++) {
      //check if name already before
      for (let r = 1; r <= i; r++) {
        for (let f = 0; f < values[r][nameInd].length; f++) {
          if ((r < i || f < k) && cellList[k]==values[r][nameInd][k]) {
            inconsist_replacing_elem(i, nameInd, k, findNewName(cellList[k]), `name "${cellList[k]} at ${getAddress(i, j)} already at ${getAddress(r, nameInd)}`);
            return;
          }
        }
      }
    }
    
    let valuesI = values[i][nameInd];
    for(let k = 0; k < valuesI.length; k++) {
      let val = sheet[k];
      let countOpenBr = (val.match(new RegExp(`\\[`, 'g')) || []).length;
      let countCloseBr = (val.match(new RegExp(`\\]`, 'g')) || []).length;
      if(countOpenBr || countCloseBr) {
        let ind = val.indexOf('[');
        if(val[-1] != ']' || countOpenBr != 1 || countCloseBr != 1 || ind == 0) {
          val = val.replace(new RegExp(`[\\[\\]]`, 'g'), replacementChar);
          inconsist_replacing_elem(i, nameInd, k, findNewName(val, false), `The name at ${getAddress(i, nameInd)} includes square brackets, thus should be in the format "attribute[column title]".`)
          return -1;
        }
        valuesI[k] = val.split('[');
        let columnTitle = valuesI[k][0].slice(0, -1);
        let found = false;
        for(let j = 0; j < colNumb; j++) {
          if(columnTypes[j] == allColumnTypes.ATTRIBUTES && values[0][j].includes(columnTitle)) {
            found = true;
            break;
          }
        }
        if(!found) {
          inconsist_replacing_elem(i, nameInd, k, findNewName(val, false), `The column title "${columnTitle}" in the name at ${getAddress(i, nameInd)} is not found in the attributes.`)
          return -1;
        }
        valuesI[k][0] = valuesI[k][0].trim();
        valuesI[k] = valuesI[k][0].trim() + (valuesI[k][0].includes(" ")?" ":"") + "[" + valuesI[k][1];
      }
    }
  }
  stop = false;
  let attNames = [];
  attributes = [];
  let acc = 0;
  for (let j = 0; j < colNumb; j++) {
    attNames.push([]);
    for (let i = 1; i < nbLineBef; i++) {
      for (let k = 1; k < values[i][j].length; i++) {
        let val = values[i][j][k].toLowerCase();
        if (columnTypes[j] == allColumnTypes.CONDITIONS) {
          
        } else if (columnTypes[j] == allColumnTypes.ATTRIBUTES) {
          for(let r = 0; r < attNames[j].length; r++) {
            if(attNames[j][r] == val) {

            }
          }
          data[i][allColumnTypes.ATTRIBUTES]
        }
      }
      for (let k = 0; k < values[i][j].length; k++) {
        val = values[i][j][k].toLowerCase();
        for (let r = 1; r < nbLineBef; r++) {
          for (let f = 0; f < values[r][nameInd].length; f++) {
            if (val==values[i][nameInd][k].toLowerCase()) {
            }
          }
        }
        if (values[0][j][k] == headerNames[allColumnNames.PERIODS]) {
          for (let r = 0; r < attNames[j - 4].length; r++) {
            if (attNames[j - 4][r] == val) {
              data[i][2].push(acc + r);
              attributes[acc + r]++;
              stop = true;
              break;
            }
          }
          if (!stop) {
            data[i][2].push(acc + attNames[j - 4].length);
            attNames[j - 4].push(val);
            attributes.push(1);
          }
        } else {
          for (let r = 1; r < nbLineBef; r++) {
            for (let f = 0; f < values[r][0].length; f++) {
              if (!searching(r, 0, f)) {
                continue;
              }
              if (found(context, r, i, j, k) == -1) {
                return -1;
              }
              if(periods) {
                console.log(`1periods : ${periods[0][0]}`)
              }
              perInt[i].push(perRef[r]);
              stop = true;
              break;
            }
            if (stop) {
              break;
            }
          }
          if (!stop) {
            suggRef(context, i, j, k);
            return -1;
          }
        }
        stop = false;
      }
    }
    acc += attNames[j - 4].length;
  }
  for (let i = 1; i < nbLineBef; i++) {
    for (let j = 1; j < 4; j++) {
      for (let k = 0; k < values[i][j].length; k++) {
        if(isTip(i, j, k)) {
          return -1;
        }

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
      data[z][0].append();
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
      response.append({ "range_updateRegularity": [i + 1, colNumb-1] });
      if(colors[i]!=color || updateRegularity) {
        if(color) {
          response.append({ "color_updateRegularity": [color] });
        } else {
          response.append({ "clear_updateRegularity": [] });
        }
      }
      if(fonts[i]!=font || updateRegularity) {
        response.append({ "font_updateRegularity": [font] });
      }
      colors[i]=color
      fonts[i]=font
    }
  }
  iL = Math.min(prevLine[sheetCodeName], prevLineBef);
  iU = Math.max(prevLine[sheetCodeName], prevLineBef);
  for(let i = iL; i<iU; i++) {
    response.append({ "styleBorders": [i, 0, colNumb, iL == prevLineBef] });
  }
  response.append({ "checkSuccess": [] });
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
  values[i][j].splice(k, 1);
  inconsist(i, j, [values[i][j].join("; ")], message);
}

function inconsist_replacing_elem(i, j, k, newElement, message) {
  values[i][j][k] = newElement;
  inconsist(i, j, [values[i][j].join("; ")], message);
}

function sugg(i, j) {
  suggSet(i, j, [values[i][j].join("; ")]);
}

function suggSet(i, j, sugg) {
  let cellAddress = `${getColumnTag(j)}${i + 1}`;
  inconsist(i, j, "Error at " + cellAddress, sugg);
}

function inconsist(i, j, message, sugg, chgBckCol = false) {
  response.append({ "linkToCell": [j + 1, i + 1, message, sugg, chgBckCol] });
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
        response.append({ "chgValue": [i + 1, j + 1, value] });
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
  response.append({ "sort": [result] });


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
  response.append({ "Renamings": [`renamed (${nbRen})`] });
  check();
}

function handleoldNameInputClick() {

  let rowId = range.rowIndex; // Directly use rowIndex property
  let totalRows = range.rowCount; // Get total rows in the selection
  for (let i = 0; i < totalRows; i++) {
    // Add each row index to rowNumbers

    rowId = Math.min(rowId, range.rowIndex);
  }
  response.append({ "oldNameInput": [values[rowId][0].join("; ")] });
  response.append({ "newNameInput": [values[rowId][0].join("; ")] });
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

function getColumnColor(body, headerColors) {
  // Extract the header row
  const headers = body.values[0];

  // Find the index of the column titled "name"
  const nameColumnIndex = headers.indexOf("name");

  // If the column is found, get the corresponding color from headerColors
  if (nameColumnIndex !== -1) {
    return headerColors[nameColumnIndex];
  } else {
    // If the column is not found, return a default value or handle the error

    return null;
  }
}

function checkSquarBrackets(i) {
  return 1;
}

function executeJS(body) {
  response = [];
  const funcName = body.functionName;
  if (funcName === 'handleChange') {
    handleChange(body.changes);
  } else if (funcName === 'selectionChange') {
    selectedCells[sheetCodeName] = body.selection;
    handleSelectLinks();
  } else if (funcName === 'chgSheet') {
    values = values0[sheetCodeName];
    sheetCodeName = body.sheetCodeName;

    // Label the columns "names" and "media"
    headerNames = {};
    let namesColor;
    let alreadyLabelled;
    let missingNamesNb = Object.keys(allColumnNames).length;
    resolved[sheetCodeName] = false;
    for(let j = 0; j < values[0].length; i++) {
      alreadyLabelled = "";
      values[0][j] = values[0][j].split(";").map((value) => value.trim()).filter((value) => value !== "");
      for(let k = 0; k < values[0][j].length; k++) {
        let name = values[0][j][k];
        if(name in allColumnNames) {
          if(alreadyLabelled) {
            inconsist_removing_elem(0, j, k, `Column ${j} labelled as "${alreadyLabelled}" and ${name}.`);
            return;
          }
          if(headerNames[name] !== undefined) {
            inconsist_removing_elem(0, j, k, `Column ${j} labelled as "${name}" already at column ${headerNames[name]} .`);
            return;
          }
          if(namesColor !== undefined && namesColor !== headerColors[j]) {
            inconsist(0, j, `Column ${j} must have the same background color as the attributes.`, [namesColor], true);
            return;
          }
          if(name == allColumnNames.NAMES) {
            nameInd = j;
          }
          missingNamesNb--;
          namesColor = headerColors[j];
          alreadyLabelled = name;
          headerNames[name] = alreadyLabelled;
        }
      }
    }
    if(missingNamesNb) {
      inconsist(0, 0, `Missing columns: ${Object.keys(allColumnNames).filter((name) => headerNames[name] === undefined).join(", ")}.`, []);
      return;
    }
    if (namesColor === undefined) {
      inconsist(0, 0, `No "names" header found`, [allColumnNames.NAMES]);
    }
    columnTypes = [];
    let condColor;
    let headerColors = body.headerColors;
    for (let j = 0; j < headerColors.length; j++) {
      if (headerColors[j] === namesColor) {
        columnTypes.push(allColumnTypes.ATTRIBUTES);
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
    check();
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
    values0[sheetCodeName] = body.values;
    sorting[sheetCodeName] = false;
    prevLine[sheetCodeName] = 0;
  }
}

// Endpoint to call JavaScript functions
app.post('/execute', (req, res) => {
  executeJS(req.body);
  res.json(response);
});

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});