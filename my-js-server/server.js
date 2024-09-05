const express = require('express');
const fs = require('fs');
const { exec } = require('child_process');
const path = require('path');
const abcPath = path.resolve('../abc-master/abc-master/abc');
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
var attrOfAttr;
var headerNames;
var nameSt = "names";
var nameInd;
var mediaSt = "media";
var mediaInd = -1;
var headerColors;
var attNames;

// [node.js]
var response;


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

async function handleChange(updates) {
  nbLineBef = values.length;
  if(nbLineBef) {
    colNumb = values[0].length;
  } else {
    colNumb = 0;
  }
  for(let i = 0; i < updates.length; i++) {
    var row = updates[i][0] - 1;  // Adjusting for zero-based index
    var column = updates[i][1] - 1;  // Adjusting for zero-based index
    var value = updates[i][2];
    if(value) {
      while(headerColors.length <= column) {
        headerColors.push(null);
        columnTypes.push(null);
      }
      while(values.length <= row) {
        values.push([]);
        values0[sheetCodeName].push([]);
      }
      for (let i = 0; i < values.length; i++) {
        while (values[i].length <= colNumb) {
          values[i].push([]);
          values0[sheetCodeName][i].push(null);
        }
      }
      let splitValue = [];
      splitValue = value
          .split(";")
          .map(v => v.trim().toLowerCase())
          .filter(e => e !== "");
      values[row][column] = splitValue;
      values0[sheetCodeName][row][column] = value;
    } else {
      values[row][column] = [];
      values0[sheetCodeName][row][column] = value;
      const nbLineBefBef = nbLineBef;
      for (let i = nbLineBef - 1; i > 0; i--) {
        if (values[i].some(v => v.length != 0)) {
          break;
        }
        nbLineBef--;
      }
      values.splice(nbLineBef, nbLineBefBef - nbLineBef);
      values0[sheetCodeName].splice(nbLineBef, nbLineBefBef - nbLineBef);
      let colNumbMax = 0;
      for(let i = 0; i < nbLineBef; i++) {
        let colNumbRow = colNumb;
        for(let j = colNumb - 1; j > -1; j--) {
          if(values[i][j].length == 0) {
            colNumbRow--;
          } else {
            break;
          }
        }
        colNumbMax = Math.max(colNumbMax, colNumbRow);
      }
      if(colNumbMax != colNumb) {
        for(let i = 0; i < nbLineBef; i++) {
          values[i].splice(colNumbMax, colNumb - colNumbMax);
          values0[sheetCodeName][i].splice(colNumbMax, colNumb - colNumbMax);
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
        inconsist_removing_elem(0, j, k, `Column ${getColumnTag(j)} labelled as "${alreadyInd[0]}" already at column ${getColumnTag(alreadyInd[1])}.`);
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



  await check();
  oldValues.splice(indOldValues, oldValues.length - 1 - indOldValues);
  oldValues.push(values);
  if(oldVersionsMaxNb < oldValues.length) {
    oldValues.shift();
  }
}

function getAddress(i, j) {
  return getColumnTag(j) + ":" + (i + 1);
}

function checkBrack(i, j, k, val, columnTitle) {
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
      inconsist_removing_elem(i, j, k, `No attribute column named "${columnTitle}" at ${getAddress(i, j)}.`)
      return -1;
    }
    for(let m = 0; m < attNames[columnInd].length; m++) {
      if(attNames[columnInd][m] == val) {
        isAtt = m;
        break;
      }
    }
    if(isAtt == -2) {
      inconsist_replacing_elem(i, j, k, `The column named "${columnTitle}" at ${getAddress(i, j)} doesn't have the attribute "${val}".`)
      return -1;
    }
  } else {
    for(let key in attNames) {
      for(let m = 0; m < attNames[key].length; m++) {
        if(attNames[key][m] == theVal) {
          if(isAtt != -2) {
            inconsist_replacing_elem(i, nameInd, k, findNewName(val, false), `The attribute at ${getAddress(i, nameInd)} exists in multiple columns.`);
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

// Function to convert the Formula object into a nested list structure
function formulaToList(formula, reverseReplacements) {
  if (formula.isVar()) {
      return formula.name; // Return the variable name as a string
  } else if (formula.isNot()) {
      return [0, "not", formulaToList(formula.formula, reverseReplacements)]; // Negation with nested formula
  } else if (formula.isAnd()) {
      return ["and", ...formula.terms.map(term => formulaToList(term, reverseReplacements))]; // Conjunction of terms
  } else if (formula.isOr()) {
      return ["or", ...formula.terms.map(term => formulaToList(term, reverseReplacements))]; // Disjunction of terms
  } else if (formula.isTrue()) {
      return "true"; // Return "true" for true constant
  } else if (formula.isFalse()) {
      return "false"; // Return "false" for false constant
  } else {
      throw new Error("Unknown formula type");
  }
}

function initializeData() {
  for (let i = 1; i < nbLineBef; i++) {
    data.push({attributes: new Set(), conditions: [], reverseReplacements: [], newWords: [], posteriors: [], ulteriors: 0, minPos: 0});
    // check if a non-empty row has no name
    if(!values[i][nameInd].length) {
      for (let j = 0; j < values[i].length; j++) {
        if (values[i][j].length !== 0) {
          inconsist(i, nameInd, `missing name at ${getAddress(i, nameInd)}`, [[findNewName("")]]);
          return -1;
        }
      }
      data[i] = -1;
      continue;
    }
  }
}

function getAttributes() {
  // attributes
  for (let j = 0; j < colNumb; j++) {
    let colTitle = values0[sheetCodeName][0][j];
    if(colTitle !== null){
      colTitle = colTitle.split(";");
    } else {
      colTitle = [];
    }
    for(let k = 0; k < colTitle.length; k++) {
      if(colTitle[k].match(/^[A-Z]+$/)) {
        inconsist_replacing_elem(0, j, k, findNewName(colTitle[k]), `Column name at ${getAddress(0, j)} is only upper case letters.`);
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
          data[i].attributes.add(attInd);
        }
      }
    }
  }
}

function getNames() {
  for(let i = 1; i < nbLineBef; i++) {
    let valuesI = values[i][nameInd];
    for(let k = 0; k < valuesI.length; k++) {
      let val = valuesI[k];
      const regexName = /(?<val>[^\[\]]+)(?:\s*\[\s*\"(?<columnTitle>[^\[\]]+)\s*\])?/g;
      let match = regexName.exec(val);
      let refInd_val;
      if(match !== null) {
        refInd_val = checkBrack(i, nameInd, k, match.groups.val, match.groups.columnTitle);
        if(refInd_val == -1) {
          return -1;
        }
        valuesI[k] = '"' + val + '"' + (match.groups.columnTitle===undefined?"":'["' + match.groups.columnTitle + '"]');
      }
      if(match === null || refInd_val === -2) {
        // check if the name is already used
        for (let r = 1; r <= i; r++) {
          for (let f = 0; f < values[r][nameInd].length; f++) {
            if ((r < i || f < k) && valuesI[k]==values[r][nameInd][k]) {
              inconsist_replacing_elem(i, nameInd, k, findNewName(valuesI[k]), `name "${valuesI[k]}" at ${getAddress(i, nameInd)} already at ${getAddress(r, nameInd)}`);
              return -1;
            }
          }
        }
      } else {
        attrOfAttr[refInd_val] = data[i].attributes;
        data[i].isAtt = refInd_val;
      }
    }
  }
  for(let i = 1; i < nbLineBef; i++) {
    if(data[i].isAtt !== undefined) {
      continue;
    }
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

function exprToList(expr) {
  // Remove any extra whitespace for easier processing
  expr = expr.replace(/\s+/g, '');

  // Helper function to check if a character is a letter (a variable in the expression)
  function isLetter(c) {
      return c >= 'a' && c <= 'z' || c >= 'A' && c <= 'Z';
  }

  // Recursive function to process the expression
  function parseExpression(expr) {
      let stack = [];
      let i = 0;

      while (i < expr.length) {
          let char = expr[i];

          if (char === '(') {
              // Find the matching closing parenthesis
              let j = i;
              let parenCount = 0;
              while (j < expr.length) {
                  if (expr[j] === '(') parenCount++;
                  if (expr[j] === ')') parenCount--;
                  if (parenCount === 0) break;
                  j++;
              }

              // Recursively parse the expression within the parentheses
              stack.push(parseExpression(expr.slice(i + 1, j)));
              i = j;
          } else if (isLetter(char)) {
              // Push the variable as a string
              stack.push(char);
          } else if (char === '&') {
              stack.push(-1); // Represent AND with -1
          } else if (char === '|') {
              stack.push(0); // Represent OR with 0
          } else if (char === '~') {
              stack.push(2); // Represent NOT with 2
          }

          i++;
      }

      // If there's only one item in the stack, return it as the result
      if (stack.length === 1) {
          return stack[0];
      }

      return stack;
  }

  // Parse the entire expression
  return parseExpression(expr);
}

function addFormula(i, formula) {
  for (let j = 0; j < colNumb; j++) {
    if (columnTypes[j] == allColumnTypes.CONDITIONS) {
      let colPattern = values[0][j];
      if(colPattern.length) {
        colPattern = colPattern[0].split("?!");
      }
      if (values[i][j].length + (colPattern.length?1:-1) > colPattern.length) {
        inconsist_removing_elem(i, j, k, `Too much arguments at ${getAddress(i, j)}.`);
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

async function getConditions() {
  for (let i = 1; i < nbLineBef; i++) {
    if(data[i].isAtt !== undefined) {
      continue;
    }
    let formula = [];
    addFormula(i, formula);
    for (let a of data[i].attributes) {
      addFormula(a, formula);
    }
    formula = formula.join(" && ");
    let replacedWithWords;
    let reverseReplacements;
    try{
      let reverseReplacementsI;
      ({replacedWithWords, reverseReplacementsI} = replaceWithSameWord(formula));
      reverseReplacements = reverseReplacementsI;
    } catch(error) {
      inconsist(i, j, error.message, []);
      return;
    }
    
    let direct_terms;
    try {

      // Await the result from the Python script
      direct_terms = JSON.parse(await runPythonScript("test_direct", replacedWithWords));

      // Send the result back to the client
      console.log("result");
    } catch (err) {
      inconsist_removing_elem(i, j, 0, `incorrect condition : ${err.message}.`);
      return -1;
    }
    


    let formulaList = exprToList(replacedWithWords);
    let stack = [...formulaList];
    let output = [];
    
    while (stack.length > 0) {
      let current = stack.shift();
      
      if (Array.isArray(current)) {
        // Push the current array elements back to the stack, maintaining order
        stack.unshift(...current);
      } else {
        // Process the non-array element (assuming it needs to be stored in the output)
        let groups = reverseReplacements[current];
        if(groups.after !== undefined) {
          if(groups.max_d !== undefined) {
            if(groups.min_d > groups.max_d) {
              inconsist(i, j, `min_d greater than max_d.`, []);
              return -1;
            }
            if(groups.min_d === 0 && groups.max_d === 0) {
              inconsist(i, j, `min_d and max_d are both 0.`, []);
              return -1;
            }
          }
          if(groups.attrib_ref !== undefined) {
            let refInd_val = checkBrack(i, j, 0, groups.attrib_ref, groups.attrib_ref_col);
            if(refInd_val == -1) {
              return -1;
            }
            if(refInd_val == -2) {
              inconsist_removing_elem(i, j, k, `attribute "${val}" at ${getAddress(i, j)} not found.`);
              return -1;
            }
            groups.attrib_ref = refInd_val;
          }
          let refInd_val = checkBrack(i, j, 0, groups.precedent, groups.precedent_col);
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
                  groups.precedent = m;
                  found = true;
                  break;
                }
              }
              if(found) {
                break;
              }
            }
            if(!found) {
              inconsist_removing_elem(i, j, 0, `precedent "${groups.precedent}" at ${getAddress(i, j)} not found.`);
              return -1;
            }
          }
        }
        output.push(groups);
      }
    }
    
    for(let j = 0; j < direct_terms.lenght; j++) {
      let groups = reverseReplacements[direct_terms[j]];
      if(groups.after !== undefined) {
        if(groups.max_d !== undefined && groups.max_d < 1) {
          groups.min_d = -groups.max_d;
          if(groups.min_d === undefined) {
            groups.max_d = undefined;
          } else {
            groups.max_d = -groups.min_d;
          }
          replacedWithWords = replacedWithWords.replace(direct_terms[j], "true ");
          data[groups.precedent].posteriors.append([i, groups.min_d, groups.max_d]);
          data[i].ulteriors++;
        } else if(groups.min_d !== undefined && groups.min_d > -1) {
          data[i].posteriors.append([groups.precedent, groups.min_d, groups.max_d]);
          data[groups.precedent].ulteriors++;
        }
      }
    }
    data[i].reverseReplacements = reverseReplacements;
    data[i].replacedWithWords = replacedWithWords;
  }

  for (let i = 1; i < nbLineBef; i++) {
    if(data[i].ulteriors !== 0) {
      continue;
    }
    


    
    const inDegree = Array(nbLineBef).fill(0);
    const maxDistance = Array(nbLineBef).fill(0);

    // Create adjacency list and in-degree list
    const adjList = Array.from({ length: nbLineBef }, () => []);

    // Build graph and calculate in-degree for each node
    for (let i = 0; i < nbLineBef; i++) {
        for (const [neighbor, distance] of graph[i]) {
            adjList[i].push([neighbor, distance]);
            inDegree[neighbor]++;
        }
    }

    // Queue to store nodes with in-degree of 0 (nodes with no dependencies)
    const queue = [];

    // Initialize the queue with nodes that have no incoming edges
    for (let i = 0; i < nbLineBef; i++) {
        if (inDegree[i] === 0) {
            queue.push(i);
        }
    }

    // Process nodes in topological order
    while (queue.length > 0) {
        const node = queue.shift();

        // Update the maxDistance for each neighbor
        for (const [neighbor, distance] of adjList[node]) {
            maxDistance[neighbor] = Math.max(maxDistance[neighbor], maxDistance[node] + distance);

            // Decrease in-degree and if it becomes zero, add it to the queue
            inDegree[neighbor]--;
            if (inDegree[neighbor] === 0) {
                queue.push(neighbor);
            }
        }
    }
  }

  for (let i = 1; i < nbLineBef; i++) {
    if(data[i].isAtt !== undefined) {
      continue;
    }
    let reverseReplacements = data[i].replacedWithWords;
    let simplified;
    try {

      // Await the result from the Python script
      simplified = JSON.parse(await runPythonScript("simplify_expression", reverseReplacements));

      // Send the result back to the client
      console.log("result");
    } catch (err) {
      inconsist_removing_elem(i, j, 0, `incorrect condition : ${err.message}.`);
      return -1;
    }
    let formulaList = exprToList(simplified);
    let stack = [...formulaList];
    let output = [];
    
    while (stack.length > 0) {
      let current = stack.shift();
      
      if (Array.isArray(current)) {
        // Push the current array elements back to the stack, maintaining order
        stack.unshift(...current);
      } else {
        // Process the non-array element (assuming it needs to be stored in the output)
        let groups = reverseReplacements[current];
        output.push(groups);
      }
    }
    data[i].conditions = output;
  }
}

function getSimplifiedCond() {

}

function applyCondAttToAllRows() {
  // apply conditions to an attribute to all the rows having the attribute
  for(let i = 1; i < nbLineBef; i++) {
    if(!Number.isInteger(data[i])) {
      for(let j = data[i].conditions.length - 1; j > -1; j--) {
        let cond = data[i].conditions[j];
        let att = -1;
        if(cond.precedent_is_att) {
          att = cond.precedent;
        } else if(Number.isInteger(data[cond])) {
          att = data[cond];
        }
        if(att !== -1) {
          for(let m of attributes[att]) {
            cond.precedent = m;
            data[i].conditions.push(cond);
          }
          data[i].conditions.splice(j, 1);
        }
      }
    }
  }
}

async function check() {
  resolved[sheetCodeName] = false;
  let found;
  
  data = [-1];
  var r;

  initializeData();
  stop = false;
  attributes = [];
  let acc = 0;
  let accAttr = {};

  getAttributes();

  getNames();

  await getConditions();


  applyCondAttToAllRows();

  // change negative distances to positive distances
  for(let i = 1; i < nbLineBef; i++) {
    if(!Number.isInteger(data[i])) {
      for(let j = 0; j < data[i].conditions.length; j++) {
        let cond = data[i].conditions[j];
        if(cond.max_d !== undefined && cond.max_d < 1) {
          cond.min_d = -cond.max_d;
          if(cond.min_d === undefined) {
            cond.max_d = undefined;
          } else {
            cond.max_d = -cond.min_d;
          }
          data[cond.precedent].conditions.push(cond);
          data[i].conditions.splice(j, 1);
        }
        if(cond.min_d > -1) {
          data[i].minPos = Math.max(data[i].minPos, cond.min_d);
        }
      }
    }
  }

  // check if there is a cycle
  let cycle = findCycle(data.map(row => {conditions: row.conditions.filter(condition => condition.min_d !== undefined && condition.min_d > -1)}));
  if(cycle !== null) {
    inconsist(cycle[0], 0, `Cycle detected: ${cycle.map(index => `${values[index][nameInd][0]} (${index})`).join(" -> ")}.`, []);
    return -1;
  }
}












function check0() {
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
  inconsist(i, j, message, [values[i][j].map((item, m) => (m === k ? newElement : item))]);
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

function inconsist(i, j, message, suggs, theAction = Actions.NewVal) {
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

function correct() {
  for (let i = 1; i < nbLineBef; i++) {
    for (let j = 0; j < colNumb; j++) {
      if (!Array.isArray(values[i][j])) {
        console.log("NOT");
      }
      let value = values[i][j].join("; ");
      if (values0[sheetCodeName][i][j] != value) {
        values0[sheetCodeName][i][j] = value;
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

async function renameSymbol(oldValue, newValue) {
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
  await check();
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

async function ctrlZ() {
  if (indOldValues > 0) {
    indOldValues--;
    values = oldValues[indOldValues];
    correct();
    await check();
  }
}

async function ctrlY() {
  if (indOldValues < oldValues.length - 1) {
    indOldValues++;
    values = oldValues[indOldValues];
    correct();
    await check();
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

function findNecessarySubformulas0(node) {
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

function findSubformulas(node) {
  if (node.value === '&&') {
      const left = findSubformulas(node.left);
      const right = findSubformulas(node.right);
      return {
          necessary: left.necessary.concat(right.necessary),
          rest: left.rest.concat(right.rest)
      };
  } else if (node.value === '||') {
      const left = findSubformulas(node.left);
      const right = findSubformulas(node.right);
      const common = left.necessary.filter(value => right.necessary.includes(value));
      const restLeft = left.necessary.filter(value => !common.includes(value)).concat(left.rest);
      const restRight = right.necessary.filter(value => !common.includes(value)).concat(right.rest);
      return {
          necessary: common,
          rest: [node].concat(restLeft).concat(restRight)
      };
  } else {
      return {
          necessary: [node.value],
          rest: []
      };
  }
}

function replaceWithSameWord(str) {
  const replacements = {}; // Object to store replacements
  const reverseReplacements = {}; // Object to store reverse mapping
  let wordCounter = 0; // Counter to generate new words

  const trimmedResult  = str.split(/(\s*&&\s*|\s*\|\|\s*|\s*!\s*|\s*\(\s*|\s*\)\s*)/).filter(Boolean);
  const operators = new Set(["&&", "||", "!", "(", ")"]);
  
  const nameRow = "[^\[\]&&||!]+?";
  let regex = /(?<after>a\s*(?!_\s*[^\d\s-]*$)(?<min_d>-?\d+)?\s*_?\s*(?<max_d>-?\d+)?\s*(?:\s*"\s*(?<attrib_ref>nameRow)\s*"\s*(?:\[\s*"\s*(?<attrib_ref_col>nameRow)\s*"\s*\])?)?\s*"\s*(?<precedent>nameRow)\s*"\s*(?:\[\s*"\s*(?<precedent_col>[^\[\]]+)\s*"\s*\])?(?<isAny>\(any\))?)|(?:p(?<position>-?\d+))/g;
  regex = regex.replace(/nameRow/g, nameRow);
  const replacedWithWords = trimmedResult.forEach(item => {
    if (!operators.has(item)) {
      const match = str.match(regex);
      if(!match) {
        return Error("incorrect condition.");
      }
      const groups = match.groups; // The last element in args array is the groups object

      if (!replacements[match[0]]) {
        const newWord = `word${wordCounter++} `;
        replacements[match[0]] = newWord;

        // Store the reverse mapping along with the capturing groups
        groups.max_d = groups.min_d;
        if(groups.d2 !== undefined) {
          groups.max_d = groups.max_d;
        }
        reverseReplacements[newWord] = {};
        
        // Loop through each key in the groups object and add it to reverseReplacements[newWord].groups
        for (let key in groups) {
            if (groups.hasOwnProperty(key)) {
                reverseReplacements[newWord][key] = groups[key];
            }
        }
      }
      return replacements[match[0]];
    } else {
      return item;
    }
  });

  return { replacedWithWords, reverseReplacements };
}

function findCycle(elements) {
  // Create an array to track the visitation state of each element:
  // 0 = unvisited, 1 = visiting, 2 = visited
  const visited = new Array(elements.length).fill(0);
  const stack = []; // To keep track of the current path

  // Helper function to perform DFS
  function dfs(index) {
      // If the element is already being visited, we've found a cycle
      if (visited[index] === 1) {
          const cycleStartIndex = stack.indexOf(index);
          if (cycleStartIndex !== -1) {
              // Return the cycle in order
              return stack.slice(cycleStartIndex).concat(index);
          }
      }

      // If the element has been fully visited, no cycle was found here
      if (visited[index] === 2) return null;

      // Mark the element as visiting and add to the stack
      visited[index] = 1;
      stack.push(index);

      // Iterate over all the afterIndex entries
      let afterIndexes = [];
      if(elements[index] !== undefined) {
        afterIndexes = elements[index].conditions;
      }

      for (let i = 0; i < afterIndexes.length; i++) {
          const nextIndex = afterIndexes[i].ulterior;

          // Perform DFS on each index
          if (nextIndex !== null && nextIndex !== undefined) {
              const cycle = dfs(nextIndex);
              if (cycle) return cycle;
          }
      }

      // Mark the element as fully visited and remove from stack
      visited[index] = 2;
      stack.pop();

      // No cycle was found
      return null;
  }

  // Run DFS for each element
  for (let i = 0; i < elements.length; i++) {
      if (visited[i] === 0 && elements[i].conditions.length === 0) {
          const cycle = dfs(i);
          if (cycle) {
              return cycle; // Return the cycle with element ids
          }
      }
  }

  return null; // No cycles found
}

app.get('/health', (req, res) => {
  res.status(200).send('Server is healthy');
});

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});

// Endpoint to call JavaScript functions
app.post('/execute', async (req, res) => {
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
    await handleChange(body.changes);
  } else if (funcName === 'selectionChange') {
    selectedCells[sheetCodeName] = body.selection;
    handleSelectLinks();
  } else if (funcName === 'chgSheet') {
    sheetCodeName = body.sheetCodeName;
    headerColors = body.headerColors;
    values = values0[sheetCodeName].map(r => r.map(c => c===null?[]:c.toString().split(';').map(k => k.trim().toLowerCase()).filter(k => k)));
    await handleChange([]);
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



// Function to simplify boolean expressions using ABC
function simplifyExpression(expression, callback) {
  const abcCommand = `echo "read ${expression}; strash; mfs; print_stats; quit;" | ${abcPath}`;
  exec(abcCommand, (error, stdout, stderr) => {
      if (error) {
          console.error(`Error: ${error.message}`);
          callback(null, error.message);
          return;
      }
      if (stderr) {
          console.error(`Stderr: ${stderr}`);
          callback(null, stderr);
          return;
      }
      // Process the output from ABC and extract simplified logic
      const simplifiedExpr = stdout; // Process according to ABC's output format
      callback(simplifiedExpr, null);
  });
}


const jsep = require('jsep');

function extractLinkedTerms(formula) {
  // Parse the formula into an AST
  const ast = jsep(formula);

  const linkedTerms = new Set();

  function traverse(node) {
      if (node.type === 'BinaryExpression') {
          // Traverse both sides of the binary expression
          traverse(node.left);
          traverse(node.right);
      } else if (node.type === 'UnaryExpression') {
          // If it's a 'not' expression, skip it
          // We don't want to add negated terms
          if (node.operator !== 'not') {
              traverse(node.argument);
          }
      } else if (node.type === 'Identifier') {
          // Add the term to the set
          linkedTerms.add(node.name);
      }
  }

  // Start traversal from the root of the AST
  traverse(ast);

  return Array.from(linkedTerms);
}

const { PythonShell } = require('python-shell');


function exprToList(expr) {
  function parseExpression(expression) {
      // Remove any spaces from the expression
      expression = expression.replace(/\s+/g, '');

      // Replace logical operators with symbols that are easier to work with
      expression = expression.replace(/&&/g, '&').replace(/\|\|/g, '|').replace(/!/g, '~');

      // Function to process the expression recursively
      function process(expr) {
          if (expr.startsWith('~')) { // Handle NOT (~) operator
              return [0, 2, process(expr.substring(1))];
          } else if (expr.startsWith('(') && expr.endsWith(')')) {
              expr = expr.substring(1, expr.length - 1);
          }

          // Find the main operator, respecting parentheses
          let depth = 0;
          for (let i = 0; i < expr.length; i++) {
              if (expr[i] === '(') depth++;
              if (expr[i] === ')') depth--;
              if (depth === 0) {
                  if (expr[i] === '&') { // AND operator
                      return [-1, process(expr.substring(0, i)), process(expr.substring(i + 1))];
                  } else if (expr[i] === '|') { // OR operator
                      return [0, process(expr.substring(0, i)), process(expr.substring(i + 1))];
                  }
              }
          }

          return expr; // Base case: return the literal expression
      }

      return process(expression);
  }

  return parseExpression(expr);
}

const math = require('mathjs');

function evaluateExpression(expressionStr, scope) {
  // Replace variable names with their corresponding boolean values from the scope
  Object.keys(scope).forEach(varName => {
      const value = scope[varName];
      const regex = new RegExp(`\\b${varName}\\b`, 'g');
      expressionStr = expressionStr.replace(regex, value);
  });

  // Replace logical operators with JavaScript equivalents
  expressionStr = expressionStr
      .replace(/&&/g, '&&')
      .replace(/\|\|/g, '||')
      .replace(/!/g, '!')
      .replace(/true/g, 'true')
      .replace(/false/g, 'false');

  // Use Function constructor to evaluate the logical expression
  return new Function(`return ${expressionStr};`)();
}

function findFalseTerms(expressionStr) {
  // Extract variables from the expression (assuming variables are single letters)
  const variables = expressionStr.match(/[A-Za-z]\b/g);
  const uniqueVariables = [...new Set(variables)]; // Get unique variables

  let falseTerms = [];

  uniqueVariables.forEach(varName => {
      // Create a scope where this variable is false
      let scope = {};
      scope[varName] = false;

      // Evaluate the expression with the current variable set to false
      const result = evaluateExpression(expressionStr, scope);

      // If the result is false, add this variable to the falseTerms array
      if (!result) {
          falseTerms.push(varName);
      }
  });

  return falseTerms;
}


// Function to call Python script and return a promise
function runPythonScript(pythonFile, arg1) {
  return new Promise((resolve, reject) => {
    const options = {
      mode: 'json',
      pythonOptions: ['-u'], // Unbuffered stdout
      scriptPath: './', // Directory where the script is located
    };

    const pyshell = new PythonShell(pythonFile + ".py", options);

    // Send data to the Python script
    pyshell.send({ arg1 });

    // Receive output from the Python script
    pyshell.on('message', (message) => {
      resolve(message); // Resolve the promise with the message (result) from Python
    });

    pyshell.end((err) => {
      if (err) reject(err); // Reject the promise if there's an error
    });
  });
}

app.post('/test', async (req, res) => {

  // Example usage:
  const arg1 = "B | ~C & ~B";
  const arg2 = "!(A || !A) && (B || !!C && B)";

  
  try {

    // Await the result from the Python script
    const result = await runPythonScript("simplify_expression", arg1);

    // Send the result back to the client
    console.log("result");
    res.json(result);
  } catch (err) {
    res.status(500).send(err.message);
  }
  console.log("message");



});