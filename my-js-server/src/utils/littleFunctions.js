
const { addToDic, addToSupDic, getAddress, checkBrack, initializeData, getAttributes, getNames, setAllAttr, addFormula, checkWithC, check, searching, inconsist_removing_elem, inconsist_replacing_elem, midToWhite, displayReference, inconsist, findNewName, correct, dataGeneratorSub, renameSymbol, handleoldNameInputClick, getTrimmedResult, replaceWithSameWord, runPythonScript } = require('../utils/executeService');
const { getConditions } = require('../utils/getConditions');
const { handleChange } = require('../utils/handleChange');
const { handleSelectLinks } = require('../utils/handleSelectLinks');

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
  