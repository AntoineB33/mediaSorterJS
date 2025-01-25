

const { addToDic, addToSupDic, getAddress, checkBrack, initializeData, getAttributes, getNames, setAllAttr, addFormula, checkWithC, check, searching, inconsist_removing_elem, inconsist_replacing_elem, midToWhite, displayReference, inconsist, findNewName, correct, dataGeneratorSub, renameSymbol, handleoldNameInputClick, getTrimmedResult, replaceWithSameWord, runPythonScript } = require('../utils/executeService');
const { getConditions } = require('../utils/getConditions');
const { handleChange } = require('../utils/handleChange');
const { getColumnTag, addListBoxList } = require('../utils/littleFunctions');

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

            params = {};
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
            
            displayReference(params, values, periods[perRef[row]][g]);
            values = params.values;
            i = params.i;
            j = params.j;
            k = params.k;
            val = params.val;
            columnTitle = params.columnTitle;
            message = params.message;
            suggs = params.suggs;
            msgTypeColors = params.msgTypeColors;
            values0Glob = params.values0Glob;
            valuesGlob = params.valuesGlob;
            prevLine = params.prevLine;
            lenAgg = params.lenAgg;
            attributes = params.attributes;
            sorting = params.sorting;
            sheetCodeName = params.sheetCodeName;
            editRow = params.editRow;
            editCol = params.editCol;
            resolved = params.resolved;
            oldVersionsMaxNb = params.oldVersionsMaxNb;
            columnTypes = params.columnTypes;
            data = params.data;
            allMediaRows = params.allMediaRows;
            attrOfAttr = params.attrOfAttr;
            nameSt = params.nameSt;
            nameInd = params.nameInd;
            mediaSt = params.mediaSt;
            mediaInd = params.mediaInd;
            headerColors = params.headerColors;
            attNames = params.attNames;
            actionsHistory = params.actionsHistory;
            indActHist = params.indActHist;
            handleChanges = params.handleChanges;
            rowIdMap = params.rowIdMap;
            response = params.response;
            colNumb = params.colNumb;
            
          }
        }
      }
    } else {
      for(let k = 0; k < values[row][column].length; k++) {
        if(column < 4) {
          for(let r = 1; r < values.length; r++) {
            for(let f = 0; f < values[r][0].length; f++) {

              params = {};
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
              
              let searchRes = searching(params, values, r, 0, f);
              values = params.values;
              i = params.i;
              j = params.j;
              k = params.k;
              val = params.val;
              columnTitle = params.columnTitle;
              message = params.message;
              suggs = params.suggs;
              msgTypeColors = params.msgTypeColors;
              values0Glob = params.values0Glob;
              valuesGlob = params.valuesGlob;
              prevLine = params.prevLine;
              lenAgg = params.lenAgg;
              attributes = params.attributes;
              sorting = params.sorting;
              sheetCodeName = params.sheetCodeName;
              editRow = params.editRow;
              editCol = params.editCol;
              resolved = params.resolved;
              oldVersionsMaxNb = params.oldVersionsMaxNb;
              columnTypes = params.columnTypes;
              data = params.data;
              allMediaRows = params.allMediaRows;
              attrOfAttr = params.attrOfAttr;
              nameSt = params.nameSt;
              nameInd = params.nameInd;
              mediaSt = params.mediaSt;
              mediaInd = params.mediaInd;
              headerColors = params.headerColors;
              attNames = params.attNames;
              actionsHistory = params.actionsHistory;
              indActHist = params.indActHist;
              handleChanges = params.handleChanges;
              rowIdMap = params.rowIdMap;
              response = params.response;
              colNumb = params.colNumb;
              
              if(row!=r && searchRes) {
                if(perRef[r] > -1) {
                  for(let g = 0; g < 3; g++) {
                    if(periods[perRef[r]][g]!=-1) {

                      params = {};
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
                      
                      displayReference(params, values, r);
                      values = params.values;
                      i = params.i;
                      j = params.j;
                      k = params.k;
                      val = params.val;
                      columnTitle = params.columnTitle;
                      message = params.message;
                      suggs = params.suggs;
                      msgTypeColors = params.msgTypeColors;
                      values0Glob = params.values0Glob;
                      valuesGlob = params.valuesGlob;
                      prevLine = params.prevLine;
                      lenAgg = params.lenAgg;
                      attributes = params.attributes;
                      sorting = params.sorting;
                      sheetCodeName = params.sheetCodeName;
                      editRow = params.editRow;
                      editCol = params.editCol;
                      resolved = params.resolved;
                      oldVersionsMaxNb = params.oldVersionsMaxNb;
                      columnTypes = params.columnTypes;
                      data = params.data;
                      allMediaRows = params.allMediaRows;
                      attrOfAttr = params.attrOfAttr;
                      nameSt = params.nameSt;
                      nameInd = params.nameInd;
                      mediaSt = params.mediaSt;
                      mediaInd = params.mediaInd;
                      headerColors = params.headerColors;
                      attNames = params.attNames;
                      actionsHistory = params.actionsHistory;
                      indActHist = params.indActHist;
                      handleChanges = params.handleChanges;
                      rowIdMap = params.rowIdMap;
                      response = params.response;
                      colNumb = params.colNumb;
                      
                    }
                  }
                } else {

                  params = {};
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
                  
                  displayReference(params, values, r);
                  values = params.values;
                  i = params.i;
                  j = params.j;
                  k = params.k;
                  val = params.val;
                  columnTitle = params.columnTitle;
                  message = params.message;
                  suggs = params.suggs;
                  msgTypeColors = params.msgTypeColors;
                  values0Glob = params.values0Glob;
                  valuesGlob = params.valuesGlob;
                  prevLine = params.prevLine;
                  lenAgg = params.lenAgg;
                  attributes = params.attributes;
                  sorting = params.sorting;
                  sheetCodeName = params.sheetCodeName;
                  editRow = params.editRow;
                  editCol = params.editCol;
                  resolved = params.resolved;
                  oldVersionsMaxNb = params.oldVersionsMaxNb;
                  columnTypes = params.columnTypes;
                  data = params.data;
                  allMediaRows = params.allMediaRows;
                  attrOfAttr = params.attrOfAttr;
                  nameSt = params.nameSt;
                  nameInd = params.nameInd;
                  mediaSt = params.mediaSt;
                  mediaInd = params.mediaInd;
                  headerColors = params.headerColors;
                  attNames = params.attNames;
                  actionsHistory = params.actionsHistory;
                  indActHist = params.indActHist;
                  handleChanges = params.handleChanges;
                  rowIdMap = params.rowIdMap;
                  response = params.response;
                  colNumb = params.colNumb;
                  
                }
              }
            }
          }
        } else if (column < colNumb - 1) {
          nbMediaInCat = attributes[data[row][2]].length;
          for(let r = 1; r < values.length; r++) {
            for(let d = 4; d < colNumb - 1; d++) {
              for(let f = 0; f < values[r][d].length; f++) {

                params = {};
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
                
                let searchRes = searching(params, values, r, d, f);
                values = params.values;
                i = params.i;
                j = params.j;
                k = params.k;
                val = params.val;
                columnTitle = params.columnTitle;
                message = params.message;
                suggs = params.suggs;
                msgTypeColors = params.msgTypeColors;
                values0Glob = params.values0Glob;
                valuesGlob = params.valuesGlob;
                prevLine = params.prevLine;
                lenAgg = params.lenAgg;
                attributes = params.attributes;
                sorting = params.sorting;
                sheetCodeName = params.sheetCodeName;
                editRow = params.editRow;
                editCol = params.editCol;
                resolved = params.resolved;
                oldVersionsMaxNb = params.oldVersionsMaxNb;
                columnTypes = params.columnTypes;
                data = params.data;
                allMediaRows = params.allMediaRows;
                attrOfAttr = params.attrOfAttr;
                nameSt = params.nameSt;
                nameInd = params.nameInd;
                mediaSt = params.mediaSt;
                mediaInd = params.mediaInd;
                headerColors = params.headerColors;
                attNames = params.attNames;
                actionsHistory = params.actionsHistory;
                indActHist = params.indActHist;
                handleChanges = params.handleChanges;
                rowIdMap = params.rowIdMap;
                response = params.response;
                colNumb = params.colNumb;
                
                if(row!=r && searchRes) {
                  displayReference(params, values, r);
                  values = params.values;
                  i = params.i;
                  j = params.j;
                  k = params.k;
                  val = params.val;
                  columnTitle = params.columnTitle;
                  message = params.message;
                  suggs = params.suggs;
                  msgTypeColors = params.msgTypeColors;
                  values0Glob = params.values0Glob;
                  valuesGlob = params.valuesGlob;
                  prevLine = params.prevLine;
                  lenAgg = params.lenAgg;
                  attributes = params.attributes;
                  sorting = params.sorting;
                  sheetCodeName = params.sheetCodeName;
                  editRow = params.editRow;
                  editCol = params.editCol;
                  resolved = params.resolved;
                  oldVersionsMaxNb = params.oldVersionsMaxNb;
                  columnTypes = params.columnTypes;
                  data = params.data;
                  allMediaRows = params.allMediaRows;
                  attrOfAttr = params.attrOfAttr;
                  nameSt = params.nameSt;
                  nameInd = params.nameInd;
                  mediaSt = params.mediaSt;
                  mediaInd = params.mediaInd;
                  headerColors = params.headerColors;
                  attNames = params.attNames;
                  actionsHistory = params.actionsHistory;
                  indActHist = params.indActHist;
                  handleChanges = params.handleChanges;
                  rowIdMap = params.rowIdMap;
                  response = params.response;
                  colNumb = params.colNumb;
                  
                }
              }
            }
          }
        }
      }
    }
  }
  