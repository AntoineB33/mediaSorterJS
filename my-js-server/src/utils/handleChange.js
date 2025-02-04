
const { addToDic, addToSupDic, getAddress, checkBrack, initializeData, getAttributes, getNames, setAllAttr, addFormula, checkWithC, check, searching, inconsist_removing_elem, inconsist_replacing_elem, midToWhite, displayReference, inconsist, findNewName, correct, dataGeneratorSub, renameSymbol, handleoldNameInputClick, getTrimmedResult, replaceWithSameWord, runPythonScript } = require('../utils/executeService');
const { getConditions } = require('../utils/getConditions');
const { handleSelectLinks } = require('../utils/handleSelectLinks');
const { getColumnTag, addListBoxList } = require('../utils/littleFunctions');

async function handleChange(params, updates) {
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

        params = {};
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
        
        inconsist_removing_elem(values, 0, j, k, `Column ${getColumnTag(j)} labelled as "${alreadyInd[0]}" already at column ${getColumnTag(alreadyInd[1])}.`);

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
        
        return;
      }
    }
  }
  if(nameInd == -1) {

    params = {};
    params.values = values;
    params.i = 0;
    params.j = values[0].length;
    params.k = k;
    params.val = val;
    params.columnTitle = columnTitle;
    params.message = `Missing column "${nameSt}".`;
    params.suggs = [[nameSt]];
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
    
    inconsist(params);
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
    
    return;
  }
  if(headerColors[nameInd] === 16777215) {

    params = {};
    params.values = values;
    params.i = 0;
    params.j = nameInd;
    params.k = k;
    params.val = val;
    params.columnTitle = columnTitle;
    params.message = `Column ${getColumnTag(nameInd)} labelled as "${nameSt}" must have a background color.`;
    params.suggs = [['']];
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
    
    inconsist(params);
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
    
    return;
  }
  if(nameInd == mediaInd) {

    params = {};
    params.values = values;
    params.i = 0;
    params.j = j;
    params.k = k;
    params.val = val;
    params.columnTitle = columnTitle;
    params.message = `Column ${getColumnTag(nameInd)} labelled as "${nameSt}" and "${mediaSt}".`;
    params.suggs = [['']];
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
    
    inconsist(params);
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
    
    return;
  }
  if(mediaInd != -1 && namesColor !== headerColors[mediaInd]) {

    params = {};
    params.values = values;
    params.i = 0;
    params.j = mediaInd;
    params.k = k;
    params.val = val;
    params.columnTitle = columnTitle;
    params.message = `Column ${getColumnTag(mediaInd)} must have the same background color as the attributes.`;
    params.suggs = [headerColors[nameInd]];
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
    
    inconsist(params, Actions.NewBgCol);
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

      params = {};
      params.values = values;
      params.i = 0;
      params.j = j;
      params.k = k;
      params.val = val;
      params.columnTitle = columnTitle;
      params.message = `Third background color at column ${j + 1} (must be only two, for the conditions and the attributes)`;
      params.suggs = [condColor];
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
      
      inconsist(params, true);
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
      
      return;
    }
  }

  params = {};
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
  
  await check(params, sheetCodeName);
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
  
  // oldValues.splice(indOldValues, oldValues.length - 1 - indOldValues);
  // oldValues.push(values);
  // if(oldVersionsMaxNb < oldValues.length) {
  //   oldValues.shift();
  // }
}
