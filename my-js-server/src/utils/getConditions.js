
const { addToDic, addToSupDic, getAddress, checkBrack, initializeData, getAttributes, getNames, setAllAttr, addFormula, checkWithC, check, searching, inconsist_removing_elem, inconsist_replacing_elem, midToWhite, displayReference, inconsist, findNewName, correct, dataGeneratorSub, renameSymbol, handleoldNameInputClick, getTrimmedResult, replaceWithSameWord, runPythonScript } = require('../utils/executeService');
const { handleChange } = require('../utils/handleChange');
const { handleSelectLinks } = require('../utils/handleSelectLinks');
const { getColumnTag, addListBoxList } = require('../utils/littleFunctions');

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
      
      let formula = [];
      params.values = values;
      params.i = i;
      params.formula = formula;
      addFormula(params);
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
      
      for (let a of data[i].attributes) {

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
        
        params.i = a;
        addFormula(params);
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
      formula = formula.join(" && ");
      let replacedWithWords;
      let reverseReplacements;
      try{
        ({replacedWithWords, reverseReplacements} = replaceWithSameWord(formula));
      } catch(error) {

        params = {};
        params.values = values;
        params.i = i;
        params.j = 0;
        params.k = k;
        params.val = val;
        params.columnTitle = columnTitle;
        params.message = error.message;
        params.suggs = [];
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
      
      let direct_terms = [];
      if(replacedWithWords!=='') {
        try {
  
          direct_terms = await runPythonScript("test_direct", replacedWithWords)
        } catch (err) {

          params = {};
          params.values = values;
          params.i = i;
          params.j = 0;
          params.k = k;
          params.val = val;
          params.columnTitle = columnTitle;
          params.message = `incorrect condition : ${err.message}.`;
          params.suggs = [];
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
            
            params.values = values;
            params.i = i;
            params.j = 0;
            params.k = 0;
            params.val = groups.attrib_ref;
            params.columnTitle = groups.attrib_ref_col;
            let refInd_val = checkBrack(params);
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
            
            if(refInd_val == -1) {
              return -1;
            }
            if(refInd_val == -2) {

              params = {};
              params.values = values;
              params.i = i;
              params.j = 0;
              params.k = k;
              params.val = val;
              params.columnTitle = columnTitle;
              params.message = `attribute "${val}" at row ${i + 1} not found.`;
              params.suggs = [];
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
              
              return -1;
            }
            groups.attrib_ref = refInd_val;
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
          
          params.values = values;
          params.i = i;
          params.j = 0;
          params.k = 0;
          params.val = groups.precedent;
          params.columnTitle = groups.precedent_col;
          let refInd_val = checkBrack(params);
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

              params = {};
              params.values = values;
              params.i = i;
              params.j = 0;
              params.k = k;
              params.val = val;
              params.columnTitle = columnTitle;
              params.message = `precedent "${groups.precedent}" at row ${i + 1} not found.`;
              params.suggs = [];
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

              params = {};
              params.values = values;
              params.i = i;
              params.j = 0;
              params.k = k;
              params.val = val;
              params.columnTitle = columnTitle;
              params.message = `position "${groups.position}" at ${getAddress(i, j)} is not unique.`;
              params.suggs = [];
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

      params = {};
      params.values = values;
      params.i = incoherence.row;
      params.j = 0;
      params.k = k;
      params.val = val;
      params.columnTitle = columnTitle;
      params.message = incoherence.result;
      params.suggs = [];
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

          params = {};
          params.values = values;
          params.i = i;
          params.j = 0;
          params.k = k;
          params.val = val;
          params.columnTitle = columnTitle;
          params.message = `incorrect condition : ${err.message}.`;
          params.suggs = [];
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
          
          return -1;
        }
  
        let formulaList = transformLogicalFormula(getTrimmedResult(simplified));
  
        // change the 
        data[i].conditions = transformTerms(formulaList, exampleTransform, reverseReplacements, newInds);
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

    checkWithC(params, sheetCodeName);
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
    
    response.push({ "listBoxList": [] });
    resolved[sheetCodeName] = true;
  }
  