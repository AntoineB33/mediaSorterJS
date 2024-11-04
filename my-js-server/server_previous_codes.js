

var msgTypeColors = [{ r: 255, g: 0, b: 0 }, { r: 255, g: 137, b: 0 }, { r: 0, g: 145, b: 255 }];
var updateRegularity = 1;

var values0 = {};
var valuesGlob = {};
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
var editRow = {};
var editCol = {};
var resolved = {};
var oldVersionsMaxNb = 20;
var columnTypes;
var child;
var data;
var allMediaRows;
var attrOfAttr = [];
var headerNames;
var nameSt = "names";
var nameInd;
var mediaSt = "media";
var mediaInd = -1;
var headerColors;
var attNames;
var actionsHistory = {};
var indActHist = 0;
let handleChanges = {};




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








  

function sugg(i, j) {
    suggSet(i, j, [values[i][j]]);
  }
  
  function suggSet(i, j, sugg) {
    let cellAddress = `${getColumnTag(j)}${i + 1}`;
    inconsist(i, j, "Error at " + cellAddress, sugg);
  }




  

async function ctrlZ(sheetCodeName) {
    if (indOldValues > 0) {
      indOldValues--;
      values = oldValues[indOldValues];
      correct(sheetCodeName);
      await check(sheetCodeName);
    }
  }
  
  async function ctrlY(sheetCodeName) {
    if (indOldValues < oldValues.length - 1) {
      indOldValues++;
      values = oldValues[indOldValues];
      correct(sheetCodeName);
      await check(sheetCodeName);
    }
  }



  

async function handleChangeCaller(updates, sheetCodeName) {
    if(sorting[sheetCodeName]) {
      addToSupDic(handleChanges, sheetCodeName, updates);
      setTimeout(() => {
        // This will run scheduledFunction after 10 seconds
        handleChangeCaller(updates, sheetCodeName);
      }, 1000);
    }
    for(let sheet in handleChanges) {
      handleChange(handleChanges[sheet], sheet);
    }
    for(let wb in handleChanges) {
      for(let sheet in handleChanges[wb]) {
        handleChange(handleChanges[wb][sheet], wb, sheet);
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