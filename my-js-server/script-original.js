/* global console, document, Excel, Office */

// there are "Excel API limitations" comments

// npx webpack serve --mode development   - to run the server
// npm start                              - to create and open an Excel file with the add-in

// 5 : receive "computing"
// 6 : write "=not_updating_mode()" or "=updating_mode()"
// 7 : receive "finished" (sorting finished)
// 8 : write "=GetColumnText()" (start the sorting)
// 9 : write "=insertInPowerPoint()" (create a PowerPoint with the sorted media)
// 10 : write "=quit_sort()" (to stop the sorting)
// 11 : receive "updated" (now analyse the updated version)
// 12 : receive "sorting" (not analysing because not completly updated)

// 0: color when changed, 1: all the colors, 2: timer
var updateRegularity = 1;
var rootXMLJSName = "sorterRootJS";
var rootXMLVBAName = "sorterRootVBA";
var rootXMLMsgName = "sorterRootMsg";
var values;
var oldValues = [];
var indOldValues = 0;
var perRef = [];
var periods = [];
var colNumb;
var perNot;
var stop = false;
var data = [];
var val2;
var val;
var elem2;
var elem;
var stat;
var stat2;
var perLen = 0;
var perInt = [];
var perIntCont = [];
var prevLine = 0;
var nbLineBef;
var fileId = "1EoCDqXsL0tqAW6M7qVQT_N7L_NsBirMJ";
var values0;
var lenAgg;
var attributes;
var dataAgg;
var progChg = 0;
var sorting = false;
var colors;
var fonts;
var place;
var steps;
var visited;
var notInitialized = true;
var updating = false;
var handleRunning = false;
var checkRunning = false;
var colNb;
var sheetNames = {};
let customXMLPartDict = {};
let initCustomXMLPartDict = {};
var prevColNumb = 0;
var resolved;


/**
 * Converts a given integer to its corresponding Excel column tag.
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

// Office.onReady((info) => {
//   if (info.host === Office.HostType.Excel) {
//     // Initialization check
//     if (window.hasInitialized) {
//       console.log("Already initialized.");
//       return;
//     }
//     window.hasInitialized = true; // Set initialization flag
//     Excel.run(async (context) => {
//     }).catch(error => console.error("Error:", error));
//   }
// });




// class Mutex {
//   constructor() {
//     this.locked = false;
//     this.queue = [];
//   }

//   async lock() {
//     while (this.locked) {
//       await new Promise(resolve => this.queue.push(resolve));
//     }
//     this.locked = true;
//   }

//   unlock() {
//     console.log(`len : ${this.queue.length}`)
//     if (this.queue.length > 0) {
//       const next = this.queue.shift();
//       next();
//     } else {
//       this.locked = false;
//     }
//   }

//   async runWithLock(func) {
//     await this.lock();
//     try {
//       await func();
//     } finally {
//       console.log("to unlock")
//       this.unlock();
//     }
//   }
// }

// const mutex = new Mutex();
// const mutexPush = new Mutex();



class Mutex {
  constructor() {
      this.queue = [];
      this.locked = false;
  }

  lock() {
      const unlock = () => {
          const nextResolve = this.queue.shift();
          if (nextResolve) {
              nextResolve(unlock);
          } else {
              this.locked = false;
          }
      };

      if (this.locked) {
          return new Promise(resolve => this.queue.push(resolve)).then(() => unlock);
      } else {
          this.locked = true;
          return Promise.resolve(unlock);
      }
  }

  unlock(unlock) {
      if (typeof unlock === 'function') {
          unlock();
      }
  }
}



const mutex = new Mutex();
const mutexPush = new Mutex();

// async function doWork(v) {
//     try {
//         console.log("hey " + v);
//         await delay(1000000);
//     } finally {
//         mutex.unlock(unlock);
//     }
// }

class AsyncFunctionQueue {
  constructor() {
    this.queue = [];
    this.processing = false;
  }

  async enqueue(func, isHandle = 0, ...args) {
    const unlock = await mutexPush.lock();
    console.log(`enqueue : ${isHandle} ${this.queue.length} ${this.queue.length?this.queue[this.queue.length - 1][1]:"None"} `)
    if(isHandle && this.queue.length && this.queue[this.queue.length - 1][1]) {
      this.queue[this.queue.length - 1][1]++;
      mutexPush.unlock(unlock);
      return;
    }
    // Push the function and its resolver into the queue
    this.queue.push([async () => {
      await Excel.run(async (context) => {
        if(args) {
          console.log(`args[0] : ${args[0]} ${args.length}`)
        }
        await func(context, ...args);
        await context.sync();
      });
    }, isHandle]);
    console.log(`function name : ${func.name}`)
    mutexPush.unlock(unlock);
    this.processQueue();
    // return new Promise((resolve, reject) => {
    // });
  }

  async processQueue() {
    console.log(`mutex`)
    const unlock = await mutex.lock();
    // await mutex.runWithLock(async () => {
    if(this.processing) {
      console.log(`retun`)
      return;
    }
    this.processing = true;
    console.log("processing is true")
    while (this.queue.length > 0) {
      console.log(`this.queue.length : ${this.queue.length}`)
      const unlock = await mutexPush.lock();
      const lastEvent = this.queue.shift();
      mutexPush.unlock(unlock);
      console.log(`lastEvent : ${lastEvent[1]} ${progChg} `)
      if(lastEvent[1] && progChg) {
        progChg-=lastEvent[1];
        if(progChg<0) {
          continue;
        }
        progChg = 0;
      }
      console.log(`lastEvent : ${lastEvent[1]} ${progChg} `)
      while (1) {
        try {
          console.log(`try`)
          await lastEvent[0]();
          break;
        } catch (error) {
          console.error("Error during function execution", error);
          await delay(100000000);
        }
      }
    }
    this.processing = false; // Set processing to false only when queue is empty

    // In case items were added while processing was being set to false
    console.log("processing is false")
    if (this.queue.length > 0) {
      this.processQueue();
    }
    
    mutex.unlock(unlock);
    console.log(`after mutex`)
    
    // mutex.lock(`mutex`)
    // console.log("hey there")
    // mutex.unlock();
  }
}
const myQueue = new AsyncFunctionQueue();

document.getElementById("updateButton").style.display = 'none';
document.getElementById("linkToCell").style.display = 'none';
document.getElementById("sortButton").addEventListener("click", () =>
  myQueue.enqueue(dataGenerator)
);
document.getElementById("updateButton").addEventListener("click", () =>
  myQueue.enqueue(updateFunc)
);
document.getElementById("ctrlZ").addEventListener("click", () =>
  myQueue.enqueue(ctrlZ)
);
document.getElementById("ctrlY").addEventListener("click", () =>
  myQueue.enqueue(ctrlY)
);
document.getElementById("oldNameInput").addEventListener("click", () => myQueue.enqueue(handleoldNameInputClick));
document.getElementById("newNameInput").addEventListener("click", () => myQueue.enqueue(handlenewNameInputClick));
document.getElementById("getPowerPoint").addEventListener("click", () => myQueue.enqueue(callVBA, 0, "insertInPowerPoint"));
$(document).ready(() => myQueue.enqueue(init));
// $(document).ready(() => 
// {
//   Excel.run(function(context) {
//     let sheet = context.workbook.worksheets.getActiveWorksheet();
//     let rangeAddress = `A1:B2`; // Adjust range as needed
//     let range = sheet.getRange(rangeAddress);
    
//     range.load(['rowCount', 'columnCount']);
//     range.format.fill.load('color'); // Load the background color of each cell in the range

//     return context.sync().then(function() {
//         let rows = range.rowCount;
//         let cols = range.columnCount;
//         let backgroundColors = [];

//         for (let i = 0; i < rows; i++) {
//             let rowColors = [];
//             for (let j = 0; j < cols; j++) {
//                 rowColors.push(range.getCell(i, j).format.fill.color);
//             }
//             backgroundColors.push(rowColors);
//         }

//         console.log(backgroundColors); // Output the 2D list of background colors
//         // You can also return this list or process it further as needed
//     });
// }).catch(function(error) {
//     console.log(error);
// });
// });


document.getElementById('newNameInput').addEventListener('keydown', function (event) {
  if (event.key === "Enter") {
    // Prevent default to avoid any other actions associated with pressing Enter
    event.preventDefault();
    myQueue.enqueue(renameSymbol, 0, "ImportVBAComponents", document.getElementById("oldNameInput").value, event.target.value);
  }
});

document.getElementById('newWorkbookName').addEventListener('keydown', function (event) {
  if (event.key === "Enter") {
    // Prevent default to avoid any other actions associated with pressing Enter
    event.preventDefault();
    myQueue.enqueue(callVBA, 0, "ImportVBAComponents", event.target.value);
  }
});

async function renamingSheet(context) {
  console.log("RENAMING")
}

function addItem(row) {
  const listElem = document.createElement("li");
  listElem.classList.add("ms-ListItem", "suggestion-item");
  link = document.createElement("a");
  link.setAttribute("href", "#");
  link.setAttribute("data-text", `${row}`);
  link.textContent = values[row][0][0]
  listElem.appendChild(link);
  return listElem;
}

/**
 *  Update the links to show at the menu each time a cell is selected
 * 
 */
async function handleSelectLinks(context, address) {
  // only if no errors in the sheet
  if(resolved) {
    const match = address.match(/([A-Z]+)(\d+)/);
    const columnLetters = match[1];
    const row = +match[2] - 1;
    
    // Convert column letters to a number
    let column = 0;
    for (let i = 0; i < columnLetters.length; i++) {
      column = column * 26 + (columnLetters.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    column -= 1
    
    const suggestionList = document.getElementById("relatives-list");
    suggestionList.innerHTML = ""; // Clear previous suggestions
    if(row>=nbLineBef || column >= colNumb) {
      return -1;
    }
    

    const listItem = document.createElement("ul");
    if(column==0) {
      if(perRef[row] > -1) {
        for(let g = 0; g < 3; g++) {
          if(periods[perRef[row]][g] < nbLineBef && periods[perRef[row]][g]!=row) {
            listItem.appendChild(addItem(periods[perRef[row]][g]));
          }
        }
      }
    } else {
      for(let k = 0; k < values[row][column].length; k++) {
        const listItemSub = document.createElement("li");
        const listItemSubUl = document.createElement("ul");
        if(isTip(context, row, column, k) == -1) {
          return -1;
        }
        if(column < 4) {
          for(let r = 1; r < values.length; r++) {
            for(let f = 0; f < values[r][0].length; f++) {
              if(row!=r && searching(r, 0, f)) {
                if(perRef[r] > -1) {
                  for(let g = 0; g < 3; g++) {
                    if(periods[perRef[r]][g]!=-1) {
                      listItemSubUl.appendChild(addItem(r));
                    }
                  }
                } else {
                  listItemSubUl.appendChild(addItem(r));
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
                  listItemSubUl.appendChild(addItem(r));
                }
              }
            }
          }
        }
        listItemSub.appendChild(listItemSubUl);
        listItem.appendChild(listItemSub);
      }
    }
    suggestionList.appendChild(listItem);
    const suggestionItems = document.querySelectorAll(".suggestion-item");
    suggestionItems.forEach((item) => {
      item.addEventListener("click", (event) => myQueue.enqueue(
        async (context, text) => {
          if (text) {
            let sheet = context.workbook.worksheets.getActiveWorksheet();
            // let cell = sheet.getRange(`${getColumnTag(j)}${i + 1}`);
            let cell = sheet.getCell(+text, 0);
            cell.select();
          }
        }, 0, event.target.getAttribute("data-text")
      ));
    });
  }
}

async function init(context) {
  
  
  // const workbook2 = context.workbook;
  // const customXmlParts2 = workbook2.customXmlParts;

  // // Load the id of each custom XML part
  // customXmlParts2.load("items/id");

  // await context.sync();

  // // Loop through each custom XML part and delete it
  // customXmlParts2.items.forEach(part => {
  //     part.delete();
  // });

  // // Sync the context to apply the deletions
  // await context.sync();



  initCustomXMLPartDict[rootXMLJSName] = "false";
  initCustomXMLPartDict[rootXMLVBAName] = "true";
  initCustomXMLPartDict[rootXMLMsgName] = "";
  const sheets = context.workbook.worksheets;
  sheets.load('items/name');
  await context.sync();
  sheets.items.forEach(sheet => {
      sheetNames[sheet.id] = sheet.name;
  });
  // Get the worksheet collection
  const onChangedSheets = sheets.onChanged;
  // Add an event handler for the onActivated event
  onChangedSheets.add((args) => {
    Excel.run(async function (context) {
      let changedSheet = sheets.getItem(args.worksheetId);
      changedSheet.load('name');
      await context.sync();
      if (changedSheet.name !== sheetNames[changedSheet.id]) {
        changedSheet.name = sheetNames[changedSheet.id];
        myQueue.enqueue(renamingSheet);
        return 1;
      }
    });
    myQueue.enqueue(async (context) => await handleChange(context, args.address),1);
  });
  console.log("Initialization complete 0");
  colors = [];
  fonts = [];
  await handleChange(context);
  sheets.onSelectionChanged.add((event) => myQueue.enqueue(handleSelectLinks, 0, event.address));
}

function delay(ms = 200) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function accessXML(context) {
  var customXmlParts = context.workbook.customXmlParts;
  customXmlParts.load('items');
  await context.sync();

  for (let key in initCustomXMLPartDict) {
    if (initCustomXMLPartDict.hasOwnProperty(key)) {
      customXMLPartDict[key] = 0;
    }
  }
  let customXmlPartsHasChg = 1;
  var xmlDocMsg, rootMsg;
  while(customXmlPartsHasChg) {
    customXmlPartsHasChg = 0;
    for (let i = 0; i < customXmlParts.items.length; i++) {
      const xmlPart = customXmlParts.items[i];
      console.log("heyhey")
      const xml = xmlPart.getXml();
      await context.sync();

      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(xml.value, "application/xml");
      const root = xmlDoc.documentElement;
      const rootNodeName = root.nodeName;

      
      for (let key in customXMLPartDict) {
        if (customXMLPartDict.hasOwnProperty(key)) {
          if(rootNodeName === key) {
            console.log(`key : ${key}`)
            customXMLPartDict[key] = xmlPart;
            if(key == rootXMLMsgName) {
              rootMsg = root;
              xmlDocMsg = xmlDoc;
              console.log(`root : ${typeof rootMsg === 'undefined'}`)
              console.log(`youuuu`)
            }
          }
        }
      }
    }
    for (let key in customXMLPartDict) {
      if (customXMLPartDict.hasOwnProperty(key) && customXMLPartDict[key] == 0) {
        customXmlPartsHasChg = 1;
        customXmlParts.add("<" + key + ">" + initCustomXMLPartDict[key] + "</" + key + ">");
      }
    }
    console.log("while")
    // await delay(1000)
    if(customXmlPartsHasChg) {
      customXmlParts.load('items');
      await context.sync();
    }
  }
  customXMLPartDict["sorterRootJS"].setXml("<sorterRootJS>false</sorterRootJS>");
  await context.sync();
  let xmlBlob = await customXMLPartDict[rootXMLMsgName].getXml();
  await context.sync();
  console.log(xmlBlob.value)
  console.log("end access")
  console.log(`root : ${typeof rootMsg === 'undefined'}`)
  return [xmlDocMsg, rootMsg];
}

async function releaseLockXML(context) {
  console.log("???")
  customXMLPartDict["sorterRootJS"].setXml("<sorterRootJS>true</sorterRootJS>");
  await context.sync();
  console.log("???")
}

async function handleChange(context, address = "") {
  resolved = false;
  console.log(`address : ${address}`)
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.load("name");
  let range = sheet.getUsedRange();
  range.load("values");
  await context.sync();
  values = range.values;
  nbLineBef = values.length;
  colNb = values[0].length;
  // console.log(`A1:${String.fromCharCode(65 + Math.max(stat2, 7))}${Math.max(nbLineBef, prevLine)}`);
  // range = sheet.getRange(`A1:${String.fromCharCode(65 + Math.max(stat2, 7))}${Math.max(nbLineBef, prevLine)}`);

  // range.load({
  //   select: "format/fill/color",  // Specifies exactly what to load
  //   expand: "rows/items"  // This attempts to expand rows and cells; adjust based on actual needs and API version
  // });

  // await context.sync();
  // console.log("sync")

  // // Access the colors
  // colors = [];
  // for (let row = 0; row < Math.max(nbLineBef, prevLine); row++) {
  //     let rowColors = [];
  //     for (let col = 0; col < Math.max(stat2, 7); col++) {
  //         let cellColor = range.getCell(row, col).format.fill.color;
  //         rowColors.push(cellColor);
  //         console.log(`color : ${cellColor}`)
  //     }
  //     colors.push(rowColors);
  // }
  // return -1;






  // range = sheet.getRange(`A1:B2`);
  // // range.format.fill.load("color");
  // // range.format.fill.load("font");
  // range.load("format/fill/color");
  // // range.load("format/fill/font");
  // await context.sync();
  // colors = range.format.fill.color;
  // console.log(`fonts : ${colors.length}`);
  // for(let i = 0; i < colors.length; i++) {
  //   console.log(colors[i]);
  // }
  // return -1;
  // fonts = range.format.fill.font;
  // console.log(`fonts : ${fonts}`);
  if (nbLineBef == 0 || colNb < 7) {
    range = sheet.getRange(`A1:${getColumnTag(Math.max(colNb, 7))}${Math.max(1, nbLineBef)}`);
    range.load("values");
    await context.sync();
    values = range.values;
    nbLineBef = values.length;
    colNb = values[0].length;
  }
  // return;
  colNumb = colNb;
  for (let i = 6; i < colNb; i++) {
    if (values[0][i] == "url") {
      colNumb = i + 1;
      break;
    }
  }

  // erase or colorize when url is displaced
  if (colNumb != prevColNumb) {
    let jL = Math.min(colNumb, prevColNumb);
    let jU = Math.max(colNumb, prevColNumb);
    if (colNumb < prevColNumb) {
      let rangeCol = sheet.getRange(`${getColumnTag(jL)}1:${getColumnTag(jU - 1)}${prevLine + 1}`);
      rangeCol.format.fill.clear();
    } else {
      for(let i = 0; i<prevLine; i++) {
        let rangeCol = sheet.getRange(`${getColumnTag(jL)}${i + 1}:${getColumnTag(jU - 1)}${i + 1}`);
        if(colors[i]) {
          rangeCol.format.fill.color = colors[i];
        } else {
          rangeCol.format.fill.clear();
        }
      }
    }
    for(let i = 0; i<prevLine; i++) {
      styleBorders(sheet, i, jL, jU, colNumb > prevColNumb);
    }
    prevColNumb = colNumb;
  }
  // var workbook = context.workbook;
  // var customXmlParts = workbook.customXmlParts;
  // customXmlParts.load("items");
  // await context.sync().then(function() {
  //     var promises = customXmlParts.items.map(function(part) {
  //         part.delete();
  //     });
  //     return context.sync(Promise.all(promises));
  // });




  const [xmlDoc, root] = await accessXML(context);

  // Initialize a list to hold the "msg" elements.
  let vbaMessages = [];
  console.log(`root : ${typeof root === 'undefined'}`)

  // First loop: Add all "msg" XML children to the list.
  const msgElements = root.getElementsByTagName("vbaMessage");
  console.log("???")
  for (let i = 0; i < msgElements.length; i++) {
    vbaMessages.push(msgElements[i]);
  }

  // Second loop: Delete "msg" elements from the root element.
  for (let i = vbaMessages.length - 1; i >= 0; i--) {
      root.removeChild(vbaMessages[i]);
  }

  // Serialize the modified XML document back to a string.
  const serializer = new XMLSerializer();
  const newXml = serializer.serializeToString(xmlDoc);
  
  // Set the new XML back to the custom XML part.
  customXMLPartDict[rootXMLMsgName].setXml(newXml);

  await context.sync();
  

  await releaseLockXML(context);
  // Log each line's text content
  for (let i = 0; i < vbaMessages.length; i++) {
    const message = vbaMessages[i];
    const functionName = message.getElementsByTagName("functionName")[0].textContent;
    const sheetId = message.getElementsByTagName("sheetId")[0].textContent;
    switch (functionName) {
      case "finished":
        if(sheetId == sheet.name) {
          let sortButton = document.getElementById("sortButton");
          sortButton.textContent = "Sort";
        }
        sorting = false;
        break;
      case "sorting":
        const vbaValue = message.getElementsByTagName("arg")[0].textContent;
        progChg+=+vbaValue * colNumb - 1;
        if (!sorting) {
          // that means the sorting started by double clicking sortFromCopy.exe
          await dataGenerator(context);
        }
        return -1;
    }
  }

  let headLine = ["names", "following", "after", "before", ""];
  for (let i = 0; i < 4; i++) {
    if (headLine[i] != values[0][i]) {
      let cell = sheet.getRange(`${getColumnTag(i)}1`);
      progChg++;
      cell.values = [[`${headLine[i]}`]];
    }
  }
  // sheet.getRange("A2").values = [[`${colNumb}`]];
  // sheet.getRange("A3").values = [[`${prevLine}`]];
  // prevLine++;
  if ("periods" != values[0][colNumb - 2]) {
    progChg++;
    range.getCell(0, colNumb - 2).values = [["periods"]];
  }
  if ("url" != values[0][colNumb - 1]) {
    progChg++;é
    range.getCell(0, colNumb - 1).values = [["url"]];
  }

  // Each cell becomes a list
  values0 = JSON.parse(JSON.stringify(values));
  for (let i = 0; i < nbLineBef; i++) {
    for (let j = 0; j < colNumb; j++) {
      values[i][j] = values[i][j]
        .toString()
        .split(";")
        .map(function (v, i) {
          // console.log("return2");
          return v.trim();
        })
        .filter(function (e) {
          // console.log("return3");
          return e !== "";
        });
    }
    let surp = values[i].length - colNumb;
    // console.log("surp")
    // console.log(`splice : ${colors}`);
    // colors[i].splice(colNumb, surp);
    // console.log("splice")
    // fonts[i].splice(colNumb, surp);
    values[i].splice(colNumb, surp);
    // if(values[0].length>=colNumb+2 && values[i][colNumb+2]) {
    //   let int = parseInt(values[i][colNumb+2]);
    //   if(!isNaN(int)) {
    //   }
    // }
  }
  var nbLineBefBef = nbLineBef;
  for (let i = nbLineBef - 1; i > 0; i--) {
    if (values[i].some((v, j) => v.length != 0 && j < colNumb)) {
      break;
    }
    nbLineBef--;
  }
  values.splice(nbLineBef, nbLineBefBef - nbLineBef);
  oldValues.splice(indOldValues, oldValues.length - 1 - indOldValues);
  oldValues.push(values);
  console.log("!!!")
  await check(context);
  console.log("...")
}

async function check(context) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.load("name");
  perNot = ["", "start:", "end:"];
  periods = [];
  perRef = [0];
  data = [0];
  perInt = [0];

  perIntCont = [0];
  var r;

  for (let i = 1; i < nbLineBef; i++) {
    perInt.push([]);
    perIntCont.push([[], [], []]);
    perRef.push(-1);
    if (values[i][colNumb - 1].length == 0) {
      // if there is no url
      perRef[i] = -2;
    }

    // data[i][0] : if it doesn't follow any row
    // data[i][1] : the rows it come after
    // data[i][2] : the attributes
    // data[i][3] : the row that follows it, -1 if no one
    // data[i][4] : the tuples [the row that follows it in respect of an attribute; the attribute]
    data.push([true, [], [], -1, [], []]);
    data.push({"follows":[], "comesAfter":[], "attributes":[], });
    for (let k = 0; k < values[i][0].length; k++) {
      if(isTip(context, i, 0, k)) {
        return -1;
      }
      stop = false;
      //check if already before
      for (let r0 = 1; r0 <= i; r0++) {
        if (r0 == 1) {
          r = i;
        } else {
          r = r0 - 1;
        }
        for (let f = 0; f < values[r][0].length; f++) {
          if ((r == i && f >= k) || !searching(r, 0, f)) {
            continue;
          }
          if (
            i == r ||
            (stat == 0 &&
              (perRef[r] < 0 ||
                (perRef[i] > -1 && (perRef[i] != perRef[r] || periods[perRef[r]][0] != i)) ||
                (perRef[i] < 0 && periods[perRef[r]][0] > -1) ||
                stat2 == 0)) ||
            (stat > 0 &&
              ((perRef[i] > -1 && perRef[r] > -1 && perRef[i] != perRef[r] && periods[perRef[r]][stat] > -1) ||
                (perRef[i] < 0 && perRef[r] > -1 && periods[perRef[r]][stat] > -1)))
          ) {
            elem = values[r][0][f];
            putSugg(context, i, k);
            //   console.log("return4");
            return -1;
          }
          console.log("stop")
          stop = true;
          break;
        }
        if (stop) {
          break;
        }
      }
      if (stat || stop) {
        if (!stop) {
          r = i;
          stat2 = stat;
        }
        if (found(context, r, i, 0, k) == -1) {
          // console.log("return5");
          return -1;
        }
        if(periods) {
          console.log(`0periods : ${periods[0][0]} ${stat} ${stop}`)
        }
      }
    }
  }
  stop = false;
  // console.log(`prev ${prevLine}`);
  let attNames = [];
  attributes = [];
  let acc = 0;
  for (let j = 4; j < colNumb - 1; j++) {
    attNames.push([]);
    for (let i = 1; i < nbLineBef; i++) {
      for (let k = 0; k < values[i][j].length; k++) {
        if(isTip(context, i, j, k)) {
          return -1;
        }
        if (stat) {
          values[i][j][k] = elem;
          console.log("sugg2");
          sugg(context, i, j);
          return -1;
        }
        if (j < colNumb - 2) {
          if (perRef[i] != -1) {
            console.log("suggREf0");
            suggSet(context, i, j, []);
            return -1;
          }
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
        if(isTip(context, i, j, k)) {
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
              suggRef(context, i, j, k);
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
          suggRef(context, i, j, k);
          return -1;
        }
        stop = false;
        if (stat > 0) {
          if (found(context, r, i, j, k) == -1) {
            return -1;
          }
          if(periods) {
            console.log(`2periods : ${periods[0][0]}`)
          }
        }
        if (isCatF) {
          if (data[r][4].some((v) => v[1] == catId)) {
            suggRef(context, i, j, k);
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
    /*for(let q = 1; q<3; q++) {
      for(let m = 0; m<values[c][q].length; m++) {
        values[z][q].push(values[c][q][m]);
      }
    }
    for(let m = 0; m<values[c][3].length; m++) {
      values[s][3].push(values[c][3][m]);
    }
    for(let m = 1; m<3; m++) {
      for(let j = 4; j<colNumb; j++) {
        for(let k = 0; k<values[c][j].length; k++) {
          if(!values[periods[i][m]][j].map((e)=>e.toLowerCase()).includes(values[c][j][k].toLowerCase())) {
            values[periods[i][m]][j].push(values[c][j][k]);
          }
        }
      }
    }*/
  }
  console.log(" sssssssssss ")
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
              sugg(context, iR, 1);
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
        inconsist(context,
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
  // console.log(`prev ${prevLine}`);
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
    inconsist(context, values[i][0][0], 0, allLines, []);
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
                inconsist(context,
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
                  inconsist(context, dataAgg[i].lines[int], 0, "[" + cat + "]\n    " + allLines.slice(5), []);
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
                inconsist(context, dataAgg[path[1]].lines[0], 0, "   " + allLines.slice(3), []);
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
  let maxL = Math.max(nbLineBef, prevLine);
  let prevLineBef = prevLine;
  prevLine = 1;
  // let range = sheet.getRange(`A1:${getColumnTag(colNb + 10)}${Math.max(prevLine, nbLineBef)}`);
  // // Math.max(0, colors.length - 1) because of Excel API limitation

  // if(colors.length <= maxL) {
  //   // Create a 2D array to hold the range objects for each cell
  //   let cellRangeArray = [];
  //   for (var row = colors.length; row < maxL; row++) {
  //     cellRangeArray.push([]);
  //     // Loop through each column in the range
  //     for (let col = 0; col < colNumb + 10; col++) {
  //       let cell = range.getRange(`${getColumnTag(col)}${row + 1}`);
  //       // cell.load(["format/fill/color", "format/font/color", "format/borders/*"]);
  //       cell.load(["format/fill/color", "format/font/color"]);
  //       cellRangeArray[row].push(cell);
  //     }
  //   }
  //   await context.sync();
  //   for (var row = colors.length; row < maxL; row++) {
  //     colors.push([]);
  //     fonts.push([]);
  //     // Loop through each column in the range
  //     for (let col = 0; col < colNumb + 10; col++) {
  //       let cellBackgroundColor = cellRangeArray[row][col].format.fill.color;
  //       colors[row].push(cellBackgroundColor);
  //       let font = cellRangeArray[row][col].format.font.color;
  //       fonts[row].push(font);
  //     }
  //   }
    
  //   // // Build the new sheet data XML string
  //   // let sheetXml = `<sheet name="${escapeXml(sheet.name)}"><color>`;
  //   // colors.forEach(line => {
  //   //     sheetXml += `<message>${escapeXml(line)}</message>`;
  //   // });
  //   // sheetXml += '</color><font>';
  //   // colors.forEach(line => {
  //   //     sheetXml += `<message>${escapeXml(line)}</message>`;
  //   // });
  //   // sheetXml += '</sheet>';

  // //   var customXmlParts = context.workbook.customXmlParts;
  // //   try {
  // //       await customXmlParts.load('id');
  // //       await context.sync();

  // //       let xml = '<data><sheets></sheets><requests></requests></data>'; // Default XML structure if none exists

  // //       if (customXmlParts.items.length > 0) {
  // //           const part = customXmlParts.items[0];
  // //           await part.getXml();
  // //           await context.sync();
  // //           xml = part.xml;

  // //           // Replace the specific sheet part in the existing XML
  // //           const sheetRegex = new RegExp(`<sheet name="${escapeXml(sheet.name)}">.*?<\\/sheet>`, "s");
  // //           if (sheetRegex.test(xml)) {
  // //               xml = xml.replace(sheetRegex, sheetXml);
  // //           } else {
  // //               // If the sheet does not exist, append it within <sheets>
  // //               const sheetsEndTag = '</sheets>';
  // //               if (xml.includes(sheetsEndTag)) {
  // //                   xml = xml.replace(sheetsEndTag, sheetXml + sheetsEndTag);
  // //               } else {
  // //                   // If <sheets> part doesn't exist, create the structure
  // //                   xml = xml.replace('</data>', `<sheets>${sheetXml}</sheets></data>`);
  // //               }
  // //           }

  // //           part.delete();
  // //           await context.sync();
  // //       } else {
  // //           // Add the new sheet data to the default XML structure
  // //           xml = xml.replace('</sheets>', sheetXml + '</sheets>');
  // //       }

  // //       customXmlParts.add(xml);
  // //       await context.sync();

  // //       console.log("Sheet data replaced in custom XML part.");
  // //   } catch (error) {
  // //       console.error("An error occurred: " + error.message);
  // //   }
  // }
  console.log(`color loaded...`);
  // if (nbLineBefBef == -1) {
  //   nbLineBefBef = nbLineBef;
  // }
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
      prevLine = i + 1;
    }
    if(colors[i]!=color || fonts[i]!=font || updateRegularity) {
      var range = sheet.getRange(`A${i + 1}:${getColumnTag(colNumb-1)}${i + 1}`);
      if(colors[i]!=color || updateRegularity) {
        if(color) {
          range.format.fill.color = "#" + color;
        } else {
          range.format.fill.clear();
        }
      }
      if(fonts[i]!=font || updateRegularity) {
        range.format.font.color = "#" + font;
      }
      colors[i]=color
      fonts[i]=font
    }
  }
  iL = Math.min(prevLine, prevLineBef);
  iU = Math.max(prevLine, prevLineBef);
  for(let i = iL; i<iU; i++) {
    styleBorders(sheet, i, 0, colNumb, iL == prevLineBef);
  }
  document.getElementById("linkToCell").style.display = "none";
  updateSuggestions(context, 0, 0, []);
  range = context.workbook.getSelectedRange();
  range.load("address");
  await context.sync();
  resolved = true;
  handleSelectLinks(context, range.address);
}

async function styleBorders(sheet, i, jL, jU, draw) {
  for(let j = jL; j<jU; j++) {
    let cell = sheet.getRange(`${getColumnTag(j)}${i + 1}`);
    if(draw) {
      cell.format.borders.getItem('EdgeTop').style = Excel.BorderLineStyle.thin;
      cell.format.borders.getItem('EdgeTop').color = "black";
      cell.format.borders.getItem('EdgeBottom').style = Excel.BorderLineStyle.thin;
      cell.format.borders.getItem('EdgeBottom').color = "black";
      cell.format.borders.getItem('EdgeLeft').style = Excel.BorderLineStyle.thin;
      cell.format.borders.getItem('EdgeLeft').color = "black";
      cell.format.borders.getItem('EdgeRight').style = Excel.BorderLineStyle.thin;
      cell.format.borders.getItem('EdgeRight').color = "black";
    } else {
      cell.format.borders.getItem('EdgeTop').style = Excel.BorderLineStyle.none;
      cell.format.borders.getItem('EdgeBottom').style = Excel.BorderLineStyle.none;
      cell.format.borders.getItem('EdgeLeft').style = Excel.BorderLineStyle.none;
      cell.format.borders.getItem('EdgeRight').style = Excel.BorderLineStyle.none;
    }
  }
}

async function callVBA(context, functionName, ...args) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.load("name");
  console.log(11)
  await context.sync();

  args.unshift(sheet.name);
  const [xmlDoc, root] = await accessXML(context);


  
  // Create a new "msg" element
  const msgElement = xmlDoc.createElement("callVBA");

  // Create a new "name" element
  const nameElement = xmlDoc.createElement("functionName");
  nameElement.textContent = functionName;

  // Append the "name" element to the "msg" element
  msgElement.appendChild(nameElement);
  for (let i = 0; i < args.length; i++) {
    const argElement = xmlDoc.createElement("arg");
    argElement.textContent = args[i];
    msgElement.appendChild(argElement);
  }

  root.appendChild(msgElement);

  // Serialize the XML document back to a string
  const serializer = new XMLSerializer();
  const updatedXml = serializer.serializeToString(xmlDoc);

  // Set the updated XML back to the custom XML part
  customXMLPartDict[rootXMLMsgName].setXml(updatedXml);
  await context.sync();
  await releaseLockXML(context);
  let cell = sheet.getRange(`${getColumnTag(colNumb + 10)}1`);
  progChg++;
  console.log(`=callFunc()`)
  cell.formulas = [[`=callFunc()`]];
  await context.sync();
}

// // Function to display the entered text in a cell
// function displayTextInCell() {
//   // Get the input text value
//   var inputText = document.getElementById("inputText").value;

//   // Call Office JavaScript API to set the text in a cell (Replace this with your actual code to set text in a cell)
//   Office.context.document.setSelectedDataAsync(inputText, function(result) {
//     if (result.status === "succeeded") {
//       console.log("Text has been set in the cell.");
//     } else {
//       console.error("Error setting text in the cell:", result.error.message);
//     }
//   });
// }

// document.getElementById("displayButton").addEventListener("click", displayTextInCell);

// $("#run").on("click", () => tryCatch(run));


// /** Default helper for invoking an action and handling errors. */
// async function tryCatch(callback) {
//   try {
//     await callback();
//   } catch (error) {
//     // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
//     console.error(error);
//   }
// }

function updateSuggestions(i, j, suggestions) {
  const suggestionList = document.getElementById("suggestion-list");
  suggestionList.innerHTML = ""; // Clear previous suggestions

  suggestions.forEach(function (suggestionText) {
    const listItem = document.createElement("li");
    listItem.classList.add("ms-ListItem", "suggestion-item");
    const link = document.createElement("a");
    link.setAttribute("href", "#");
    link.setAttribute("data-text", suggestionText);
    link.textContent = suggestionText;
    listItem.appendChild(link);
    suggestionList.appendChild(listItem);
  });
  const suggestionItems = document.querySelectorAll(".suggestion-item");
  suggestionItems.forEach((item) => {
    item.addEventListener("click", (event) => myQueue.enqueue(handleSuggestionItemClick, 0, event, i, j));
  });
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
    sugg(context, i, j);
  } else {
    console.log("suggRef");
    suggSet(context, i, j, suggWords);
  }
}

function found(context, r, i, j, k) {
  if (perRef[r] == -1) {
    values[i][j].splice(k, 1);
    console.log(`sugg5 ${r} ${perRef[r]}`);
    sugg(context, i, j);
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
    console.log(`2- periods : ${periods[0][0]}`)
  }
  if (j == 0) {
    periods[perRef[r]][stat] = i;
    periods[perRef[r]][3] = periods[perRef[r]][3].concat(
      values[i][0]
        .filter((e) => e.startsWith(perNot[stat]))
        .map((e) => e.slice(perNot[stat].length))
        .filter((e) => !periods[perRef[r]][3].map((e) => e.toLowerCase()).includes(e.toLowerCase()))
    );
    console.log(`3- periods : ${periods[0][0]}`)
    perRef[i] = perRef[r];
  }
  if (i != r) {
    values[i][j][k] = perNot[stat] + elem2;
  }
  if(periods) {
    console.log(`periods : ${periods[0][0]}`)
  }
  return 1;
}

/**
  * elem <- the value of the referenced cell minus possible prefix
  * val  <- the value of elem in lowercase
  * 
  * @param {integer} i - the row index
  * @param {integer} j - the column index
  * @param {integer} k - the index in the cell
  * @return {integer} - -1 if it is only a prefix, 1 otherwise
  */
function isTip(context, i, j, k) {
  stat = 0;
  elem = values[i][j][k];
  val = elem.toLowerCase();
  for (let q = 1; q < 3; q++) {
    if (val.startsWith(perNot[q])) {
      elem = values[i][j][k].slice(perNot[q].length).trim();
      val = elem.toLowerCase();
      if (val == "") {
        elem = values[i][j][k];
        putSugg(context, i, k);
        return -1;
      }
      stat = q;
      break;
    }
  }
}

function searching(r, d, f) {
  console.log("searching")
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

function sugg(context, i, j) {
  console.log("sugg");
  suggSet(context, i, j, [values[i][j].join("; ")]);
}

function suggSet(context, i, j, sugg) {
  let cellAddress = `${getColumnTag(j)}${i + 1}`;
  inconsist(context, i, j, "Error at " + cellAddress, sugg);
}

function inconsist(context, i, j, message, sugg) {
  updateSuggestions(context, i, j, sugg);
  var linkToCell = document.getElementById("linkToCell");
  console.log(`prep : ${getColumnTag(j)}${i + 1}`);
  linkToCell.addEventListener("click", function (event) {
    event.preventDefault(); // Prevent default link behavior

    // Call Office JavaScript API to select the cell
    Excel.run(async context => {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let range = sheet.getRange(`${getColumnTag(j)}${i + 1}`);
      range.select();
    });
  });
  linkToCell.textContent = message;
  linkToCell.style.display = "block";
  // if (prevLine < i) {
  //   prevLine = i;
  // }
  exit();
}

// Function to toggle the visibility of the app body pane
function togglePaneVisibility() {
  var appBody = document.getElementById("app-body");
  if (appBody.style.display === "none") {
    appBody.style.display = "block";
  } else {
    appBody.style.display = "none";
  }
}

function exit() {
  // sheet.getRange(`${String.fromCharCode(65 + colNumb)}4:A${nbLineBef+10}`).format.fill.color = "#ff0000";
  // sheet.getRange(`${String.fromCharCode(65 + colNumb)}4`).format.font.color = "#000000";
}

function putSugg(context, i, k) {
  while (!stop) {
    let theMatch = elem.match(/(.*? -)(\d)$/);
    if (theMatch) {
      elem = theMatch[1] + (+theMatch[2] + 1);
    } else {
      elem += " -2";
    }
    val = elem.toLowerCase();
    if (stat > 0) {
      val = val.slice(perNot[stat].length);
    }
    stop = true;
    for (let m = 1; m < values.length; m++) {
      for (let p = 0; p < values[m][0].length; p++) {
        if ((m != i || p != k) && searching(m, 0, p)) {
          stop = false;
          break;
        }
      }
      if (!stop) {
        break;
      }
    }
  }
  if (stat) {
    elem = perNot[stat] + elem;
  }
  values[i][0][k] = elem;
  console.log("sugg1");
  sugg(context, i, 0);
}

function correct(context) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  for (let i = 1; i < nbLineBef; i++) {
    for (let j = 0; j < colNumb; j++) {
      if (!Array.isArray(values[i][j])) {
        console.log("NOT");
      }
      let value = values[i][j].join("; ");
      if (values0[i][j] != value) {
        values0[i][j] = value;
        progChg++;
        let cell = sheet.getRange(`${getColumnTag(j)}${i + 1}`);
        cell.values = [[value]];
      }
    }
  }
}

async function dataGenerator(context) {
  if (sorting) {
    await callVBA(context, "quit_sort");
    document.getElementById('updateButton').style.display = 'none';
  } else {
    let sortButton = document.getElementById("sortButton");
    sortButton.textContent = "Stop sorting";
    await dataGeneratorSub(context);
    document.getElementById('updateButton').style.display = 'block';
    sorting = true;
  }
}

async function updateFunc(context) {
  let sortButton = document.getElementById("updateButton");
  if (updating) {
    await callVBA(context, "not_updating_mode");
    sortButton.textContent = "Update";
  } else {
    await callVBA(context, "updating_mode");
    sortButton.textContent = "Stop updating";
  }
  updating = !updating;
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

function escapeXml(unsafe) {
  console.log(`unsafe : ${unsafe}`)
  return unsafe.replace(/[<>&'"]/g, function (c) {
    switch (c) {
      case '<': return '&lt;';
      case '>': return '&gt;';
      case '&': return '&amp;';
      case '\'': return '&apos;';
      case '"': return '&quot;';
    }
  });
}

async function dataGeneratorSub(context) {
  let resCheck = await check(context);
  if (resCheck == -1) {
    return -1;
  }

  correct(context);
  console.log(`data ${data.length}`);
  // return;
  console.log(`att ${attributes.length}`);
  let workbook = context.workbook;

  // Get the name of the Excel file
  console.log("1");
  workbook.load("name");
  await context.sync();
  let fileName = workbook.name;
  fileName = fileName.substring(0, fileName.lastIndexOf("."));
  console.log("2");

  // Get the name of the active sheet
  // let activeSheet = workbook.worksheets.getActiveWorksheet();
  // for (let i = 0; i < lenAgg; i++) {
  //   let difference = new Set(dataAgg[i].before);
  //   dataAgg[i].shortBefore.forEach(function(value) {
  //     difference.delete(value);
  //   });
  //   dataAgg[i].before = difference;
  // }


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
  // Adding xml attributes to transfer result through xml


  // var customXmlParts = context.workbook.customXmlParts;
  // console.log(`there2`);
  // customXmlParts.load('id'); // Load all parts to access their IDs

  // await context.sync().then(function () {
  //     // Check each part and delete if necessary
  //     console.log(`there2`);
  //     customXmlParts.items.forEach(function (part) {
  //         part.delete(); // Delete each found part
  //     });
  //     console.log(`there2`);

  //     return context.sync();
  // }).then(function () {
  //     // After all parts are deleted, add the new XML part
  //     console.log(`there2`);
  //     var newPart = customXmlParts.add(result);
  //     console.log(`there3`);
  //     return context.sync();
  // }).then(function () {
  //     console.log("Data added to custom XML part.");
  // });




  let sheet = context.workbook.worksheets.getActiveWorksheet();
  
  sheet.protection.protect({
      selections: Excel.ProtectionSelectionMode.none
  });
  await callVBA(context, "GetColumnText", result);

  
  // cell.values = [[`0`]];

  //const { exec } = require('child_process');
  // fetch('http://localhost:3001/command', { // Assuming your server is running on localhost:3000
  //   method: 'POST',
  //   headers: {
  //       'Content-Type': 'application/json'
  //   },
  //   body: JSON.stringify({
  //       result
  //   })
  // })
  // .then(response => {
  //   console.log("yes!");
  //     // Handle response from the server
  // })
  // .catch(error => {
  //     // Handle error
  //   console.log("no!");
  // });

  // sheet.getRange(1,colNumb+2).setValue("loading...");
  // let file = DriveApp.getFileById(fileId);
  // if (file) {
  //   let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //   //sheet.getRange(nbLineBef+10,colNumb + 1).setValue(result);
  //   file.setContent(result);
  // }

  // let currentDate = new Date();

  // let hours = ('0' + currentDate.getHours()).slice(-2);
  // let minutes = ('0' + currentDate.getMinutes()).slice(-2);
  // let seconds = ('0' + currentDate.getSeconds()).slice(-2);

  // let day = ('0' + currentDate.getDate()).slice(-2);
  // let year = currentDate.getFullYear();

  // let threeLetterMonth = currentDate.toLocaleString('en-US', { month: 'short' });

  // let formattedDate = hours + ':' + minutes + ':' + seconds + ' ' + day + ' ' + threeLetterMonth + ' ' + year;
  // sheet.getRange(1,colNumb+2).setValue("updated at "+formattedDate);
}

async function renameSymbol(context, oldValue, newValue) {
  let resCheck = await check(context);
  if (resCheck == -1) {
    console.log(-1)
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
  console.log(`renamed (${nbRen})`);
  correct(context);
  console.log(`renamed (${nbRen})`);
  document.getElementById("renamings").textContent = `renamed (${nbRen})`;
  check(context);
}

async function goToPeriodHead(context) {
  // await Excel.run(async (context) => {
  //   var customXmlParts = context.workbook.customXmlParts;
  //   try {
  //       await customXmlParts.load('id');
  //       await context.sync();

  //       console.log("Deleting existing XML parts...");
  //       customXmlParts.items.forEach(part => part.delete());
  //       await context.sync();

  //       console.log("Adding new XML part...");
  //       progChg++;
  //       var newPart = customXmlParts.add('<data><message>e/h\y</message></data>');
  //       await context.sync();

  //       console.log("Data added to custom XML part.");
  //   } catch (error) {
  //       console.error("An error occurred: " + error.message);
  //   }
  // });
  // return;
  console.log(`PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP`)
  let resCheck = await check(context);
  if (resCheck == -1) {
    return -1;
  }
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  console.log(`values.length : ${values.length}`)
  for (let t = 0; t < values[1][colNumb].length; t++) {
    val = values[1][colNumb][t].toLowerCase();
    for (let i = 1; i < nbLineBef; i++) {
      for (let k = 0; k < values[i][0].length; k++) {
        if (searching(i, 0, k)) {
          if (stat2 != 0 && periods[perRef[i]][0] != -1) {
            i = periods[perRef[i]][0];
          }
          let range = sheet.getRange(`A${i + 1}`);
          range.load("address");
          range.select();
          stop = 1;
          break;
        }
      }
      if (stop) {
        break;
      }
    }
    if (stop) {
      break;
    }
  }
  correct(context);
}

function handleoldNameInputClick(context) {
  // Get the current worksheet and then the selected range
  let range = context.workbook.getSelectedRange();

  // Load the row index and row count properties of the range
  range.load("rowIndex, rowCount");

  let rowId = range.rowIndex; // Directly use rowIndex property
  let totalRows = range.rowCount; // Get total rows in the selection
  for (let i = 0; i < totalRows; i++) {
    // Add each row index to rowNumbers

    rowId = Math.min(rowId, range.rowIndex);
  }
  document.getElementById("oldNameInput").value = values[rowId][0].join("; ");
  document.getElementById("newNameInput").value = values[rowId][0].join("; ");
}

function handlenewNameInputClick() {
  document.getElementById("newNameInput").value = document.getElementById("oldNameInput").value;
}

async function ctrlZ() {
  if (indOldValues > 0) {
    indOldValues--;
    values = oldValues[indOldValues];
    correct(context);
    check(context);
  }
}

async function ctrlY() {
  if (indOldValues < oldValues.length - 1) {
    indOldValues++;
    // values, colors, fonts = oldValues[indOldValues];
    values = oldValues[indOldValues];
    correct(context);
    check(context);
  }
}