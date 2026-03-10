// Presented by BrilliantPy

// Editable
const sheetName = 'พัสดุ_2';
const index_col = {
    'ประทับเวลา': 0,
    'ชื่อหน้ากล่อง': 2,
    'เบอร์โทร': 3,
    'หน้ากล่อง': 4,
    'ของที่อยู่ในกล่อง': 5,
    'ID': 6,
    'ราคาเทิร์น': 7,
    'ราคาขาย': 8,
    'ยอดฝากรวม': 9,
    'สถานะ': 10,
    'ยอดฝาก': 11
};
const order_select_col = [index_col['ID'], index_col['ประทับเวลา'], index_col['เบอร์โทร'], index_col['ชื่อหน้ากล่อง'], index_col['หน้ากล่อง'], index_col['ของที่อยู่ในกล่อง'], index_col['ราคาเทิร์น'], index_col['ราคาขาย'], index_col['ยอดฝากรวม'], index_col['ยอดฝาก'], index_col['สถานะ']];
// let isFirstNumCol = false; 

const idColNum = 7, turnColNum = 8, sellColNum = 9, totalColNum = 10, statColNum = 11, depositColNum = 12;

// function doGet() {
//   return HtmlService.createTemplateFromFile('page1').evaluate()
//     .addMetaTag('viewport', 'width=device-width, initial-scale=1')
//     .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
// }

// function doGet(e) {
//   let page = (e && e.parameter && e.parameter.page) ? e.parameter.page : "index";
//   var output = HtmlService.createTemplateFromFile(page);
//   return output.evaluate()
//       .setTitle('พัสดุ')
//       .addMetaTag('viewport', 'width=device-width, initial-scale=1')
//       .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
// }

function doGet(e) {
    let page = e.parameter.page || "index";
    let html = HtmlService.createTemplateFromFile(page).evaluate();
    let htmlOutput = HtmlService.createHtmlOutput(html);
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1').setTitle('พัสดุ');

    //Replace {{NAVBAR}} with the Navbar content
    htmlOutput.setContent(htmlOutput.getContent().replace("{{NAVBAR}}", getNavbar(page)));
    return htmlOutput;
}

function getNavbar(activePage) {
    var scriptURLHome = getScriptURL();
    var scriptURLPage1 = getScriptURL("page=upload");

    var navbar =
        `
    <style>
.navbar-toggler:focus,
.navbar-toggler:active,
.navbar-toggler-icon:focus {
    outline: none !important;
    box-shadow: none !important;
    border: 0 !important;
}

</style>
    <nav class="navbar navbar-dark bg-primary sticky-top navbar-expand-lg">
      <div class="container-fluid">
        <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarNav"
          aria-controls="navbarNav" aria-expanded="false" aria-label="Toggle navigation">
          <span class="navbar-toggler-icon"></span>
        </button>
        <div class="collapse navbar-collapse" id="navbarNav">
          <ul class="navbar-nav">
            <li class="nav-item">
              <a class="nav-link ${activePage === 'index' ? 'active' : ''}" href="${scriptURLHome}">พัสดุ</a>
            </li>
            <li class="nav-item">
              <a class="nav-link ${activePage === 'upload' ? 'active' : ''}" href="${scriptURLPage1}">อัพโหลด</a>
            </li>
          </ul>
        </div>
      </div>
    </nav>
    <br>
    <script src="https://code.jquery.com/jquery-3.7.1.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"
      integrity="sha384-YvpcrYf0tY3lHB60NNkmXc5s9fDVZLESaAA55NDzOxhy9GkcIdslK1eN7N6jIeHz"
      crossorigin="anonymous"></script>`;
    return navbar;
}

function getScriptURL(qs = null) {
    var url = ScriptApp.getService().getUrl();
    if (qs) {
        if (qs.indexOf("?") === -1) {
            qs = "?" + qs;
        }
        url = url + qs;
    }
    return url;
}

// function getURL() {
//   var url = ScriptApp.getService().getUrl();
//   return url;
// }


function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Init
let ss, sheet, lastRow, lastCol, range, values;
function initSpreadSheet() {
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1ofRxVvq3e_bnb5Bol3kNpmhEsdytNBQaiGdY7pwdykU/edit?gid=1754669522#gid=1754669522');
    sheet = ss.getSheetByName(sheetName);
    lastRow = sheet.getLastRow();
    lastCol = sheet.getLastColumn();
    range = sheet.getDataRange();
    values = range.getDisplayValues();
    // console.log("lastRow:", lastRow, "lastCol:", lastCol); 
}

function getData() {
    initSpreadSheet();
    let format_values = formatData(values);
    format_values = JSON.stringify(format_values);
    return format_values;
}

function formatData(_data) {
    let data = _data;
    data.shift();
    for (let i = 0; i < data.length; i++) {
        let cur_data = data[i];
        let new_data = [];
        // if (isFirstNumCol) {
        //   new_data.push(i+1);
        // }
        for (let j = 0; j < order_select_col.length; j++) {
            new_data.push(cur_data[order_select_col[j]]);
        }
        data[i] = new_data;
    }
    return data;
}

function getIdData(id) {
    initSpreadSheet();
    let dataArr = [];

    for (var i = 2; i <= lastRow; i++) {
        let currId = sheet.getRange(i, 1).getValue();
        if (currId == id) {
            dataArr = sheet.getRange(i, 1, 1, lastCol).getValues();
            dataArr = dataArr[0];
            break;
        }
    }

    dataArr = JSON.stringify(dataArr);
    // Logger.log(dataArr);
    return dataArr;
}

//Update
function updateData(id, turn, sell, stat, total, deposit) {
    initSpreadSheet();
    for (var i = 2; i <= lastRow; i++) {
        let currId = sheet.getRange(i, idColNum).getValue();
        if (currId == id) {
            if (stat != "") {
                sheet.getRange(i, statColNum).setValue(stat);
            }
            if (turn != "") {
                sheet.getRange(i, turnColNum).setValue(turn);
            }
            if (sell != "") {
                sheet.getRange(i, sellColNum).setValue(sell);
            }
            sheet.getRange(i, totalColNum).setValue(total);
            sheet.getRange(i, depositColNum).setValue(deposit);

            return "SUCCESS";
        }
    }
    return "Error";
}

function removeData(id) {
    initSpreadSheet();
    for (var i = 2; i <= lastRow; i++) {
        let currId = sheet.getRange(i, idColNum).getValue();
        if (currId == id) {
            sheet.deleteRow(i);
            break;
        }
    }
    return "SUCCESS";
}

function saveFile1(obj) {
    var blob = Utilities.newBlob(Utilities.base64Decode(obj.data), obj.mimeType, obj.fileName);
    var folder = DriveApp.getFolderById('1x4_KoHW1lZt-AmApCAFxnCOaQjs01V-PxJjEdD_wSbXhLla7-NGmIg-ufJtAobMPiC0l32Lh');
    var driveFile = folder.createFile(blob);
    return driveFile.getUrl();
}

function saveFile2(obj) {
    var blob = Utilities.newBlob(Utilities.base64Decode(obj.data), obj.mimeType, obj.fileName);
    var folder = DriveApp.getFolderById('1rJMTqw1JpNfKAH_rOBt-RLqoQgUXQgeVLka760zszh0DuopNlvPU7VybxLmj0gfiMguMMroJ');
    var driveFile = folder.createFile(blob);
    return driveFile.getUrl();
}

function addData(name, phone, parcelName, item, total) {
    const formattedPhone = `'${phone}`;
    try {
        initSpreadSheet(); // Assuming this function initializes the spreadsheet
        sheet.appendRow([new Date(), "", name, formattedPhone, saveFile1(parcelName), saveFile2(item), "", "", "", "", "ไม่เคลียร์", total]);
        return "SUCCESS";
    } catch (error) {
        console.error("Error adding data:", error);
        return "ERROR";
    }
}

function test() {
    return getURL();
}
