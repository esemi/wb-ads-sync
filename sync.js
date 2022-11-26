

const SOURCES_HOST = "https://www.wildberries.ru/webapi/security/login/data?returnUrl=https%3A%2F%2Fwww.wildberries.ru%2F";
const DOCUMENT_ID = "1rDjtvS4Em9JDvJULABNndxK0rY0HAMMGvee2nk7GWFI";
const SHEET_NAME_DATA = "ads_source"
const SHEET_NAME_APP = "main"
const SYNC_DATE_CELL = "B1"

function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Custom WB menu')
        .addItem('Sync WB run', 'syncRun')
        .addToUi();
}


function syncRun() {
    console.log('sync start');

    //todo login to wb
    let loginResponseRaw = UrlFetchApp.fetch(SOURCES_HOST, {
        'method': 'POST',
    }).getContentText();
    console.log('login response %s', loginResponseRaw);
    let loginResponse = JSON.parse(loginResponseRaw);
    console.log('login response data %s', loginResponse);

    //todo fetch webtoken
    //todo fetch ads-campaigns data (with pagination?)

    let googleSpreadSheet = SpreadsheetApp.openById(DOCUMENT_ID);
    let dataSheet = googleSpreadSheet.getSheetByName(SHEET_NAME_DATA);
    let data = [
        ['id', 'title', 'start date', 'end date', 'amount']
    ];

    //fill values
    console.log('parse values %s', data);
    dataSheet.clear({contentsOnly: true});
    dataSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    console.log('source up success');

    //up sync date
    let currentDateTime = (new Date()).toLocaleString();
    googleSpreadSheet.getSheetByName(SHEET_NAME_APP).getRange(SYNC_DATE_CELL).setValue(
        currentDateTime
    );
    console.log('sync complete');
}
