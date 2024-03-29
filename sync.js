const DOCUMENT_ID = "1rDjtvS4Em9JDvJULABNndxK0rY0HAMMGvee2nk7GWFI";
const GET_COMPANIES_URL = "https://cmp.wildberries.ru/backend/api/v3/stats/atrevds?pageNumber={page}&pageSize=100&search=&type=null";
const GET_ITEMS_URL = "https://cmp.wildberries.ru/backend/api/v3/fullstat/{company_id}";
const SHEET_NAME_COMPANIES = "ads_source"
const SHEET_NAME_ITEMS = "items_source_{page}"
const SHEET_NAME_APP = "main"
const SYNC_DATE_CELL = "B1"
const WB_TOKEN_STATS_CELL = "B4"
const WB_TOKEN_ITEMS_CELL = "B5"
const COMPANIES_PER_PAGE = 100;
const ITEMS_PER_CALL = 30;
const SLEEP_TIME = 5 * 1000;

function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Custom WB menu')
        .addItem('Sync WB adv-companies only', 'syncCompanies')
        .addItem('Sync WB items page 1', 'syncItemsPage1')
        .addItem('Sync WB items page 2', 'syncItemsPage2')
        .addItem('Sync WB items page 3', 'syncItemsPage3')
        .addItem('Sync WB items page 4', 'syncItemsPage4')
        .addItem('Sync WB items page 5', 'syncItemsPage5')
        .addItem('Sync WB items page 6', 'syncItemsPage6')
        .addItem('Sync WB items page 7', 'syncItemsPage7')
        .addItem('Sync WB items page 8', 'syncItemsPage8')
        .addItem('Sync WB items page 9', 'syncItemsPage9')
        .addItem('Sync WB items page 10', 'syncItemsPage10')
        .addToUi();
}


function syncCompanies() {
    console.log('sync companies start');

    let googleSpreadSheet = SpreadsheetApp.openById(DOCUMENT_ID);

    //load companies
    let companies = _fetchCompanies(
        _getStatsToken(googleSpreadSheet)
    );
    console.log('parse companies %s', companies.length);

    //fill companies sheet
    let companiesSheet = googleSpreadSheet.getSheetByName(SHEET_NAME_COMPANIES);
    companiesSheet.clear({contentsOnly: true});
    companiesSheet.getRange(1, 1, 1, 5).setValues([['id', 'title', 'start date', 'end date', 'amount']]);
    if (companies.length) {
        companiesSheet.getRange(2, 1, companies.length, companies[0].length).setValues(companies);
    }
    console.log('companies source up success');

    //up sync date
    let currentDateTime = (new Date()).toLocaleString();
    googleSpreadSheet.getSheetByName(
        SHEET_NAME_APP,
    ).getRange(
        SYNC_DATE_CELL,
    ).setValue(
        currentDateTime,
    );
    console.log('sync companies complete');
}


function syncItemsPage1(page = 1) {
    syncItems(page);
}
function syncItemsPage2(page = 2) {
    syncItems(page);
}
function syncItemsPage3(page = 3) {
    syncItems(page);
}
function syncItemsPage4(page = 4) {
    syncItems(page);
}
function syncItemsPage5(page = 5) {
    syncItems(page);
}
function syncItemsPage6(page = 6) {
    syncItems(page);
}
function syncItemsPage7(page = 7) {
    syncItems(page);
}
function syncItemsPage8(page = 8) {
    syncItems(page);
}
function syncItemsPage9(page = 9) {
    syncItems(page);
}
function syncItemsPage10(page = 10) {
    syncItems(page);
}


function syncItems(page) {
    console.log('sync companies start page=%d', page);
    const startRowOffset = ITEMS_PER_CALL * (page - 1);

    let googleSpreadSheet = SpreadsheetApp.openById(DOCUMENT_ID);

    let companiesSheet = googleSpreadSheet.getSheetByName(SHEET_NAME_COMPANIES);
    let companies = companiesSheet.getRange(2 + startRowOffset, 1, ITEMS_PER_CALL, 5).getValues().filter(el => el[0]);
    console.log('debug', companies);

    //load items for all companies
    let items = _fetchItems(
        _getItemsToken(googleSpreadSheet),
        companies
    );
    console.log('parse items %s', items.length);

    //fill items sheet
    let itemsSheet = googleSpreadSheet.getSheetByName(
        SHEET_NAME_ITEMS.replace('{page}', page)
    );
    itemsSheet.clear({contentsOnly: true});
    itemsSheet.getRange(1, 1, 1, 6).setValues([['company id', 'id', 'title', 'start date', 'end date', 'amount']]);
    if (items.length){
        itemsSheet.getRange(2, 1, items.length, items[0].length).setValues(items);
    }
    console.log('items page %d updated', page);
}


function _getStatsToken(document) {
    // fixme refresh token by WB API
    let token = document.getSheetByName(
        SHEET_NAME_APP,
    ).getRange(
        WB_TOKEN_STATS_CELL,
    ).getValue();
    console.log('get stats token %s', token)
    return token;
}


function _getItemsToken(document) {
    // fixme refresh token by WB API
    let token = document.getSheetByName(
        SHEET_NAME_APP,
    ).getRange(
        WB_TOKEN_ITEMS_CELL,
    ).getValue();
    console.log('get items token %s', token)
    return token;
}


function _compareStatRow(a, b) {
    if (a[0] < b[0]) {
        return 1;
    }
    if (a[0] > b[0]) {
        return -1;
    }
    // a must be equal to b
    return 0;
}


function _fetchCompanies(token){
    let pageNumber = 1;
    let data = [];
    let hasNextPage = true;
    while (hasNextPage) {
        let statsResponseRaw = UrlFetchApp.fetch(
            GET_COMPANIES_URL.replace('{page}', pageNumber),
            {
                'headers': {
                    'accept': 'application/json',
                    'content-type': 'application/json',
                    'referer': 'https://cmp.wildberries.ru/statistics',
                    'user-agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.75 Safari/537.36',
                    'X-User-Id': '97739323',
                    'cookie': `BasketUID=4547d2b3-790c-4a23-8553-c5ff2bc843e5; _wbauid=4482704701664446372; ___wbu=a3cc44fd-6839-4e14-902c-a38a0e9d8d8a.1664555179; __bsa=basket-ru-43; __tm=1669556224; um=uid%3Dw7TDssOkw7PCu8K4wrbCtsKywrjCssKzwrI%253d%3Aproc%3D100%3Aehash%3Dd41d8cd98f00b204e9800998ecf8427e; WILDAUTHNEW_V3=13E92D91E7061BC866357F7B04DD2F1CE18AD44FB0CF8538C268CA19646FC6053A03CAC24DD96DF3225A7C0256D868E37169F3E0EAD7A1D788F90EB2B0F443FBE9EDA434C095E143256EAE796F3B8C775719DF25DE7E8A23DC5A839585091DF17F1332066DB9B5584B471500C3B08E1087222F5A507CE5FCC570B17448DAE3DE5773259031BF491BEE17AA6AAB34DF9FE2E912F3B5E78FE9727DB717D6C786CCE7CA62DBD4FE369E2995E88E01A5A6572DD60A142D7E200B101946ABA1515A6D78D37EAE92D23ABE57D17DCCE0A37B71BF86A99C87F3DFB6815A6DB27B2A29A92A1D14B2D1DD3EC96503BBAF63A91F982D27354DD0472A8A7C1585C29A8BFB3E9F5501F709799786FFE624CA41E75211F94D6DC28CD88F532DC47AFCDC8E9450DF01741A;__wbl=cityId%3D0%26regionId%3D0%26city%3D%D0%9C%D0%BE%D1%81%D0%BA%D0%B2%D0%B0%26phone%3D84957755505%26latitude%3D55%2C7558%26longitude%3D37%2C6176%26src%3D1; x-supplier-id-external=40b958e6-2e4b-4ca9-8712-75d775012b97; WBToken=${token}; x-supplier-id=40b958e6-2e4b-4ca9-8712-75d775012b97`,
                },
                'contentType': 'application/json',
                'method': 'GET',
                'validateHttpsCertificates': false,
            },
        ).getContentText();
        console.log('stats response %s', statsResponseRaw);
        let statsResponse = JSON.parse(statsResponseRaw);
        pageNumber += 1;
        if (statsResponse.pageCount < COMPANIES_PER_PAGE) {
            hasNextPage = false;
        }

        statsResponse.content.forEach((element) => {
            data.push([
                parseInt(element.Id),
                element.CampaignName,
                element.Begin,
                element.End,
                parseFloat(element.Sum),
            ])
        });
    }
    data.sort(_compareStatRow);
    return data;
}


function _fetchItems(token, companies){
    let data = [];
    for (const company of companies) {
        let headers = {
            'accept': 'application/json',
            'content-type': 'application/json',
            'referer': `https://cmp.wildberries.ru/statistics/${company[0]}`,
            'user-agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.75 Safari/537.36',
            'X-User-Id': '97739323',
            'cookie': `BasketUID=4547d2b3-790c-4a23-8553-c5ff2bc843e5; _wbauid=4482704701664446372; ___wbu=a3cc44fd-6839-4e14-902c-a38a0e9d8d8a.1664555179; __bsa=basket-ru-43; WILDAUTHNEW_V3=13E92D91E7061BC866357F7B04DD2F1CE18AD44FB0CF8538C268CA19646FC6053A03CAC24DD96DF3225A7C0256D868E37169F3E0EAD7A1D788F90EB2B0F443FBE9EDA434C095E143256EAE796F3B8C775719DF25DE7E8A23DC5A839585091DF17F1332066DB9B5584B471500C3B08E1087222F5A507CE5FCC570B17448DAE3DE5773259031BF491BEE17AA6AAB34DF9FE2E912F3B5E78FE9727DB717D6C786CCE7CA62DBD4FE369E2995E88E01A5A6572DD60A142D7E200B101946ABA1515A6D78D37EAE92D23ABE57D17DCCE0A37B71BF86A99C87F3DFB6815A6DB27B2A29A92A1D14B2D1DD3EC96503BBAF63A91F982D27354DD0472A8A7C1585C29A8BFB3E9F5501F709799786FFE624CA41E75211F94D6DC28CD88F532DC47AFCDC8E9450DF01741A; x-supplier-id-external=40b958e6-2e4b-4ca9-8712-75d775012b97; x-supplier-id=40b958e6-2e4b-4ca9-8712-75d775012b97; WBToken=${token}`,
        };
        Utilities.sleep(SLEEP_TIME);
        let itemsResponseRaw = UrlFetchApp.fetch(
            GET_ITEMS_URL.replace('{company_id}', company[0]),
            {
                'headers': headers,
                'contentType': 'application/json',
                'method': 'GET',
                'validateHttpsCertificates': false,
            },
        ).getContentText();
        console.info('items %d response %s', company[0], itemsResponseRaw);
        let itemsResponse = JSON.parse(itemsResponseRaw);
        let items = itemsResponse.content.days.map(el => el.apps).flatMap(el => el).map(el => el.nm).flatMap(el => el);
        let localSummary = _sumByKey(items, 'nmId', 'sum');
        console.log('items summary', localSummary);
        for (const item of localSummary) {
            let itemId = parseInt(item.nmId);
            if (item.sum && itemId) {
                data.push([
                    company[0],
                    itemId,
                    company[1],
                    company[2],
                    company[3],
                    item.sum,
                ]);
            }
        }
    }
    return data;
}


function _sumByKey(arr, key, value) {
    const map = new Map();
    for(const obj of arr) {
        const currSum = map.get(obj[key]) || 0;
        map.set(obj[key], currSum + obj[value]);
    }
    return Array.from(map, ([k, v]) => ({[key]: k, [value]: v}));
}
