const SOURCES_HOST = "https://cmp.wildberries.ru/backend/api/v3/stats/atrevds?pageNumber={page}&pageSize=100&search=&type=null";
const DOCUMENT_ID = "1rDjtvS4Em9JDvJULABNndxK0rY0HAMMGvee2nk7GWFI";
const SHEET_NAME_DATA = "ads_source"
const SHEET_NAME_APP = "main"
const SYNC_DATE_CELL = "B1"
const WB_TOKEN_CELL = "B4"
const ITEMS_PER_PAGE = 100;


function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Custom WB menu')
        .addItem('Sync WB run', 'syncRun')
        .addToUi();
}


function syncRun() {
    console.log('sync start');

    // get saved token
    let googleSpreadSheet = SpreadsheetApp.openById(DOCUMENT_ID);
    let token = googleSpreadSheet.getSheetByName(
        SHEET_NAME_APP,
    ).getRange(
        WB_TOKEN_CELL,
    ).getValue();
    console.log('get token %s', token)

    let pageNumber = 1;
    let data = [
        ['id', 'title', 'start date', 'end date', 'amount']
    ];
    let hasNextPage = true;
    while (hasNextPage) {
        let statsResponseRaw = UrlFetchApp.fetch(
            SOURCES_HOST.replace('{page}', pageNumber),
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
        if (statsResponse.pageCount < ITEMS_PER_PAGE) {
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
    data.sort(compareStatRow)
    console.log('parse values %s', data.length);

    let dataSheet = googleSpreadSheet.getSheetByName(SHEET_NAME_DATA);

    //fill values
    dataSheet.clear({contentsOnly: true});
    dataSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    console.log('source up success');

    //up sync date
    let currentDateTime = (new Date()).toLocaleString();
    googleSpreadSheet.getSheetByName(
        SHEET_NAME_APP,
    ).getRange(
        SYNC_DATE_CELL,
    ).setValue(
        currentDateTime,
    );
    console.log('sync complete');
}


function compareStatRow(a, b) {
    if (a[0] < b[0]) {
        return 1;
    }
    if (a[0] > b[0]) {
        return -1;
    }
    // a must be equal to b
    return 0;
}