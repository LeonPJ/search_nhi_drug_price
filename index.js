const request = require('request');
const XLSX = require('xlsx');
const path = require('path');

// file path
const drugListFileName = 'drug_list.xlsx';
const drugPriceFileName = 'drug_price.xlsx';

// set file path
const drugListFilePath = path.join(process.cwd(), drugListFileName);
const drugPriceFilePath = path.join(process.cwd(), drugPriceFileName);

// 读取药品列表文件
const workbook = XLSX.readFile(drugListFilePath);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const jsonData = XLSX.utils.sheet_to_json(worksheet);

// create work space and sheet
const newWorkbook = XLSX.utils.book_new();
const newWorksheet = XLSX.utils.aoa_to_sheet([['NUM', 'CODE', 'NAME', 'MANUFACTURE', 'PRICE']]);

async function delay(milliseconds) {
    return new Promise(resolve => setTimeout(resolve, milliseconds));
}

// search formate
async function searchDrugPrice(num, drugCode) {
    const payload = JSON.stringify({
        "DRUG_CODE": drugCode,
        "DRUG_NAME": "",
        "DRUG_DOSE": "",
        "DRUG_CLASSIFY_NAME": "",
        "DRUG_ING": "",
        "DRUG_ING_QTY": "",
        "DRUG_ING_UNIT": "",
        "DRUG_STD_QTY": "",
        "DRUG_STD_UNIT": "",
        "DRUGGIST_NAME": "",
        "MIXTURE": "",
        "PAY_START_DATE_YEAR": "",
        "PAY_START_DATE_MON": "",
        "ORAL_TYPE": "",
        "ATC_CODE": "",
        "SHOWTYPE": "Y",
        "CURPAGE": 1,
        "PAGESIZE": 10
    });

    return new Promise((resolve, reject) => {
        request.post({
            url: 'https://info.nhi.gov.tw/api/INAE3000/INAE3000S01/SQL0001',
            body: payload,
            headers: {
                'Accept': '*/*',
                'User-Agent': 'Thunder Client (https://www.thunderclient.com)',
                'Content-Type': 'application/json',
                'Content-Length': Buffer.byteLength(payload)
            }
        }, (err, res, body) => {
            if (err) {
                XLSX.utils.sheet_add_aoa(newWorksheet, [[num, drugCode, null, 'ERROR']], { origin: -1 });
                console.error(`Error fetching data for drug code ${drugCode}:`, err);
                return reject(err);
            }

            try {
                const data = JSON.parse(body);
                if (data.data && data.data.length > 0) {

                    const drugData = data.data[0];
                    const drugCode = drugData.druG_CODE;
                    const drugName = drugData.druG_ENAME;
                    const drugPrice = drugData.paY_PRICE;
                    const drugManufacturer = drugData.druggisT_NAME;

                    XLSX.utils.sheet_add_aoa(newWorksheet, [[num, drugCode, drugName, drugManufacturer, drugPrice]], { origin: -1 });
                    console.log(`${drugCode} ${drugName} ${drugManufacturer} ${drugPrice}`);
                } else {
                    XLSX.utils.sheet_add_aoa(newWorksheet, [[num, drugCode, null, 'NOT FOUND']], { origin: -1 });
                    console.log(`${drugCode} not found`);
                }
                resolve();
            } catch (parseError) {
                XLSX.utils.sheet_add_aoa(newWorksheet, [[num, drugCode, null, 'PARSE ERROR']], { origin: -1 });
                console.error(`Error parsing response for drug code ${drugCode}:`, parseError);
                resolve(parseError);
            }
        });
    });
}

// main function
async function processDrugList() {
    console.log('**** DRUG SEARCH PRICE V1.1 ****');
    for (let i = 0; i < jsonData.length; i++) {
        const drugCode = jsonData[i].code;
        await searchDrugPrice(i + 1, drugCode);
        // await delay(1000);
    }

    // add to workspace
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1');
    XLSX.writeFile(newWorkbook, drugPriceFilePath);
    console.log('**** SEARCH FINISHED ****');
}

// 
processDrugList();
