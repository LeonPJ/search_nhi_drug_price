const axios = require('axios');
const XLSX = require('xlsx');
const path = require('path');

const drugListFileName = 'drug_list.xlsx';
const drugPriceFileName = 'drug_price.xlsx';

const drugListFilePath = path.join(__dirname, drugListFileName);
const drugPriceFilePath = path.join(__dirname, drugPriceFileName);

const workbook = XLSX.readFile(drugListFilePath);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const jsonData = XLSX.utils.sheet_to_json(worksheet);

const newWorkbook = XLSX.utils.book_new();
const newWorksheet = XLSX.utils.aoa_to_sheet([['NUM', 'CODE', 'NAME', 'PRICE']]);

async function delay(milliseconds) {
    return new Promise(resolve => {
        setTimeout(resolve, milliseconds);
    });
}

async function searchDrugPrice(num, drugCode) {
    const payload = {
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
    };

    try {
        const res = await axios.post('https://info.nhi.gov.tw/api/INAE3000/INAE3000S01/SQL0001', payload, {
            headers: {
                'Accept': '*/*',
                'User-Agent': 'Thunder Client (https://www.thunderclient.com)',
                'Content-Type': 'application/json'
            }
        });

        const drugCode = res.data.data[0].druG_CODE;
        const drugName = res.data.data[0].druG_ENAME;
        const drugPrice = res.data.data[0].paY_PRICE;
        XLSX.utils.sheet_add_aoa(newWorksheet, [[num, drugCode, drugName, drugPrice]], { origin: -1 });

        console.log(`${drugCode} ${drugName} ${drugPrice}`);
    } catch (err) {
        // console.log(err);
        XLSX.utils.sheet_add_aoa(newWorksheet, [[num, drugCode, null, null]], { origin: -1 });
        console.log(`${drugCode}`);
    }
}

async function processDrugList() {
    console.log('**** DRUG SEARCH PRICE V1.0 ***');
    for (let i = 0; i < jsonData.length; i++) {
        const drugCode = jsonData[i].code;
        await searchDrugPrice(i + 1, drugCode);
    }

    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1');
    XLSX.writeFile(newWorkbook, drugPriceFilePath);
    console.log('**** SEARCH FINISHED ***');
}

processDrugList();
