const axios = require('axios');
const fsp = require('fs/promises');
const { EOL } = require('os');
const { XMLParser, XMLValidator } = require("fast-xml-parser");
require('dotenv').config();

const EBAY_APP_ID = process.env.EBAY_APP_ID || '';
const EBAY_DEV_ID = process.env.EBAY_DEV_ID || '';
const EBAY_CERT_ID = process.env.EBAY_CERT_ID || '';
const EBAY_AUTH_TOKEN = process.env.EBAY_AUTH_TOKEN || '';
const EBAY_TRADING_URL = process.env.EBAY_TRADING_URL || '';
const EBAY_SEARCH_URL = process.env.EBAY_SEARCH_URL || '';
const EBAY_API_SITEID = process.env.EBAY_API_SITEID || '0';
const SUCCESS_MSG = 'Success';

const searchConfig = {
    keywords: process.env.KEYWORDS,
    locatedIn: process.env.LOCATED_IN,
    listingType: process.env.LISTING_TYPE,
    minPrice: process.env.MIN_PRICE || '0',
    maxPrice: process.env.MAX_PRICE || '100000000',
    currency: process.env.CURRENCY || 'USD',
    condition: process.env.CONDITION,
    sortOrder: process.env.SORT_ORDER,
    maxNumEntries: parseInt(process.env.MAX_NUM_ENTRIES),
    soldItemsOnly: process.env.SOLD_ITEMS_ONLY,
};

function readXml(xml) {
    if (!XMLValidator.validate(xml))
        console.error('Not XML format');
    const parsed = new XMLParser().parse(xml);
    const item = parsed.GetItemResponse.Item;
    const nameValueList = item.ItemSpecifics.NameValueList;
    const mpn = nameValueList && nameValueList instanceof Array
        ? nameValueList.find(elem => elem.Name === 'Manufacturer Part Number')?.Value
        : '';
    const sold = item.SellingStatus?.QuantitySold || '';

    return { mpn, sold };
}

async function getMpnAndSold(itemId = "") {
    let payload = `<?xml version="1.0" encoding="utf-8"?>
    <GetItemRequest xmlns="urn:ebay:apis:eBLBaseComponents">
        <RequesterCredentials>
            <eBayAuthToken>${EBAY_AUTH_TOKEN}</eBayAuthToken>
        </RequesterCredentials>
        <ErrorLanguage>en_US</ErrorLanguage>
        <WarningLevel>High</WarningLevel>
        <DetailLevel>ItemReturnAttributes</DetailLevel>
        <ItemID>${itemId}</ItemID>
        <IncludeItemSpecifics>true</IncludeItemSpecifics>
    </GetItemRequest>`;

    const headers = {
        'X-EBAY-API-APP-NAME': EBAY_APP_ID,
        'X-EBAY-API-DEV-NAME': EBAY_DEV_ID,
        'X-EBAY-API-CERT-NAME': EBAY_CERT_ID,
        'X-EBAY-API-CALL-NAME': 'GetItem',
        'X-EBAY-API-SITEID': EBAY_API_SITEID,
        'X-EBAY-API-REQUEST-Encoding': 'XML',
        'X-EBAY-API-COMPATIBILITY-LEVEL': '1227',
        'Content-Type': 'text/xml',
    };
    const resp = await axios.post(EBAY_TRADING_URL, payload, { headers });
    let ret;
    try {
        ret = readXml(resp.data);
    } catch (err) {
        console.error(`Error: ItemId=${itemId}`);
        throw err;
    }
    return ret;
}

function yyyymmddhhmiss() {
    const pad2 = (n) => (n < 10 ? '0' : '') + n;
    const d = new Date();
    return `${d.getFullYear()}${pad2(d.getMonth() + 1)}${pad2(d.getDate())}`
        + `${pad2(d.getHours())}${pad2(d.getMinutes())}${pad2(d.getSeconds())}`;
}

function getBaseUrl() {
    let baseUrl = EBAY_SEARCH_URL;
    baseUrl += '?OPERATION-NAME=findItemsByKeywords';
    baseUrl += '&SERVICE-VERSION=1.13.0';
    baseUrl += '&SECURITY-APPNAME=' + EBAY_APP_ID;
    baseUrl += '&RESPONSE-DATA-FORMAT=JSON';
    baseUrl += '&REST-PAYLOAD';
    baseUrl += '&keywords=' + searchConfig.keywords;
    baseUrl += '&itemFilter(0).name=LocatedIn';
    baseUrl += '&itemFilter(0).value=' + searchConfig.locatedIn;
    baseUrl += '&itemFilter(1).name=ListingType';
    baseUrl += '&itemFilter(1).value=' + searchConfig.listingType;
    baseUrl += '&itemFilter(2).name=MinPrice';
    baseUrl += '&itemFilter(2).value=' + searchConfig.minPrice;
    baseUrl += '&itemFilter(2).paramName=Currency';
    baseUrl += '&itemFilter(2).paramValue=' + searchConfig.currency;
    baseUrl += '&itemFilter(3).name=MaxPrice';
    baseUrl += '&itemFilter(3).value=' + searchConfig.maxPrice;
    baseUrl += '&itemFilter(3).paramName=Currency';
    baseUrl += '&itemFilter(3).paramValue=' + searchConfig.currency;
    baseUrl += '&itemFilter(4).name=SoldItemsOnly';
    baseUrl += '&itemFilter(4).value=' + searchConfig.soldItemsOnly;
    baseUrl += '&itemFilter(5).name=Condition';
    baseUrl += '&itemFilter(5).value=' + searchConfig.condition;
    baseUrl += '&outputSelector(0)=SellerInfo';
    baseUrl += '&sortOrder=' + searchConfig.sortOrder;
    baseUrl += '&paginationInput.entriesPerPage=100';
    baseUrl += '&paginationInput.pageNumber=';
    return baseUrl;
}

async function convert2CSV(searchResult) {
    const count = searchResult[0]['@count'];
    if (count === '0') {
        console.error('Result is empty');
        process.exit(1);
    }
    const items = searchResult[0].item;

    let csv = `itemId,title,viewItemURL,sellerUserName,price,`
    csv += `categoryName,watchCount,startTime,endTime,mpn,sold${EOL}`;
    for (const item of items) {
        const itemId = item.itemId[0];
        const title = item.title[0].replace(/,/g, '');
        const viewItemURL = item.viewItemURL[0];
        const sellerUserName = item.sellerInfo[0].sellerUserName;
        const price = item.sellingStatus[0].convertedCurrentPrice[0].__value__;
        const categoryName = item.primaryCategory[0].categoryName[0];
        const watchCount = item.listingInfo[0].watchCount
            ? item.listingInfo[0].watchCount[0]
            : '';
        const startTime = item.listingInfo[0].startTime
            ? item.listingInfo[0].startTime[0]
            : '';
        const endTime = item.listingInfo[0].endTime
            ? item.listingInfo[0].endTime[0]
            : '';
        const { mpn, sold } = await getMpnAndSold(itemId);

        csv += `${itemId},${title},${viewItemURL},${sellerUserName},${price},`;
        csv += `${categoryName},${watchCount},${startTime},${endTime},`;
        csv += `${mpn || ''},${sold}${EOL}`;
    }
    return csv;
}

async function main() {

    let baseUrl = getBaseUrl();
    let pageNumber = 1;
    const entriesPerPage = 100;
    let maxNumEntries = searchConfig.maxNumEntries;

    const fileName = `./data/${searchConfig.keywords.replace(' ', '_')}_${yyyymmddhhmiss()}.csv`;

    /* Call synchronously (pageNumber increment) */
    while (pageNumber * entriesPerPage <= maxNumEntries) {

        try {
            /* Call Rest API */
            const encodedUrl = encodeURI(`${baseUrl}${pageNumber}`);
            const resp = await axios.get(encodedUrl, {
                responseType: 'json',
                responseEncoding: 'utf8',
            });

            if (resp.status !== 200)
                throw new Error(`Status=${resp.status}, ${resp.statusText}`);

            const itemsResp = resp.data?.findItemsByKeywordsResponse;
            if (!itemsResp) throw new Error('No data');

            const _ack = itemsResp[0].ack[0];
            if (_ack !== SUCCESS_MSG) {
                const errMsg = itemsResp[0].errorMessage[0].error[0].message[0];
                console.error(errMsg);
                throw new Error('Request failed.');
            }

            const { searchResult, paginationOutput } = itemsResp[0];

            const totalEntries = parseInt(paginationOutput[0].totalEntries[0]);
            if (maxNumEntries > totalEntries) maxNumEntries = totalEntries;

            const csv = await convert2CSV(searchResult);

            /* Write to file */
            try {
                await fsp.appendFile(fileName, csv, 'utf-8');
            } catch (ferr) {
                throw ferr;
            }

        } catch (err) {
            console.error(err);
            throw err;
        }

        ++pageNumber;
    }

}

main();