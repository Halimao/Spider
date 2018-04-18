// @ts-check

const cheerio = require('cheerio');
const request = require('request-promise-native');
const xlsx = require('node-xlsx');
const fs = require('fs');

// 在此处填写表单内的搜索条件
var qs = {
    'searchtype': 'subsearch',
    'searchValue': 'Romy',
    'section': 'exhibitorkeyword',
    'ShowReturnToTop': 'false',
    'fpsearchtype': 'keyword',
    'fpkeyword': 'Romy',
    'searchTime': '1516698455531',
    'increment': '100'
};
var headers = {
    'Cache-Control': 'no-cache',
    'x-requested-with': 'XMLHttpRequest'
};
var companyArray = initialCompanyData();
//var companyArray = new Array('2realistic', 'Romy');
// 执行数据爬取
spider(qs, headers, companyArray);

async function spider(qs, headers, companyArray) {
    var startTime = new Date().getTime();
    var detailUrlObj = {};
    try {
        for (var i = 0; i < companyArray.length; i++) {
            try {
                qs['searchValue'] = qs['fpkeyword'] = companyArray[i];
                var url = 'https://ces18.mapyourshow.com/7_0/search/_search-results.cfm';
                var body = await retry(request.get, 5)(url, {
                    qs: qs,
                    headers: headers,
                    timeout: 60 * 1000
                }); 
                var data = JSON.parse(body);
                var dataHtml = data['DATA']['BODYHTML'];
                var $ = cheerio.load(dataHtml);
                var detailUrl = 'https://ces18.mapyourshow.com' + $('.mys-table-exhname').find('a').attr('href');
                detailUrlObj[companyArray[i]] = detailUrl;
                console.log('Inital ' + i + 'th Detail Info Url-------');
            } catch (iErr) {
                console.log('Initial detail info url about company named ' + companyArray[i] + ' failed!');
                continue;
            }
        }

        var resultObj = await spiderDetailInfo(detailUrlObj);
        var infosArray = resultObj['infosArray'];
        var failedArray = resultObj['failedArray'];
        
        var excelData = [{
            name: 'The King',// 构造sheet名称
            // 构造数据
            data: [
                ['Company', 'Phone', 'Address', 'Web Url', 'About Corporation', 'Product Categories', 'Company Contacts']// 构造标题行
            ]
        }];
        infosArray.forEach(function(value, index){
            excelData[0].data.push([value['Company'], value['Phone'], value['Address'], value['Web Url'], value['About Corporation'], value['Product Categories'], value['Company Contacts']]);
        });
            
        // 生成数据缓冲区
        var buffer = xlsx.build(excelData);
        // 生成Excel文件
        fs.writeFile('./result.xlsx', buffer, function (err) {
            if (err) {
                throw err;
            }
            console.log('Write to result.xlsx has finished');
        });
        var endTime = new Date().getTime();
        console.log('Time spent: ' + (endTime - startTime) + 'ms');
    } catch (oErr) {
        console.error(oErr);
    }
}

/**
 * @method 爬取数据, node版本要支持async关键字
 * @param {object} detailUrlObj 
 * @return object
 */
async function spiderDetailInfo(detailUrlObj) {
    var count = 0;
    var infosArray = new Array();
    var failedArray = new Array();
    console.log('Start to get detail info......');
    for (var key in detailUrlObj) {
        let infoObj = {
            'Company': '',
            'Phone': '',
            'Address': '',
            'Web Url': '',
            'About Corporation': '',
            'Product Categories': '',
            'Company Contacts': ''
        };
        console.log(++count + "th, company------" + key + ",detailUrlObj[key]------" + detailUrlObj[key]);
        infoObj['Company'] = key;// 赋值公司名
        try {
            await sleep(2000);// 停留2秒再请求,防止IP被屏蔽
            const body = await retry(request.get, 5)(detailUrlObj[key], {
                timeout: 60 * 1000
            });
            var $ = cheerio.load(body);
            // 获取顶部的地址、电话信息、官网地址
            $('#jq-sc-Mobile-ExhContactInfo').find('p').each(function () {
                var cssName = $(this).attr('class');
                if(cssName == 'sc-Exhibitor_Address'){
                    //console.log('Address------' + $(this).text().trim());
                    infoObj['Address'] = $(this).text().trim();
                } else if(cssName == 'sc-Exhibitor_Url') {
                    //console.log('Url-------' + $(this).text().trim());
                    infoObj['Web Url'] = $(this).text().trim();
                } else if(cssName == 'sc-Exhibitor_PhoneFax') {
                    //console.log('Phone-------' + $(this).text().trim());
                    infoObj['Phone'] = $(this).text().trim();
                }
            });
            // 获取底部About Corporation
            $('#mys-exhibitor-details-wrapper').find('div.mys-taper-measure').each(function () {
                //console.log('About Corporation------' + $(this).text().trim());
                infoObj['About Corporation'] = $(this).text().trim();
            });
            // 获取底部Product Categories、Company Contacts信息
            $('#mys-exhibitor-details-wrapper').find('.mys-insideToggle').each(function () {
                var curCate = $(this).parent().prev().find('strong').text().trim();
                if(curCate.indexOf('Product Categories') > -1){
                    //console.log('Product Categories------' + $(this).text().trim());
                    infoObj['Product Categories'] = $(this).text().trim();
                } else if(curCate.indexOf('Company Contacts') > -1){
                    //console.log('Company Contacts------' + $(this).text().trim());
                    infoObj['Company Contacts'] = $(this).text().trim();
                }
            });
            infosArray.push(infoObj);
        } catch (err) {
            //console.error(err);
            console.log('Getting detail info about company named ' + key + ' failed');
            failedArray.push(key);
            // 失败一个跳过不抛出异常
            continue;
        }
    }
    return {
        'infosArray': infosArray,
        'failedArray': failedArray
    };
}

/**
 * @method 读取Excel文件里的所有公司
 */
function initialCompanyData() {
    //读取Excel文件里的所有公司
    var obj = xlsx.parse('ces18_Exhibitor.xlsx');
    var excelObj = obj[0].data;
    var companyArray = [];
    for(var i in excelObj){
        var value = excelObj[i];
        companyArray.push(value[0]);
    }
    console.log('Initial company data success, start to spider......');
    return companyArray;
}

/**
 * @method sleep休眠方法
 * @param {number} delay, 单位ms 
 */
function sleep(delay) {
    return new Promise(resolve => setTimeout(resolve, delay));
}

/**
 * @method 失败后重试
 * @param {function} fn 
 * @param {number} times 
 */
function retry(fn, times = 10) {
    return async function (...params) {
        while (times > 0) {
            try {
                return await fn(...params);
            } catch (err) {
                await sleep(1000);
                times--;
                if (times === 0) {
                    throw err;
                }
            }
        }
    }
}