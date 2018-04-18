const cheerio = require('cheerio');
const request = require('request-promise-native');
const xlsx = require('node-xlsx');
const fs = require('fs');


/**
 * @method sleep方法
 * @param {*} ms 
 */
function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

async function test(){
    //读取文件内容
    var obj = xlsx.parse('ces18_Exhibitor.xlsx');
    var excelObj = obj[0].data;
    var companyArray = [];
    for(var i in excelObj){
        var value = excelObj[i];
        companyArray.push(value[0]);
    }
    console.log(companyArray);
    return companyArray;
}

test();