// @ts-check

const cheerio = require('cheerio');
const request = require('request-promise-native');
const xlsx = require('node-xlsx');
const fs = require('fs');

// 在此处填写表单内的搜索条件
var formData = {
    grantee_code: '',
    product_code: '',
    applicant_name: '',
    comments: '',
    application_purpose: '',
    application_status: '',
    modular_type_description: '',
    grant_code_1: '',
    grant_code_2: '',
    grant_code_3: '',
    equipment_class: '',
    lower_frequency: '',
    upper_frequency: '',
    emission_designator: '',
    bandwidth_from: '',
    tolerance_from: '',
    tolerance_to: '',
    power_output_from: '',
    power_output_to: '',
    rule_part_1: '',
    rule_part_2: '',
    rule_part_3: '',
    product_description: '',
    tcb_code: '',
    tcb_scope: '',
    application_purpose_description: '',
    equipment_class_description: '',
    tcb_code_description: '',
    grant_date_from: '01/01/2015', // 起始日期
    grant_date_to: '12/31/2015', // 结束日期
    fetchfrom: 0,
    show_records: 500, // 显示多少条数据,超出总数则显示所有条数
    test_firm: 'BTL',
    outputformat: 'HTML',
    tolerance_exact_match: 'on', // Ignore
    freq_exact_match: 'on', // Ignore
    power_exact_match: 'on', // Ignore
    rule_part_exact_match: 'on', // Ignore
    calledFromFrame: 'N' // Ignore
};
// 执行数据爬取
spider(formData, true, 486);

/**
 * @method 爬取数据, node版本要支持async关键字
 * @param {object} formData 查询过滤条件
 * @param {boolean} needValidate 设置是否需要结果验证
 * @param {number} actuallyTotalCount 实际数据总数, 当且仅当needValidate为true时有效
 * @return void
 */
async function spider(formData, needValidate, actuallyTotalCount) {
    const url = 'https://apps.fcc.gov/oetcf/eas/reports/GenericSearchResult.cfm?RequestTimeout=500';
    const body = await retry(request.post, 5)(url, {
        formData: formData,
        timeout: 60 * 1000
    });
    try {
        const $ = cheerio.load(body);
        var startTime = new Date().getTime();
        var fccIDPurposeArray = [];
        var grantUrlGroup = {};
        var count = 0;
        var totalCount = 0;
        console.log('Request ' + url + ' has returned data, start to data processing......');
        $('#offTblBdy tr').each(function () {
            totalCount++;
            var dispGrantUrl = $(this).find('td').eq(3).find('a').attr('href');//获取Display Grant
            var fccId = $(this).find('td').eq(11).text().trim();//获取FCC ID
            var appPurpose = $(this).find('td').eq(12).text().trim();//获取Application Purpose
            var fccPurpose = fccId + '_##_' + appPurpose;
            if(typeof(dispGrantUrl) == 'undefined' || dispGrantUrl.length == 0){
                throw new Error(fccPurpose + '--dispGrantUrl is null!');
            }
            // 如果是已有旧数据
            if (fccIDPurposeArray.indexOf(fccPurpose) > -1) {
                grantUrlGroup[fccPurpose]['num'] = grantUrlGroup[fccPurpose]['num'] + 1;
            } else {
                // 如果是新数据
                count++;
                fccIDPurposeArray.push(fccPurpose);
                grantUrlGroup[fccPurpose] = {
                    'num': 1,
                    'fccId': fccId,
                    'purpose': appPurpose,
                    'url': 'https://apps.fcc.gov' + dispGrantUrl
                };
            }
        });
        console.log('Total number: ' + totalCount);
        console.log('(Duplicate Removal)Total number: ' + count);
        // 数据验证,如果查询返回的结果总条数与实际数据总条数不一致则抛出异常
        if(needValidate && totalCount != actuallyTotalCount){
            throw new Error("totalCount doesn't match ActuallyTotalCount!");
        }
        if (count > 0) {
            const displayGrantsObject = await initialSpiderData(grantUrlGroup);
            // 开始写入Excel文件里,此处使用node-xlsx第三方库
            var excelData = [{
                name: 'Handsome Baby Monkey',// 构造sheet名称
                // 构造数据
                data: [
                    ['FCC ID', 'Application Purpose', 'Name of Grantee', 'Equipment Class', 'Notes', 'TCB', 'Count']// 构造标题行
                ]
            }];
            for (const key in displayGrantsObject) {
                if (displayGrantsObject.hasOwnProperty(key)) {
                    const element = displayGrantsObject[key];
                    excelData[0].data.push([element['FCC ID'], element['Application Purpose'], element['Name of Grantee'], element['Equipment Class'], element['Notes'], element['TCB'], element['Num']]);
                }
            }
            // 生成数据缓冲区
            var buffer = xlsx.build(excelData);
            // 生成Excel文件
            fs.writeFile('./result.xlsx', buffer, function (err) {
                if (err) {
                    throw err;
                }
                console.log('Write to result.xlsx has finished');
            });
        }
        console.log('Total number: ' + totalCount);
        console.log('(Duplicate Removal)Total number: ' + count);
        var endTime = new Date().getTime();
        console.log('Time spent: ' + (endTime - startTime) + 'ms');
    } catch (err) {
        console.error(err);
    }
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

/**
 * @method 爬取数据, node版本要支持async关键字
 * @param {object} grantUrlGroup 
 * @return object
 */
async function initialSpiderData(grantUrlGroup) {
    var displayGrantsObject = {};
    var count = 0;
    for (var key in grantUrlGroup) {
        console.log(++count + "th, grantUrlGroup[key]---------" + grantUrlGroup[key]['url']);
        try {
            await sleep(2000);// 停留2秒再请求
            const body = await retry(request.get, 5)(grantUrlGroup[key]['url'], {
                timeout: 60 * 1000
            });
            const $ = cheerio.load(body);
            var tcb = $('body').find('table').eq(2).find('td').eq(1).html();
            tcb = tcb.split('<br')[0].trim();// 取出第一行，并去除左右空格
            var fccId = grantUrlGroup[key]['fccId'];
            var appPurpose = grantUrlGroup[key]['purpose'];
            var name = $('body').find('div[align="CENTER"]').find('tr').eq(2).find('td').eq(1).text();
            var equipmentCls = $('body').find('div[align="CENTER"]').find('tr').eq(4).find('td').eq(1).text();
            var notes = $('body').find('div[align="CENTER"]').find('tr').eq(5).find('td').eq(1).text();
            var num = grantUrlGroup[key]['num'];
            /* console.log("key---------" + key);
            console.log("fccId---------" + fccId);
            console.log("appPurpose--------" + appPurpose);
            console.log('name---------' + name);
            console.log('equipmentCls-------' + equipmentCls);
            console.log('notes--------' + notes);
            console.log('tcb-------' + tcb);
            console.log("num---------" + num); */
            var displayGrants = {
                'FCC ID': fccId,
                'Application Purpose': appPurpose,
                'Name of Grantee': name,
                'Equipment Class': equipmentCls,
                'Notes': notes,
                'TCB': tcb,
                'Num': num
            }
            displayGrantsObject[key] = displayGrants;
        } catch (err) {
            console.error(err);
        }
    }
    return displayGrantsObject
}