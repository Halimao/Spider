// @ts-check

const cheerio = require('cheerio');
const request = require('request');
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
    grant_date_from: '01/01/2017', // 起始日期
    grant_date_to: '12/31/2017', // 结束日期
    FromRec: 1, // 从第几条开始查询
    show_records: 10, // 显示多少条数据,超出总数则显示所有条数
    prior_value: 'Show Previous 500 Rows', // Ignore
    test_firm: 'BTL',
    tolerance_exact_match: 'on', // Ignore
    freq_exact_match: 'on', // Ignore
    power_exact_match: 'on', // Ignore
    rule_part_exact_match: 'on', // Ignore
    calledFromFrame: 'N' // Ignore
};
// 执行数据爬取
spider(formData, false, []);

/**
 * @method 爬取数据
 * @param {object} formData 查询过滤条件
 * @param {boolean} needFilterSite true/false,是否需要根据国家筛选
 * @param {Array<string>} givenSiteArray 当needFilterSite为true时有效,传入需要筛选的国家名称数组
 * @return void
 */
function spider(formData, needFilterSite, givenSiteArray) {
    const url = 'https://apps.fcc.gov/oetcf/eas/reports/GenericSearchResult.cfm?RequestTimeout=500'
    request.post(url, {
        formData: formData
    }, (err, res, body) => {
        if (!err) {
            const $ = cheerio.load(body)
            var companyArray = [];
            var companyGroup = {};
            var count = 0;
            $('#offTblBdy tr').each(function () {
                count++;
                // 判断是否需要根据国家进行筛选
                if (needFilterSite) {
                    // 需要筛选国家，判断国家是否在指定国家列表里
                    // 如果设置了需要筛选国家，但是没有指定筛选的国家则提示
                    if (givenSiteArray === undefined || givenSiteArray.length == 0) {
                        console.error('You need define the country you wanna to filter!');
                        return;
                    } else {
                        var countryName = $(this).find('td').eq(9).text();//获取每行的国家
                        if (givenSiteArray.indexOf(countryName) > -1) {
                            var companyName = $(this).find('td').eq(5).text();//获取每行的公司名
                            if (companyArray.indexOf(companyName) > -1) {
                                // 如果数组里已经存在这个公司了，则将公司的出现次数加1
                                companyGroup[companyName] = companyGroup[companyName] + 1;
                            } else {
                                // 如果数组里不存在这个公司, 则将公司的出现次数设置为1，并加入数组
                                companyGroup[companyName] = 1;
                                companyArray.push(companyName);
                            }
                        }
                    }
                    // 不需要筛选国家，只筛选公司
                } else {
                    var companyName = $(this).find('td').eq(5).text();//获取每行的公司名
                    if (companyArray.indexOf(companyName) > -1) {
                        // 如果数组里已经存在这个公司了，则将公司的出现次数加1
                        companyGroup[companyName] = companyGroup[companyName] + 1;
                    } else {
                        // 如果数组里不存在这个公司, 则将公司的出现次数设置为1，并加入数组
                        companyGroup[companyName] = 1;
                        companyArray.push(companyName);
                    }
                }
            }
            );
            console.log('Total number: ' + count);
            // 开始写入Excel文件里,此处使用node-xlsx第三方库
            var excelData = [{
                name: 'Handsome Baby Monkey',// 构造sheet名称
                // 构造数据
                data: [
                    ['Company Name', 'Count']// 构造标题行
                ]
            }];
            // 将查询返回的公司数据存入
            for (const key in companyGroup) {
                if (companyGroup.hasOwnProperty(key)) {
                    const element = companyGroup[key];
                    excelData[0].data.push([key, element]);
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
                // 读result.xlsx
                //var obj = xlsx.parse("./" + "result.xlsx");
                //console.log(JSON.stringify(obj));
            });
        } else {
            console.error(err)
        }
    })
}