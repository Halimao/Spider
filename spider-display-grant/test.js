
const request = require('request-promise-native');


/**
 * @method sleep方法
 * @param {*} ms 
 */
function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

async function test(){
    console.log(new Date());
    await sleep(5000);
    console.log(new Date());
}

test();