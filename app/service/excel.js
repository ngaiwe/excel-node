'use strict';

const Service = require('egg').Service;
const xlsx = require('node-xlsx').default;
const fs = require('fs');

class ExcelService extends Service {
  async getData(stream) {
    let buffers = [],
      data;
    await new Promise(resolve => {
      stream.on('data', chunk => {
        buffers.push(chunk);
      }).on('end', () => {
        buffers = Buffer.concat(buffers);
        data = xlsx.parse(buffers);
        resolve();
      }).on('error', err => {
        throw err;
      });
    });
    return data;
  }
  async realExcel(res) {
    const value = res[0];
    const datas = value.data.slice(4, value.data.length);
    const length = value.data[2].length - 23;
    let index = 0;
    let firstTitle = ['序号', '姓名'];
    let firstEndTitle = ['合计', '事假/年假/病假/丧假/婚假'];
    let cacheTitle = new Array(length);
    let secondTitle = [null, null];
    let secondEndTitle = ['出勤天数'];
    let secondTitleCache = [];
    while (index < length) {
      secondTitleCache.push(++index)
    }
    let first = [...firstTitle, ...cacheTitle, ...firstEndTitle];
    let second = [...secondTitle, ...secondTitleCache, ...secondEndTitle];
    let array = [first, second];
    datas.forEach((item, idx) => {
      let name = item[0];
      let workDays = item[4];
      let restDays = item[5];
      let synthesize = '';
      item = item.slice(23)
      item = item.map(el => {
        let text = '',
          chidao = /上班迟到(\d*)[小时]*(\d*)[分钟]*/,
          shijia = /事假.+ (\d+[.]?\d*)小时/,
          waichu = /外出.+ (\d+[.]?\d*)小时/,
          tiaoxiu = /调休.+ (\d+[.]?\d*)小时/,
          nianjia = /年假.+ (\d+[.]?\d*)天/,
          hunjia = /婚假.+ (\d+[.]?\d*)天/,
          sangjia = /丧假.+ (\d+[.]?\d*)天/,
          bingjia = /病假.+ (\d+[.]?\d*)小时/,
          chuchai = /出差.+ (\d+[.]?\d*)天/,
          xiuxi = /^休息/
        switch (el) {
          case '正常':
            text = '√'
            break;
          case '休息':
            text = null
            break;
          case '休息并打卡':
            text = '加班'
            break;
          default:
            if (xiuxi.test(el)) {
              text = null
            } else {
              text = el
            }
        }
        if (chidao.test(el)) {
          let chidaoM = chidao.exec(el)[2]
          if (chidaoM && chidaoM != 0) {
            text = `迟到${chidao.exec(el)[1]}.${chidaoM}h`
          } else {
            text = `迟到${chidao.exec(el)[1]}分钟`
          }
        }
        if (shijia.test(el)) {
          let val = shijia.exec(el)[1]
          text = `事假${bigEight(val, 'h')}`
        }
        if (waichu.test(el)) {
          text = `O`
        }
        if (tiaoxiu.test(el)) {
          text = `调休${bigEight(tiaoxiu.exec(el)[1], 'h')}`
        }
        if (nianjia.test(el)) {
          text = `年假1天`
        }
        if (hunjia.test(el)) {
          text = `婚假1天`
        }
        if (sangjia.test(el)) {
          text = `丧假1天`
        }
        if (bingjia.test(el)) {
          text = `病假${bigEight(bingjia.exec(el)[1], 'h')}`
        }
        if (chuchai.test(el)) {
          text = `出差1天`
        }
        return text
      })
      synthesize = filterArrEach(['事假', '年假', '病假', '丧假', '婚假'], item)
      let status = [++idx, name, ...item, workDays, synthesize];
      array.push(status)
    });
    return {
      name: '出勤表',
      data: array,
    };
  }
  async writeFile(res) {
    var buffer = xlsx.build([res]);
    fs.writeFileSync('./zhoubao.xlsx', buffer);
  }
}

// 处理大于8小时
var bigEight = (value, exec = null) => {
  let boolean = Number(value) >= 8
  if (exec) {
    return boolean ? '1天' : `${value}${exec}`
  } else {
    return boolean ? 8 : value
  }
}

var filterArr = (title, item) => {
  let value = 0
  item.forEach(el => {
    let reg = RegExp(`^(${title})(\\d+)(.)$`);
    if (reg.test(el)) {
      let arr = reg.exec(el),
        title = arr[1],
        val = arr[2];
      val = arr[3] == '天' ? val * 8 : val
      value += Number(val)
    } else {
      void null
    }
  })
  if (Boolean(value)) {
    let day = ''
    value = Number(value)
    if (value >= 8) {
      let h = value % 8
      if (Boolean(h)) {
        h = `${h}h`
      } else {
        h = ''
      }
      day = `${Math.floor(value/8)}天${h}`
    } else {
      day = `${value}h`
    }
    return `${title}${day}`
  } else {
    return ''
  }
}

var filterArrEach = (arrs = [], item) => {
  return  arrs.map(arr => filterArr(arr, item)).join(' ').trim()
}

module.exports = ExcelService;