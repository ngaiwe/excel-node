'use strict';

const Controller = require('egg').Controller;
const xlsx = require('node-xlsx').default;
const fs = require('fs');

class ExcelController extends Controller {
  async add() {
    const {
      ctx
    } = this;
    try {
      const stream = await ctx.getFileStream();
      let res = await ctx.service.excel.getData(stream);
      let data = await ctx.service.excel.realExcel(res);
      ctx.service.excel.writeFile(data)
      ctx.body = data
    } catch (error) {
      throw error;
    }
  }
}

module.exports = ExcelController;