'use strict';

const Controller = require('egg').Controller;

class ExcelController extends Controller {
  async index() {
    const { ctx } = this;
    await ctx.render('excel/list.tpl.html');
  }
}

module.exports = ExcelController;
