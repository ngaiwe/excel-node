'use strict';

/**
 * @param {Egg.Application} app - egg application
 */
module.exports = app => {
  const { router, controller } = app;
  router.get('/', controller.excel.index);
  router.post('/upload/excel', controller.upload.add);
};
