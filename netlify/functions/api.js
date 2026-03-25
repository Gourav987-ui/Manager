const serverless = require('serverless-http');
const { connectLambda } = require('@netlify/blobs');
const { app } = require('../../server');

const handler = serverless(app, {
  binary: [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel',
    'application/zip',
  ],
});

exports.handler = async (event, context) => {
  if (event?.blobs) {
    connectLambda(event);
  }
  return handler(event, context);
};
