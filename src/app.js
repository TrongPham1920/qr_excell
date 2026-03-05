const express = require('express');
const routes = require('./server/routes');

const app = express();

app.use(routes);

module.exports = app;