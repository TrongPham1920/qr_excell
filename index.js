const app = require('./src/app');
const routes = require('./src/server/routes');

const PORT = 3000;

app.use('/', routes);

app.listen(PORT, () => {
  console.log(`Server is running on port http://localhost:${PORT}`);
});