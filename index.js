const app = require('./src/app');
const swagger = require('./swagger');

swagger(app);

const PORT = 3000;

app.listen(PORT, () => {
  console.log(`Server is running on port http://localhost:${PORT}`);
});