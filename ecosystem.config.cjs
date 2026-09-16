module.exports = {
  apps: [
    {
      name: "qr-excell",
      script: "index.js",
      instances: 1,
      autorestart: true,
      watch: false,
      max_memory_restart: "256M",
      env: {
        NODE_ENV: "production",
        PORT: 8001,
      },
      env_file: ".env",
    },
  ],
};
