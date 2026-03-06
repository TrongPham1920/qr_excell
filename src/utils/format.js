const path = require("path");
const { PNG } = require("pngjs");
const jsQR = require("jsqr");

const normalizeFileName = (fileName) => {
  let name = path.parse(fileName).name;

  if (name.startsWith("._")) {
    name = name.replace("._", "");
  }

  name = name.replace(/'/g, "").trim();

  return name;
}

const readQrFromPng = (buffer) => {
  return new Promise((resolve) => {
    new PNG().parse(buffer, (err, data) => {
      if (err) return resolve(null);

      const qr = jsQR(
        new Uint8ClampedArray(data.data),
        data.width,
        data.height,
      );

      resolve(qr ? qr.data : null);
    });
  });  
}

module.exports = {
  normalizeFileName,
  readQrFromPng,
};