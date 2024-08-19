const GasPlugin = require("gas-webpack-plugin");
module.exports = {
  context: __dirname,
  entry: "./main.js",
  output: {
    path: __dirname ,
    filename: 'Code.gs'
  },
  plugins: [
    new GasPlugin()
  ]
}
