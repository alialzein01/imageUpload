const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = {
  mode: 'development',
  entry: {
    taskpane: './src/taskpane/taskpane.js',
  },
  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: '[name].js',
    publicPath: 'https://localhost:3000/',
    clean: true
  },
  devServer: {
    port: 3000,
    https: true,
    hot: true,
    headers: {
      'Access-Control-Allow-Origin': '*',
      'Cache-Control': 'no-cache, no-store, must-revalidate'
    },
    static: [
      {
        directory: path.join(__dirname, 'dist'),
        publicPath: '/'
      }
    ]
  },
  module: {
    rules: [
      {
        test: /\.(png|svg|jpg|jpeg|gif)$/i,
        type: 'asset/resource'
      }
    ]
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: './src/taskpane/taskpane.html',
      filename: 'taskpane.html',
      chunks: ['taskpane']
    }),
    new CopyWebpackPlugin({
      patterns: [
        {
          from: 'src/assets',
          to: 'assets'
        },
        {
          from: 'manifest.xml',
          to: 'manifest.xml'
        },
        {
          from: 'src/taskpane/taskpane.css',
          to: 'taskpane.css'
        }
      ]
    })
  ],
  resolve: {
    extensions: ['.js', '.json']
  }
};