/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const ExtractCSSPlugin = require("mini-css-extract-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");
const webpack = require("webpack");

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { cacert: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  // const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      vendor: ["react", "react-dom", "core-js", "@fluentui/react"],
      taskpane: ["react-hot-loader/patch", path.resolve(__dirname, "./src/test.index.tsx")],
    },
    output: {
      path: path.resolve(__dirname, "testBuild"),
      devtoolModuleFilenameTemplate: "webpack:///[resource-path]?[loaders]",
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
      fallback: {
        child_process: false,
        fs: false,
        os: require.resolve("os-browserify/browser"),
      },
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: ["react-hot-loader/webpack", "ts-loader"],
          exclude: /node_modules/,
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.css$/,
          use: [ExtractCSSPlugin.loader, "css-loader"],
        },
        {
          test: /\.(png|jpe?g|gif|svg|woff|woff2|ttf|eot|ico)$/,
          loader: "file-loader",
          options: {
            name: "[path][name].[ext]",
          },
        },
      ],
    },
    plugins: [
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"],
        process: "process/browser",
      }),
      new CleanWebpackPlugin(),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: path.resolve(__dirname, "./../src/taskpane/taskpane.css"),
            to: "taskpane.css",
          },
          {
            from: "./assets",
            to: "assets",
          },
        ],
      }),
      new ExtractCSSPlugin({ filename: "[name].[hash].css" }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: path.resolve(__dirname, "./src/test-taskpane.html"),
        chunks: ["taskpane", "vendor", "polyfills"],
      }),
    ],
    devServer: {
      static: {
        directory: path.join(__dirname, "testBuild"),
      },
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      https: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
