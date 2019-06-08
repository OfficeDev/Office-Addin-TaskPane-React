const devCerts = require("office-addin-dev-certs");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const ExtractTextPlugin = require('extract-text-webpack-plugin');
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require('path');
const webpack = require('webpack');

module.exports = async (env, options) => {
    const dev = options.mode === "development";
    const config = {
        devtool: "source-map",
        entry: {
            vendor: [
                'react',
                'react-dom',
                'core-js',
                'office-ui-fabric-react'
            ],
            taskpane: [
                'react-hot-loader/patch',
                path.resolve(__dirname, './src/test.index.tsx')
            ]
        },
        output: { path: path.resolve(__dirname, "testBuild") },
        resolve: {
            extensions: [".ts", ".tsx", ".html", ".js"]
        },
        node: {
            child_process: 'empty'
        },
        module: {
            rules: [
                {
                    test: /\.tsx?$/,
                    use: [
                        'react-hot-loader/webpack',
                        'ts-loader'
                    ],
                    exclude: /node_modules/
                },
                {
                    test: /\.css$/,
                    use: ['style-loader', 'css-loader']
                },
                {
                    test: /\.(png|jpe?g|gif|svg|woff|woff2|ttf|eot|ico)$/,
                    use: {
                        loader: 'file-loader',
                        query: {
                            name: 'assets/[name].[ext]'
                        }
                    }
                }
            ]
        },
        plugins: [
            new CleanWebpackPlugin(),
            new CopyWebpackPlugin([
                {
                    to: "taskpane.css",
                    from: path.resolve(__dirname, './../src/taskpane/taskpane.css')
                }
            ]),
            new ExtractTextPlugin('[name].[hash].css'),
            new HtmlWebpackPlugin({
                filename: "taskpane.html",
                template: path.resolve(__dirname, './src/test-taskpane.html'),
                chunks: ['taskpane', 'vendor', 'polyfills']
            }),
            new CopyWebpackPlugin([
                {
                    from: './assets',
                    ignore: ['*.scss'],
                    to: 'assets',
                }
            ]),
            new webpack.ProvidePlugin({
                Promise: ["es6-promise", "Promise"]
            })
        ],
        devServer: {
            contentBase: path.join(__dirname, 'testBuild'),
            headers: {
                "Access-Control-Allow-Origin": "*"
            },
            https: (options.https !== undefined) ? options.https : await devCerts.getHttpsServerOptions(),
            port: process.env.npm_package_config_dev_server_port || 3000
        }
    };

    return config;
};
