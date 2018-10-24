import path from 'path';

import HtmlWebpackPlugin from 'html-webpack-plugin';

module.exports = {
    entry: path.join(__dirname, 'src', 'index.tsx'),
    output: {
        path: path.join(__dirname, 'build'),
        filename: 'index.bundle.js'
    },
    mode: process.env.NODE_ENV || 'development',
    resolve: {
        modules: [path.resolve(__dirname, 'src'), 'node_modules']
    },
    devServer: {
        contentBase: path.join(__dirname, 'src')
    },
    module: {
        rules: [
            {
                test:/\.(ts|tsx)$/,
                exclude:/node_modules/,
                use:['ts-loader']
            },
            {
                test: /\.(js|jsx)$/,
                exclude: /node_modules/,
                use: ['babel-loader']
            },
            {
                test: /\.(css|scss)$/,
                use: [
                    "style-loader",
                    "css-loader",
                    "sass-loader"
                ]
            },
            {
                test: /\.(jpg|jpeg|png|gif|mp3|svg)$/,
                loaders: ['file-loader']
            }
        ]
    },
    resolve: {
      extensions: [ '.tsx', '.ts', '.js' ]
    },
    devServer: {
        port:3000,
        public: 'pjsummersjr.ngrok.io' //Stops 'Invalid host header' errors when using ngrok
    },
    plugins: [
        new HtmlWebpackPlugin({
            template: path.join(__dirname, 'src', 'index.html')
        })
    ]
}