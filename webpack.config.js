/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
// Production builds are served from GitHub Pages; the manifest's localhost URLs are rewritten to this.
const urlProd = "https://g34n76r.github.io/SankeyFlow/";

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    // No polyfill bundle: the targets in package.json's "browserslist" (the
    // Office webviews: WebView2 on Windows, Safari 14+ on Mac, and browsers for
    // Excel on the web) support everything the code uses.
    entry: {
      taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane.html"],
      preview: "./src/preview/preview.ts",
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader"
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          // Stylesheets linked from the HTML pages, copied as files. Fixed names:
          // with a content hash in the name, production builds renamed them after
          // minifying while the pages still linked the old names (no styles at all).
          test: /\.css$/,
          type: "asset/resource",
          generator: {
            filename: "[name][ext]",
          },
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
      // The pop-out chart window.
      new HtmlWebpackPlugin({
        filename: "preview.html",
        template: "./src/preview/preview.html",
        chunks: ["preview"],
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};