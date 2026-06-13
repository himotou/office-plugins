/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");

const urlDev = "https://localhost:3000/";

function normalizeBaseUrl(rawUrl) {
  const normalizedUrl = rawUrl.endsWith("/") ? rawUrl : `${rawUrl}/`;
  return new URL(normalizedUrl).toString();
}

function resolveProductionBaseUrl(env) {
  const rawBaseUrl = env?.addinBaseUrl || env?.ADDIN_BASE_URL || process.env.ADDIN_BASE_URL;

  if (!rawBaseUrl) {
    throw new Error(
      "Missing ADDIN_BASE_URL for production build. Example: ADDIN_BASE_URL=https://<user>.github.io/<repo> npm run build"
    );
  }

  return normalizeBaseUrl(rawBaseUrl);
}

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const productionBaseUrl = dev ? urlDev : resolveProductionBaseUrl(env);
  const productionSiteUrl = productionBaseUrl.endsWith("/") ? productionBaseUrl.slice(0, -1) : productionBaseUrl;
  const productionOrigin = new URL(productionBaseUrl).origin;

  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      react: ["react", "react-dom"],
      taskpane: {
        import: ["./src/taskpane/index.tsx", "./src/taskpane/taskpane.html"],
        dependOn: "react",
      },
      dialog: {
        import: ["./src/dialog/index.tsx", "./src/dialog/dialog.html"],
        dependOn: "react",
      },
      resourcePicker: {
        import: ["./src/resource-picker/index.tsx", "./src/resource-picker/resource-picker.html"],
        dependOn: "react",
      },
      commands: "./src/commands/commands.ts",
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
          },
        },
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: ["ts-loader"],
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|ttf|woff|woff2|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "index.html",
        inject: false,
        templateContent: `
          <!DOCTYPE html>
          <html lang="en">
            <head>
              <meta charset="UTF-8" />
              <meta http-equiv="refresh" content="0; url=taskpane.html" />
              <title>link-bind</title>
            </head>
            <body>
              <p>Redirecting to <a href="taskpane.html">taskpane.html</a>...</p>
            </body>
          </html>
        `,
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "react"],
      }),
      new HtmlWebpackPlugin({
        filename: "dialog.html",
        template: "./src/dialog/dialog.html",
        chunks: ["polyfill", "dialog", "react"],
      }),
      new HtmlWebpackPlugin({
        filename: "resource-picker.html",
        template: "./src/resource-picker/resource-picker.html",
        chunks: ["polyfill", "resourcePicker", "react"],
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
              }

              return content
                .toString()
                .replace(/<AppDomain>https:\/\/localhost:3000<\/AppDomain>/g, `<AppDomain>${productionOrigin}</AppDomain>`)
                .replace(/https:\/\/localhost:3000/g, productionSiteUrl);
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"],
      }),
    ],
    devServer: {
      hot: true,
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
