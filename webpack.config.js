/* eslint-disable @typescript-eslint/no-require-imports */
const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

const DEV_URL = "https://localhost:3000/";
// Production hosting URL (must be HTTPS and end with "/"), e.g.  ADDIN_URL=https://me.github.io/mathpaste/ npm run build
const PROD_URL = process.env.ADDIN_URL || DEV_URL;
// The hosted (production) add-in uses its own Id so Word never confuses it with the
// localhost development registration (Word caches manifests per add-in Id).
const DEV_ID = "54c02531-8874-4ffa-8ce8-d3a412be9e39";
const PROD_ID = "3079ff9f-169c-487b-a1fa-b62115fa9f1e";

async function httpsOptions() {
  const devCerts = require("office-addin-dev-certs");
  return devCerts.getHttpsServerOptions();
}

module.exports = async (env, argv) => {
  const dev = argv.mode !== "production";
  const url = dev ? DEV_URL : PROD_URL;

  return {
    devtool: dev ? "source-map" : false,
    entry: { taskpane: "./src/taskpane/index.tsx" },
    output: { path: path.resolve(__dirname, "dist"), filename: "[name].js", clean: true },
    resolve: { extensions: [".ts", ".tsx", ".js"] },
    module: {
      rules: [
        { test: /\.tsx?$/, use: { loader: "ts-loader", options: { transpileOnly: true } }, exclude: /node_modules/ },
        { test: /\.css$/, use: ["style-loader", "css-loader"] },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({ filename: "taskpane.html", template: "./src/taskpane/taskpane.html", chunks: ["taskpane"] }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets/*.png", to: "assets/[name][ext]" },
          { from: "src/shortcuts.json", to: "shortcuts.json" },
          {
            from: "manifest.xml",
            to: "manifest.xml",
            transform: (content) =>
              url === DEV_URL
                ? content
                : content
                    .toString()
                    .replace(`<Id>${DEV_ID}</Id>`, `<Id>${PROD_ID}</Id>`)
                    .replace(new RegExp(DEV_URL, "g"), url)
                    // <AppDomain> must be an origin (scheme + host)
                    .replace(/https:\/\/localhost:3000(?=<)/g, new URL(url).origin),
          },
        ],
      }),
    ],
    performance: { hints: false },
    devServer: {
      static: { directory: path.join(__dirname, "dist") },
      headers: { "Access-Control-Allow-Origin": "*" },
      server: { type: "https", options: dev ? await httpsOptions() : {} },
      port: 3000,
      hot: false,
      liveReload: true,
      client: { overlay: { warnings: false, errors: true } },
    },
  };
};
