require('dotenv').config();
const bundleAnalyzer = require('@next/bundle-analyzer');
const transpileModules = require('next-transpile-modules');

const withBundleAnalyzer = bundleAnalyzer({
  enabled: process.env.ANALYZE === 'true'
});

let HOST_URL = 'http://localhost:3000';

if (process.env.NODE_ENV === 'production') {
  HOST_URL = process.env.HOST_URL;
}

const withTM = transpileModules(['pretty-bytes']);

module.exports = withBundleAnalyzer(
  withTM({
    env: {
      STRIPE_CLIENT_TOKEN: process.env.STRIPE_CLIENT_TOKEN,
      HOST_URL
    },
    webpack(config) {
      config.node = {fs: 'empty'};
      return config;
    }
  })
);
