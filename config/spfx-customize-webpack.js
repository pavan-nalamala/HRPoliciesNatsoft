'use strict';

module.exports = function customizeWebpack(webpackConfiguration) {
  webpackConfiguration.module.rules.push({
    dependency: {
      not: ['url']
    },
    generator: {
      filename: '[name]_[contenthash][ext]'
    },
    test: /\.pdf$/i,
    type: 'asset/resource'
  });
};
