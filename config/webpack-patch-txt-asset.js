'use strict';

/**
 * Adds a webpack module rule so that importing a .txt file resolves to a
 * URL string pointing at the deployed asset, the same way SPFx already
 * handles .png/.svg/etc via the built-in `asset/resource` type.
 *
 * Why this exists: the summarization feature needs to ship pdf.js's worker
 * script as a plain-text asset (see summarization/parsers/pdfWorkerLoader.ts
 * for the full explanation of why it's not imported as a normal JS module).
 * SPFx's default webpack config has no rule for .txt files at all, so without
 * this patch the build fails with "Module not found" the moment anything
 * actually imports that file.
 *
 * This uses the SPFx Heft Webpack Patch Plugin (config/webpack-patch.json)
 * rather than `heft eject-webpack`, which would permanently fork the whole
 * build config out of SPFx's managed toolchain - overkill for one rule.
 */
module.exports = function (webpackConfig) {
  webpackConfig.module = webpackConfig.module || {};
  webpackConfig.module.rules = webpackConfig.module.rules || [];

  webpackConfig.module.rules.push({
    test: /\.txt$/,
    type: 'asset/resource',
  });

  return webpackConfig;
};
