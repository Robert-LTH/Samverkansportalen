'use strict';

const gulp = require('gulp');
const nodeSupportRange = '>=12.13.0 <13.0.0 || >=14.15.0 <15.0.0 || >=16.13.0 <17.0.0';

const createSemverOverride = (semverModule) => {
  const original = semverModule.satisfies.bind(semverModule);

  if (original(process.version, nodeSupportRange)) {
    return () => {};
  }

  semverModule.satisfies = (version, range, ...rest) => {
    const normalizedRange = (range || '').replace(/\s+/g, '');

    if (
      version === process.version &&
      normalizedRange.includes('>=12.13.0<13.0.0') &&
      normalizedRange.includes('>=14.15.0<15.0.0') &&
      normalizedRange.includes('>=16.13.0<17.0.0')
    ) {
      return true;
    }

    return original(version, range, ...rest);
  };

  return () => {
    semverModule.satisfies = original;
  };
};

const restoreRootSemver = createSemverOverride(require('semver'));
const restoreBuildSemver = createSemverOverride(require('@microsoft/sp-build-web/node_modules/semver'));

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);
build.tslintCmd.enabled = false;
build.lintCmd.enabled = false;

build.initialize(gulp);

restoreRootSemver();
restoreBuildSemver();
