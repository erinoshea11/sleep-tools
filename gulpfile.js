'use strict';

const build = require('@microsoft/sp-build-web');
build.addSuppressRule(/warn/);
build.initialize(require('gulp'));
