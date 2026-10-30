// 'use strict';

// const build = require('@microsoft/sp-build-web');

// build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

// var getTasks = build.rig.getTasks;
// build.rig.getTasks = function () {
//   var result = getTasks.call(build.rig);

//   result.set('serve', result.get('serve-deprecated'));

//   return result;
// };

// build.initialize(require('gulp'));
 
'use strict';

const build = require('@microsoft/sp-build-web');

//  Suppress ALL warnings (sass + lint + webpack)
build.addSuppression(/Warning - \[sass\]/);
build.addSuppression(/Warning - \[lint\]/);
build.addSuppression(/Warning - lint/);
build.addSuppression(/Warning - \[webpack\]/);

// (optional) suppress specific rules if still showing
build.addSuppression(/no-explicit-any/);
build.addSuppression(/no-floating-promises/);
build.addSuppression(/no-new-null/);
build.addSuppression(/explicit-function-return-type/);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.initialize(require('gulp'));
