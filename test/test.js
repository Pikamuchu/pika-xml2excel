'use strict';

var sinon = require('sinon');
var assert = require('chai').assert;
var xml2excel = require('../lib');

describe('xml2excel tests', function() {
  describe('Simple conversion case', function() {
    it('XML to Excel conversion using path', function() {
      var result = xml2excel('./test/sfcc-examples/stores.xml');
      assert.isTrue(result.success);
    });
  });
});
