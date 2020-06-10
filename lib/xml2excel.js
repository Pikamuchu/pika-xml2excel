'use strict';

var fs = require('fs');
var isValidPath = require('is-valid-path');
var XML = require("simple4x");
var json2xls = require('json2xls');

var XML_REGEX = /^\s*</;

/**
 * Main xml to excel conversion function.
 * @param {string} src - xml/excel source data/path to convert.
 * @param {object} opts - conversion options.
 */
function xml2excel(src, opts) {
  var result;
  var options = parseOptions(src, opts || {});
  var data = parseData(src, options);
  if (options.isXML) {
    console.log('Processing XML source data...');
    var xmlObj = xml2json(data, options);

    console.log('  Generating excel data ...');
    var xls = json2excel(xmlObj, options);

    if (options.outPath) {
      console.log('  Writing to file ' + options.outPath + ' ...');
      fs.writeFileSync(options.outPath, xls, 'binary');
      result = {
        success: true,
        outPath: options.outPath
      };

    } else {
      result = xls;
    }

  } else if (options.isExcel) {
    console.log('Processing Excel source data...');
  } else {
    console.log('Unknown source data.');
  }

  function xml2json(data, options) {
    return new XML(data);
  }

  function json2excel(xmlObj, options) {
    var dataElements = xmlObj && xmlObj.elements()[0];
    var excelData = parseDataElements(dataElements);
    return json2xls(excelData);
  }

  function json2xml(obj, options) {
    var builder = new xml2js.Builder();
    return builder.buildObject(obj);
  }

  function excel2json(excel, options) {
    // todo
  }

  console.log('Done.');
  return result;
}

function parseOptions(src, opts) {
  var options = {
    path: opts.path,
    isXML: opts.isXML,
    isExcel: opts.isExcel
  };
  return options;
}

function parseData(src, options) {
  var data;
  var path = options.path || (typeof src === 'string' && isValidPath(src) ? src : null);
  if (path) {
    data = fs.readFileSync(path).toString();
    // Updating missing options
    if (!options.path) {
      options.path = path;
    }
    if (options.isXML === undefined) {
      options.isXML = XML_REGEX.test(data);
    }
    if (!options.outPath) {
      if (options.isXML) {
        options.outPath = path.replace(/\.xml$/, '.xlsx');
      }
      if (options.isExcel) {
        options.outPath = path.replace(/\.xlsx$/, '.xml');
      }
    }
  } else {
    data = src;
    // Updating missing options
    if (opts.isXML === undefined) {
      options.isXML = typeof data === 'string' && XML_REGEX.test(data);
    }
  }
  return data;
}

function parseDataElements(elements) {
  var result = [];
  for (var i = 0; i < elements.length(); i++) {
    result.push(parseDataElement(elements[i]));
  }
  return result;
}

function parseDataElement(element, obj, name) {
  var object = obj || {};
  var prefix = name ? name + '>' : '';
  if (element.__class === 'XMLNode') {
    var keys = element.attributesNames();
    keys.forEach(function (key) {
      var attrKey = prefix + '@' + key;
      var attrValue = element.attribute(key);
      object[attrKey] = attrValue && attrValue.toString();
    });
    keys = element.elementsNames();
    if (keys.length === 1 && keys[0] === '_') {
      object[name || '_'] = element.text();
    } else {
      keys.forEach(function (key) {
        parseDataElement(element[key], object, prefix + key);
      });
    }

  } else if (element.__class === 'XMLList') {
    for (var i = 0; i < element.length(); i++) {
      parseDataElement(element[i], object, (name || '') + '(' + i + ')');
    }

  } else {
    object[name || '_'] = '' + element;
  }
  return object;
}

module.exports = xml2excel;
