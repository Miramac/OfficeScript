/**
 *  OfficeScript.Report.Range.js
 *
 *  Range Core: Selector and Basic Shape-Functions
 **/
var $ = require('./OfficeScript.Core')
, DataNET = require('./OfficeScript.Data.NET')
, Attributes = require('./OfficeScript.Data.Range.Attributes')
, _ = require('../lib/lodash')
;

/**
* @class $range
* @param {} selector
* @param {} context
* @return NewExpression
*/
var Range = (function () {
    var Range = function (selector, context) {
        if (selector instanceof Range) {
            return selector
        } else {
            return new Range.fn.init(selector, context);
        }
    };
    Range.fn = Range.prototype = {
        constructor: Range,
        /**
          * Description
          * @method init
          * @param {} selector
          * @param {} context
          * @return ThisExpression
          */
        init: function (selector, context) {
            // Initialize Shapeobject, operates as Selector.
            // "selector": Either String-Array or String (comma-separated) ID-Value
            // "context": Limit search to Sliderange
            var i, j, tmpRange;
            this.range = [];
            this.sheets = [];
            context = context || [];
            if (!selector) {
                return this;
            }
            if (typeof context === 'string') {
                context = context; //TODO Sheets(context);
            }
            if (typeof selector === 'string') {
                selector = selector.split(',');
                return Range(selector, context);
            }
            if (selector.toString() === 'OfficeScript.DataScript.Range') {
                selector = [selector];
            }
            if ($.isArray(selector)) {
                for (i = 0; i < selector.length; i++) {
                    if ((typeof selector[i] === 'string')) {
                        selector.splice(i, 1, DataNET.findRange(selector[i].trim(), context));
                    } else if (selector[i] instanceof Range) {
                        tmpRange = [];
                        for (j = 0; j < selector[i].range.length; j++) {
                            tmpRange.push(selector[i].range[j]);
                        }
                        selector.splice(i, 1, tmpRange);
                    }
                }
                selector = _.flatten(selector, true);
                selector = _.compact(selector);
            }

            this.range = selector;
            this.sheets = context;
            return this;
        },
        /**
          * Number of PowerPoint.Range linked in this Object.
          * @method count
          * @return MemberExpression
          */
        count: function () {
            return this.range.length;
        },
        /**
          * Iterate over all linked PowerPoint.Range using:
          * @method each
          * @param {Function} callback
          * @param {object} args (only for internal use!)
          * @return CallExpression
          */
        each: function (callback, args) {
            return $(this.range).each(callback, args);
        },
        /**
          * General Attribute getter/setter.
          * @method attr
          * @param {String} name
          * @param {String|Number} value
          * @param {Object} parent
          * @param {String} targetName
          * @return targetName {String} ['range'] Define property e.g. 'paragraphs', 'range'
          */
        attr: function (name, value, parent, targetName) {
            targetName = targetName || 'range';
            parent = parent || this;
            return $.attr(name, value, parent, targetName);
        },
        /**
          * Destroy Object
          * @method dispose
          * @return ThisExpression
          */
        dispose: function () {
            this.each(function () {
                this.Dispose();
            });
            this.range = [];
            return this;
        }
    };
    Range.fn.init.prototype = Range.fn;
    return Range;
}());

$.extend(Range.fn, Attributes);

module.exports = Range;