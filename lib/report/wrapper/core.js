/**
 *  OfficeScript.core.js
 *  Bei dem OfficeScript-Object handelt es sich um eine 1-1 Kopie ausgewählter Core-Funktionen von jQuery (http://jquery.com)
 *  Es wurden die Core Funktionen aus jQuery übernommen und das Konzept der Selectoren auf Office Objekte angewand.
 *  Abweichungen gibt es bei den Handels.
 *
 *      @file OfficeScript.core.js
 *      @fileOverview OfficeScript Object
 *      @version 0.2
 **/
/**
* @class OfficeScript
*/
var OfficeScript = (function () {
    // Define a local copy of OfficeScriptJS
    var OfficeScript = function (selector, context) {
        // The OfficeScript object is actually just the init constructor 'enhanced'
        return new OfficeScript.fn.init(selector, context, rootOfficeScript);
    },
    rootOfficeScript,
    // Save a reference to some core methods
    toString = Object.prototype.toString,
    hasOwn = Object.prototype.hasOwnProperty,
    push = Array.prototype.push,
    slice = Array.prototype.slice,
    trim = String.prototype.trim,
    indexOf = Array.prototype.indexOf,
    // A simple way to get ID(#) or class (.)
    // Result = [".myId", ".", "myId"]
    idClassExpr = /^(?:([#\.])([\w\-]*)$)/,
    // get attribute and value
    attrExpr = /^(?:\[(.*)([=])(.*)\]$)/,
    // [[Class]] -> type pairs
    class2type = {};

    OfficeScript.fn = OfficeScript.prototype = {
        constructor: OfficeScript,
        //Init Function
        /**
         * Description
         * @method init
         * @param {} selector
         * @param {} context
         * @param {} rootOfficeScript
         * @return CallExpression
         */
        init: function (selector, context, rootOfficeScript) {
            var match, elem, ret, doc;

            // Handle $(""), $(null), or $(undefined)
            if (!selector) {
                return this;
            }

            if (typeof selector === "string" && !(context || this.context)) {
                this.selector = selector;
                return this;
            }

            // Handle string selectors
            if (typeof selector === "string" && (context || this.context)) {
                //equivalent to: $(context).find(selector)
                this.selector = (selector) ? selector : this.selector;
                this.context = (context) ? context : this.context;

                return this;//OfficeScript(context).find(selector);
            }

            // HANDLE: $(function)
            if (OfficeScript.isFunction(selector)) {
                selector();
                return this;
            }
            if (selector.selector !== undefined) {
                this.selector = selector.selector;
                this.context = selector.context;
            }

            return OfficeScript.makeArray(selector, this);
        },
        // Execute a callback for every element in the matched set.
        // (You can seed the arguments with an array of args, but this is
        // only used internally.)
        /**
         * Description
         * @method each
         * @param {} callback
         * @param {} args
         * @return CallExpression
         */
        each: function (callback, args) {
            return OfficeScript.each(this, callback, args);
        },
        // Start with an empty selector
        selector: "",
        // The default length of a OfficeScript object is 0
        length: 0,
        // The number of elements contained in the matched element set
        /**
         * Description
         * @method count
         * @return MemberExpression
         */
        count: function () {
            return this.length;
        },
        /**
         * Description
         * @method size
         * @return MemberExpression
         */
        size: function () {
            return this.length;
        },

        /**
         * Description
         * @method toArray
         * @return CallExpression
         */
        toArray: function () {
            return slice.call(this, 0);
        },

        // Get the Nth element in the matched element set OR
        // Get the whole matched element set as a clean array
        /**
         * Description
         * @method get
         * @param {} num
         * @return ConditionalExpression
         */
        get: function (num) {
            return num === null ?

              // Return a 'clean' array
              this.toArray() :

              // Return just the object
              (num < 0 ? this[this.length + num] : this[num]);
        },
        // Take an array of elements and push it onto the stack
        // (returning the new matched element set)
        /**
         * Description
         * @method pushStack
         * @param {} elems
         * @param {} name
         * @param {} selector
         * @return ret
         */
        pushStack: function (elems, name, selector) {
            // Build a new OfficeScriptJS matched element set
            var ret = this.constructor();

            if (OfficeScript.isArray(elems)) {
                push.apply(ret, elems);

            } else {
                OfficeScript.merge(ret, elems);
            }

            // Add the old object onto the stack (as a reference)
            ret.prevObject = this;

            ret.context = this.context;

            if (name === "find") {
                ret.selector = this.selector + (this.selector ? " " : "") + selector;
            } else if (name) {
                ret.selector = this.selector + "." + name + "(" + selector + ")";
            }

            // Return the newly-formed element set
            return ret;
        },

        /**
         * Description
         * @method eq
         * @param {} i
         * @return ConditionalExpression
         */
        eq: function (i) {
            i = +i;
            return i === -1 ?
              this.slice(i) :
              this.slice(i, i + 1);
        },

        /**
         * Description
         * @method first
         * @return CallExpression
         */
        first: function () {
            return this.eq(0);
        },

        /**
         * Description
         * @method last
         * @return CallExpression
         */
        last: function () {
            return this.eq(-1);
        },

        /**
         * Description
         * @method slice
         * @return CallExpression
         */
        slice: function () {
            return this.pushStack(slice.apply(this, arguments),
              "slice", slice.call(arguments).join(","));
        },

        /**
         * Description
         * @method map
         * @param {} callback
         * @return CallExpression
         */
        map: function (callback) {
            return this.pushStack(OfficeScript.map(this, function (elem, i) {
                return callback.call(elem, i, elem);
            }));
        },

        /**
         * Description
         * @method end
         * @return LogicalExpression
         */
        end: function () {
            return this.prevObject || this.constructor(null);
        },

        // For internal use only.
        // Behaves like an Array's method, not like a OfficeScript method.
        push: push,
        sort: [].sort,
        splice: [].splice
    };

    // Give the init function the OfficeScript prototype for later instantiation
    OfficeScript.fn.init.prototype = OfficeScript.fn;

    /** Merge the contents of two or more objects together into the first object
     * Kopie von OfficeScript.extend (https://github.com/jquery/jquery/blob/master/src/core.js)
     * OfficeScript.extend( [deep], target, object1 [, objectN] )
     * @param {bool} deep If true, the merge becomes recursive (aka. deep copy).
     * @param {object} target The object to extend. It will receive the new properties.
     * @param {object} object1 An object containing additional properties to merge in.
     * @param {object} objectN Additional objects containing properties to merge in.
     **/
    OfficeScript.extend =
   /**
     * Description
     * @method extend
     * @return target
     */
    OfficeScript.fn.extend = function () {
        var options, name, src, copy, copyIsArray, clone,
          target = arguments[0] || {},
          i = 1,
          length = arguments.length,
          deep = false;

        // Handle a deep copy situation
        if (typeof target === "boolean") {
            deep = target;
            target = arguments[1] || {};
            // skip the boolean and the target
            i = 2;
        }

        // Handle case when target is a string or something (possible in deep copy)
        if (typeof target !== "object" && !OfficeScript.isFunction(target)) {
            target = {};
        }

        // extend OfficeScript itself if only one argument is passed
        if (length === i) {
            target = this;
            --i;
        }

        for (; i < length; i++) {
            // Only deal with non-null/undefined values
            if ((options = arguments[i]) !== null) {
                // Extend the base object
                for (name in options) {
                    src = target[name];
                    copy = options[name];

                    // Prevent never-ending loop
                    if (target === copy) {
                        continue;
                    }

                    // Recurse if we're merging plain objects or arrays
                    if (deep && copy && (OfficeScript.isPlainObject(copy) || (copyIsArray = OfficeScript.isArray(copy)))) {
                        if (copyIsArray) {
                            copyIsArray = false;
                            clone = src && OfficeScript.isArray(src) ? src : [];

                        } else {
                            clone = src && OfficeScript.isPlainObject(src) ? src : {};
                        }

                        // Never move original objects, clone them
                        target[name] = OfficeScript.extend(deep, clone, copy);

                        // Don't bring in undefined values
                    } else if (copy !== undefined) {
                        target[name] = copy;
                    }
                }
            }
        }

        // Return the modified object
        return target;
    };

    /**
     * .each() Methode aus https://github.com/jquery/jquery/blob/master/src/core.js
     */
    OfficeScript.extend({
        // args is for internal usage only
        /**
         * Description
         * @method each
         * @param {} object
         * @param {} callback
         * @param {} args
         * @return object
         */
        each: function (object, callback, args) {
            var name, i = 0,
            length = object.length,
            isObj = length === undefined || OfficeScript.isFunction(object);

            if (args) {
                if (isObj) {
                    for (name in object) {
                        if (callback.apply(object[name], args) === false) {
                            break;
                        }
                    }
                } else {
                    for (; i < length;) {
                        if (callback.apply(object[i++], args) === false) {
                            break;
                        }
                    }
                }

                // A special, fast, case for the most common use of each
            } else {
                if (isObj) {
                    for (name in object) {
                        if (callback.call(object[name], name, object[name]) === false) {
                            break;
                        }
                    }
                } else {
                    for (; i < length;) {
                        if (callback.call(object[i], i, object[i++]) === false) {
                            break;
                        }
                    }
                }
            }

            return object;
        }
    });

    /**
     * Array-methoden aus https://github.com/jquery/jquery/blob/master/src/core.js
     */
    OfficeScript.extend({
        // results is for internal usage only
        /**
         * Description
         * @method makeArray
         * @param {} array
         * @param {} results
         * @return ret
         */
        makeArray: function (array, results) {
            var ret = results || [];

            if (array !== null) {
                // The window, strings (and functions) also have 'length'
                // Tweaked logic slightly to handle Blackberry 4.7 RegExp issues #6930
                var type = OfficeScript.type(array);
                if (array.length === null || type === "string" || type === "function" || type === "regexp") {
                    push.call(ret, array);
                } else {
                    OfficeScript.merge(ret, array);
                }
            }

            return ret;
        },

        /**
         * Description
         * @method inArray
         * @param {} elem
         * @param {} array
         * @param {} i
         * @return UnaryExpression
         */
        inArray: function (elem, array, i) {
            var len;

            if (array) {
                if (indexOf) {
                    return indexOf.call(array, elem, i);
                }

                len = array.length;
                i = i ? i < 0 ? Math.max(0, len + i) : i : 0;

                for (; i < len; i++) {
                    // Skip accessing in sparse arrays
                    if (i in array && array[i] === elem) {
                        return i;
                    }
                }
            }

            return -1;
        },

        /**
         * Description
         * @method merge
         * @param {} first
         * @param {} second
         * @return first
         */
        merge: function (first, second) {
            var i = first.length,
            j = 0;

            if (typeof second.length === "number") {
                for (var l = second.length; j < l; j++) {
                    first[i++] = second[j];
                }

            } else {
                while (second[j] !== undefined) {
                    first[i++] = second[j++];
                }
            }

            first.length = i;

            return first;
        },

        /**
         * Description
         * @method grep
         * @param {} elems
         * @param {} callback
         * @param {} inv
         * @return ret
         */
        grep: function (elems, callback, inv) {
            var ret = [], retVal;
            inv = !!inv;

            // Go through the array, only saving the items
            // that pass the validator function
            for (var i = 0, length = elems.length; i < length; i++) {
                retVal = !!callback(elems[i], i);
                if (inv !== retVal) {
                    ret.push(elems[i]);
                }
            }

            return ret;
        }
    });


    /**
     * Typen-Prüfmethoden aus https://github.com/jquery/jquery/blob/master/src/core.js
     */
    OfficeScript.extend({
        /**
         * Description
         * @method isFunction
         * @param {} obj
         * @return BinaryExpression
         */
        isFunction: function (obj) {
            return OfficeScript.type(obj) === "function";
        },

        isArray: Array.isArray || function (obj) {
            return OfficeScript.type(obj) === "array";
        },

        /**
         * Description
         * @method isNumeric
         * @param {} obj
         * @return LogicalExpression
         */
        isNumeric: function (obj) {
            return !isNaN(parseFloat(obj)) && isFinite(obj);
        },

        /**
         * Description
         * @method type
         * @param {} obj
         * @return ConditionalExpression
         */
        type: function (obj) {

            return obj === null ?
            String(obj) :
            class2type[toString.call(obj)] || "object";
        },

        /**
         * Description
         * @method isPlainObject
         * @param {} obj
         * @return LogicalExpression
         */
        isPlainObject: function (obj) {
            if (!obj || OfficeScript.type(obj) !== "object") {
                return false;
            }

            try {
                // Not own constructor property must be Object
                if (obj.constructor &&
                  !hasOwn.call(obj, "constructor") &&
                  !hasOwn.call(obj.constructor.prototype, "isPrototypeOf")) {
                    return false;
                }
            } catch (e) {
                return false;
            }

            // Own properties are enumerated firstly, so to speed up,
            // if last one is own, then all properties are own.

            var key;
            for (key in obj) { }

            return key === undefined || hasOwn.call(obj, key);
        },

        /**
         * Description
         * @method isEmptyObject
         * @param {} obj
         * @return Literal
         */
        isEmptyObject: function (obj) {
            for (var name in obj) {
                return false;
            }
            return true;
        },

        /**
         * Description
         * @method error
         * @param {} msg
         * @return 
         */
        error: function (msg) {
            throw new Error(msg);
        }
    });

    /**
     * find() Methoden sucht den selector im aktuellen kontext (Achtung: diese funktion ist anders als in jquery, da kontext nicht document sonderen JS-Objeke).
     */
    OfficeScript.fn.extend({
        // results is for internal usage only
        /**
         * Description
         * @method find
         * @param {} selector
         * @param {} returnPlainObject
         * @return results
         */
        find: function (selector, returnPlainObject) {
            var self = this,
              i, e, element, match,
              attr,
              value,
              results = [],
              tempResult
            ;
            returnPlainObject = (returnPlainObject !== undefined) ? returnPlainObject : false;
            // HANDEL $("[name=attrValue]" ,context)
            //(which is just equivalent to: $(context).find(selector)
            if (selector.charAt(0) === "[" && selector.charAt(selector.length - 1) === "]" && selector.length >= 5 && selector.indexOf("=") != -1) {
                match = attrExpr.exec(selector);
                if (match) {
                    for (i = 0; i < self.length; i++) {
                        element = self[i];

                        //
                        if (hasOwn.call(element, match[1])) {
                            if (element[match[1]] == match[3]) {
                                push.call(results, element);
                            }
                        }

                        //Loop durch alle Member
                        for (e in element) {
                            //debug(e +" : " + OfficeScript.type( element[e]) );
                            if (OfficeScript.type(element[e]) === "object") {
                                tempResult = OfficeScript(element[e]).find(selector, true);
                                if (tempResult.length > 0) {
                                    OfficeScript.merge(results, tempResult);
                                }

                            }
                        }
                    }
                }
            } else {

                match = idClassExpr.exec(selector);
                if (match) {
                    // HANDLE: $("#id")
                    if (match[1] === "#") {
                        return this.findElementById(match[2]);
                    }
                    // HANDLE: $(".class")
                    if (match[1] === ".") {
                        return this.findElementsByClass(match[2]);
                    }
                }
            }

            results = (returnPlainObject) ? results : OfficeScript(results);
            return results;
        },
        /**
         * Description
         * @method findElementById
         * @param {} idValue
         * @return 
         */
        findElementById: function (idValue) {
            var self = this,
            attr = "_id",
            i;
            for (i = 0; i < self.length; i++) {
                if (hasOwn.call(self[i], attr)) {
                    if (self[i][attr] == idValue) {
                        return OfficeScript(self[i]); //Return first object
                    }
                }
            }
        },
        /**
         * Description
         * @method findElementsByClass
         * @param {} classValue
         * @return CallExpression
         */
        findElementsByClass: function (classValue) {
            var self = this,
            attr = "_class",
            results = [],
            i, element, classValues;

            for (i = 0; i < self.length; i++) {
                element = self[i];
                if (hasOwn.call(element, attr)) {
                    classValues = element[attr].split(" ");

                    if (OfficeScript.inArray(classValue, classValues) != -1) {
                        push.call(results, element);
                    }
                }
            }
            return OfficeScript(results);
        }

    });

    //getter/setter for attributes
    OfficeScript.extend({
        /**
         * Description
         * @method attr
         * @param {} name
         * @param {} value
         * @param {} parent
         * @param {} targetName
         * @return 
         */
        attr: function (name, value, parent, targetName) {
            targetName = targetName || 'target';
            parent = parent || this;
            parent.target = parent[targetName];
            if (typeof name === 'undefined') return parent;
            if (typeof value !== 'undefined' && value !== null) {
                if (parent.target) {
                    OfficeScript(parent.target).each(function () {
                        this[name] = value;
                    });
                } else {
                    parent[name] = value;
                }
                return parent;
            } else {
                if (parent.target[0]) {
                    return parent.target[0][name];
                } else {
                    return parent[name];
                }
            }
        }
    });

    // Populate the class2type map
    OfficeScript.each("Boolean Number String Function Array Date RegExp Object".split(" "), function (i, name) {
        class2type["[object " + name + "]"] = name.toLowerCase();
    });

    //point back to these default object
    rootOfficeScript = OfficeScript();

    return OfficeScript;
})();

module.exports = OfficeScript;
