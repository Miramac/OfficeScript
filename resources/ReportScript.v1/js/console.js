//https://jint.codeplex.com/discussions/263946
var console = {};
console.log = function (args) {
    var logStr = '', i;
    for (i = 0; i < arguments.length; i++) {
        if (typeof arguments[i] === 'object' || typeof arguments[i] === 'object') {
            logStr += JSON.stringify(arguments[i]) + ' ';
        } else {
            logStr += arguments[i] + ' ';
        }
       
    }
    __log(logStr);
};

console.error = function (args) {
    var logStr = '', i;
    for (i = 0; i < arguments.length; i++) {
        if (typeof arguments[i] === 'object' || typeof arguments[i] === 'object') {
            logStr += JSON.stringify(arguments[i]) + ' ';
        } else {
            logStr += arguments[i] + ' ';
        }

    }
    __error(logStr);
};


//console.log('Hello World', 3,  [1,2,"3"], {a:"1"}, false, null);