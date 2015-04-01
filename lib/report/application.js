var edge = require('edge');

var application = edge.func({
    assemblyFile: __dirname+'/../../src/OfficeScript/OfficeScript/bin/Debug/OfficeScript.dll',
    typeName: 'OfficeScript.Startup',
    methodName: 'PowerPointApplication'
});

module.exports = application
