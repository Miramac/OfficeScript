
//Simple require function factory
function req(root) {
    return (function (moduleName) {
        var moduleContent
        , module = { exports: {} }
        , exports = module.exports
        , modulePath
        , newRoot
        ;
        moduleName = moduleName.replace(new RegExp('/', 'g'), '\\');
        modulePath = (moduleName.substr(1,2) === ':\\') ? moduleName :  root + moduleName;
        if (/\.js$/.test(moduleName) !== true) {
            if (fs.existsSync( modulePath  + "\\index.js")) {
                modulePath =  modulePath  + "\\index.js";
            } else {
                modulePath += '.js';
            }
        }
        if(!fs.existsSync(modulePath)) {
            throw new Error("Cannot find module '"+moduleName+ "'");
        }
        
        newRoot = modulePath.split('\\').slice(0, modulePath.split('\\').length - 1).join('\\') + '\\';
        var require = req(newRoot);
        moduleContent = fs.readFileSync(modulePath);
        try {
            eval(moduleContent);
        } catch (e) {
            throw new Error(e);
        }
        return module.exports;
    });

}