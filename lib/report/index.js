var application = require('./application');
var app; // = application(null,true)

// function Application() {
    // var app = application(null,true);
    
    // return {
        // application: app,
        // open: app.open,
        // quit: app.quit
    // }
// }

// module.exports = Application;

module.exports = {
    application: application,
    open: function(path, cb) {
            if(!app) {
                app = application(null,true)
            }
            return app.open(path, cb)
        },
    quit: function(param, cb) {
            if(app) {
                var result = app.quit(param, cb)
                app = null;
            }
        } 
}
