var reportApp = require('../').report.application
;

reportApp(null, function(err, app) {
     app.open(__dirname+'\\data\\Testpptx_02.pptx', function(err, presentation) {
        //use presentation object
        console.log('Presentation path:', presentation.attr({name:'Path'}, true));
        presentation.slides(null, function(err, slides) {
            if(err) throw err;
            console.log('Slides count:', slides.length);
             console.log('Slides props:', slides);
            slides[1].shapes(null, function(err, shapes) {
                var shape0 = shapes[0];
                var shape1 = shapes[1];
                console.log('Shape count on slide 1:', shapes.length);
                shape0.attr({'name':'Text', 'value': 'Fu Bar'}, true); //Set text value
                console.log('get Text first shape:', shape0.attr({'name':'Text'}, true));
                
                console.log('get Text first shape:', shape0.attr({'name':'Text'}, true));
                
                console.log(slides[1].addTextbox({}, true).attr({name:'Name'},true));
                
                console.log(shape1.paragraph({'start':5}, true).attr({name:'Text', value:"test"}, true));

                // close presentation
                // presentation.close(null, function(err){
                    // if(err) throw err;
                    // app.quit()
                // });
            });
        });
    });
 })   
    
    
    // app.open(__dirname+'\\data\\Testpptx_01.pptx', function(err, presentation) {
        // if(err) throw err;
        // console.log('ppt.name', presentation.attr({name:'Path'}, true ));
        // presentation.slides(null, function(err, slides) {
            // if(err) throw err;
            // console.log('tag[x]', slides[0].tags(null,true).set({name:"x",value:"xx"}, true).get('x',true));
            // slides[1].copy(null, function(err, slide) {
                // if(err) throw err;
                // console.log('slide',  slide.attr({'name':'Pos'}, true));
                // slide.attr({'name':'Pos', value:1}, true);
                // console.log('slide Pos',  slide.attr({'name':'Pos'}, true));
                // console.log(slide.addTextbox(null,true).attr({'name':'Text', 'value': 'Fu Bar'}, true).attr({'name':'Name'},true));
                
            // });
           
        // }); 
       // presentation.close(null, function() {
            // app.open(__dirname+'\\data\\Testpptx_01.pptx', function(err, presentation) {
                // if(err) throw err;
                // console.log('ppt.name', presentation.attr({name:'Path'}, true ));
                // presentation.slides(null, function(err, slides) {
                    // if(err) throw err;
                    // console.log('tag[x]', slides[0].tags(null,true).set({name:"x",value:"xx"}, true).get('x',true));
                    // slides[1].copy(null, function(err, slide) {
                        // if(err) throw err;
                        // console.log('slide',  slide.attr({'name':'Pos'}, true));
                        // slide.attr({'name':'Pos', value:1}, true);
                         // slide.attr({'name':'Pos', value:2}, function(e,d){
                            // console.log("Slide object:", d);
                         // })
                    // });
                    
                     // slides[0].shapes(null, function(err, shapes) { 
                        // shapes.forEach(function(shape) {
                            // var test = shape.attr({name:'Text'} , true)
                        // });
                     // });
                // }); 
                
                  // app.quit(null, function(err, data) {
                    // console.log(data);
                  // })
            // });
       // });
        
    // });
    

// });
