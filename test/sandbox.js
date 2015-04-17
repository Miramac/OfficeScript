var report = require('../').report

var Shape = require('../lib/report/wrapper/shape');
//, reportApp = report.application


var presentation = report.open(__dirname+'\\data\\Testpptx_02.pptx', true);
var slides = presentation.slides(null, true)
var shapes = slides[1].shapes(null,true);

var $shape = new Shape(shapes[0]);

console.log($shape.attr('Name' , 'Foo'));
console.log($shape.name('bar'));
console.log($shape.attr('Name'));

report.quit();




/*
 report.open(__dirname+'\\data\\Testpptx_02.pptx', function(err, presentation) {
    //use presentation object
    console.log('Presentation Name:', presentation.attr({name:'Name'}, true)); 
    console.log('Presentation Path:', presentation.attr({name:'Path'}, true)); 
    console.log('Presentation FullName:',presentation.attr({'name':'FullName'}, true));
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
             presentation.close(null, function(err){
                if(err) throw err;
                report.quit()
            });
        });
    });
});

    */
    
    