/*global describe,it,after*/
var assert = require('assert')
, path = require('path')
, reportApp = require('../').report.application
, testPPT01 = 'Testpptx_01.pptx'
, testDataPath = path.join(__dirname, 'data')
;


describe('report', function(){
    this.timeout(15000);
    reportApp(null, function(err, app) {
        after( function(done) {app.quit(null, done);} );
        describe('presentation', function(){
            describe('#open&close', function(){
                it('should open and close the file', function(done){
                    app.open( path.join(testDataPath,testPPT01), function(err, presentation) {
                        if(err) throw err;
                        presentation.close(null, done);
                    });
                });
            });
            describe('#attr', function(){
                it('should get a name and path attribute from presentation', function(done){
                    app.open( path.join(testDataPath,testPPT01), function(err, presentation) {
                        if(err) throw err;
                        //get Path Sync
                        assert.equal(presentation.attr({name:'Path'}, true), testDataPath);
                        //get name async
                        presentation.attr({name:'Name'}, function(err, data) {
                            if(err) throw err;
                            assert.equal(data, testPPT01);
                            presentation.close(null, done);
                         });
                    });
                });
            });
            describe('#slides', function(){
                it('should have 2 slides', function(done){
                    app.open( path.join(testDataPath,testPPT01), function(err, presentation) {
                        if(err) throw err;
                        presentation.slides(null, function(err, slides) {
                            if(err) throw err;
                            assert.equal(slides.length, 2);
                            presentation.close(null, done);
                        });
                    });
                });
                
                it('should have the Attr Name', function(done){
                    app.open( path.join(testDataPath,testPPT01), function(err, presentation) {
                        if(err) throw err;
                        presentation.slides(null, function(err, slides) {
                            if(err) throw err;
                            assert.equal(slides[1].attr({name:'Name'}, true), 'Slide2');
                            presentation.close(null, done);
                        });
                    });
                });
                
                it('should have the Attr Pos', function(done){
                    app.open( path.join(testDataPath,testPPT01), function(err, presentation) {
                        if(err) throw err;
                        presentation.slides(null, function(err, slides) {
                            if(err) throw err;
                            slides.forEach(function(slide, index) {
                                assert.equal(slide.attr({name:'Pos'}, true), index+1);
                            })
                            presentation.close(null, done);
                        });
                    });
                });
                it('should be changeable the pos of Slide2', function(done){
                    app.open( path.join(testDataPath,testPPT01), function(err, presentation) {
                        if(err) throw err;
                        presentation.slides(null, function(err, slides) {
                            if(err) throw err;
                            assert.equal(slides[1].attr({name:'Pos'}, true), 2);
                            slides[1].attr({name:'Pos', value:1}, true);
                            assert.equal(slides[1].attr({name:'Pos'}, true), 1);
                            presentation.close(null, done);
                        });
                    });
                });
                it('should be able to delete Slide2', function(done){
                    app.open( path.join(testDataPath,testPPT01), function(err, presentation) {
                        if(err) throw err;
                        presentation.slides(null,true)[1].remove(null,function(err) {
                            presentation.slides(null, function(err, slides) {
                                if(err) throw err;
                                assert.equal(slides.length, 1);
                                presentation.close(null, done);
                            });
                        });
                    });
                });
                it('should be able to create a shape on slide1', function(done){
                    app.open( path.join(testDataPath,testPPT01), function(err, presentation) {
                        if(err) throw err;
                        var slide = presentation.slides(null,true)[1];
                        var shapeCount = slide.shapes(null, true).length;
                        slide.addTextbox(null, function(err, shape) {
                            if(err) throw err;
                            assert.equal(slide.shapes(null, true).length, shapeCount + 1);
                            presentation.close(null, done);  
                        });
                    });
                });
            });
            describe('#shapes', function(){
                it('should have 2 shapes on slide one', function(done){
                    app.open( path.join(testDataPath,testPPT01), function(err, presentation) {
                        if(err) throw err;
                        presentation.slides(null, function(err, slides) {
                            if(err) throw err;
                            slides[0].shapes(null, function(err, shapes) {
                                assert.equal(shapes.length, 2);
                                presentation.close(null, done);
                            });     
                        });
                    });
                });
                it('should have the Attr Name', function(done){
                    app.open( path.join(testDataPath,testPPT01), function(err, presentation) {
                        if(err) throw err;
                        presentation.slides(null, function(err, slides) {
                            if(err) throw err;
                            slides[0].shapes(null, function(err, shapes) {
                                assert.equal(shapes[0].attr({name:'Name'} ,true), 'Title 1');
                                presentation.close(null, done);
                            });     
                        });
                    });
                });
                it('should be changeable the Attribute Name', function(done){
                    app.open( path.join(testDataPath,testPPT01), function(err, presentation) {
                        if(err) throw err;
                        presentation.slides(null, function(err, slides) {
                            if(err) throw err;
                            slides[0].shapes(null, function(err, shapes) {
                                assert.equal(shapes[0].attr({name:'Name', value:'Test'} ,true).attr({name:'Name'}, true), 'Test');
                                presentation.close(null, done);
                            });     
                        });
                    });
                });
            });
        });
       
    });
});