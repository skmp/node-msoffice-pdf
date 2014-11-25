var mso_pdf = require('./lib');
var uuid = require('node-uuid');



var okay = 0, errors = 0;

mso_pdf(null, function (error, office) {
    if (error) { 
        console.log("Failed to init");
        return;
    }
    
    for (var i = 0; i< 45; i++) {
        office.word({
            input: "testcases\\test.docx",
            output: "output.doc." + uuid.v4() + ".pdf"
        }, function (error, pdf) {    
            if (error) { 
                console.log("Word: Failed to convert", error);
                errors++;
            }
            else {
                console.log("Converted to: " + pdf);
                okay++;
            }
        });
        
        office.excel({
            input: "testcases\\test.xlsx",
            output: "output.xls." + uuid.v4() + ".pdf"
        }, function (error, pdf) {    
            if (error) {
                console.log("Excel: Failed to convert", error);
                errors++;
            }
            else {
                console.log("Converted to: " + pdf);
                okay++;
            }
        });
        
         office.powerPoint({
            input: "testcases\\test.pptx",
            output: "output.ppt." + uuid.v4() + ".pdf"
        }, function (error, pdf) {    
            if (error) {
                console.log("PowerPoint: Failed to convert", error);
                errors++;
            }
            else {
                console.log("Converted to: " + pdf);
                okay++;
            }
        });
    }
    
    console.log("Over and queued");
    
    office.close(null, function() {
        console.log("Office finished & closed, ", okay, errors, errors*100/(okay+errors));
    });
})
