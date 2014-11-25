var edge = require('edge');
var uuid = require('node-uuid');

var Office = edge.func({
    source: 'office.cs',
    references: [ 
        'C:\\Windows\\assembly\\GAC_MSIL\\Microsoft.Office.Interop.Word\\15.0.0.0__71e9bce111e9429c\\Microsoft.Office.Interop.Word.dll',
        'C:\\Windows\\assembly\\GAC_MSIL\\Microsoft.Office.Interop.Excel\\15.0.0.0__71e9bce111e9429c\\Microsoft.Office.Interop.Excel.dll',
        'C:\\Windows\\assembly\\GAC_MSIL\\Microsoft.Office.Interop.PowerPoint\\15.0.0.0__71e9bce111e9429c\\Microsoft.Office.Interop.PowerPoint.dll',
        'C:\\Windows\\assembly\\GAC_MSIL\\Office\\15.0.0.0__71e9bce111e9429c\\Office.dll',
        'C:\\Windows\\assembly\\GAC_MSIL\\Microsoft.Vbe.Interop\\15.0.0.0__71e9bce111e9429c\\Microsoft.Vbe.Interop.dll'
    ],
});

var okay = 0, errors = 0;

Office(null, function (error, office) {
    if (error) { 
        console.log("Failed to open word");
        throw error;
    }
    
    for (var i = 0; i< 45; i++) {
        office.word({
            input: "tests\\test.docx",
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
            input: "tests\\test.xlsx",
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
            input: "tests\\test.pptx",
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