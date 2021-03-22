import edge from "edge-js";
import path from "node:path";

export default edge.func({
    source: path.join(path.dirname(module.filename), 'office.cs'),
    references: [ 
        'C:\\Windows\\assembly\\GAC_MSIL\\Microsoft.Office.Interop.Word\\15.0.0.0__71e9bce111e9429c\\Microsoft.Office.Interop.Word.dll',
        'C:\\Windows\\assembly\\GAC_MSIL\\Microsoft.Office.Interop.Excel\\15.0.0.0__71e9bce111e9429c\\Microsoft.Office.Interop.Excel.dll',
        'C:\\Windows\\assembly\\GAC_MSIL\\Microsoft.Office.Interop.PowerPoint\\15.0.0.0__71e9bce111e9429c\\Microsoft.Office.Interop.PowerPoint.dll',
        'C:\\Windows\\assembly\\GAC_MSIL\\Office\\15.0.0.0__71e9bce111e9429c\\Office.dll',
        'C:\\Windows\\assembly\\GAC_MSIL\\Microsoft.Vbe.Interop\\15.0.0.0__71e9bce111e9429c\\Microsoft.Vbe.Interop.dll'
    ],
});
