var fs = require('fs');
var bitmap = fs.readFileSync('blank.xlsx');
    // convert binary data to base64 encoded string
var str = Buffer(bitmap).toString('base64');
console.log(str);
fs.writeFileSync('blankFile.js', str);