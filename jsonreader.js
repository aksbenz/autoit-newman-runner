const fs = require('fs');

// We Use node enclose to create an exe to be used by autoit
fs.writeFileSync("args.dat", process.argv.join(','));

if (process.argv.length < 4) {
    console.log('Arg1: JSON file to read')
    console.log('Arg2: File to write comma separated values')
    process.exit(1);
}

var filepath = process.argv[2] || "";
var tgtFile = process.argv[3] || "";

var c = JSON.parse(fs.readFileSync(filepath));

var paths = [],
    stack = [],
    base = c.item;

var idx = base.length;
// Push to stack all 1st level folders in reverse order
base.reverse().forEach(function(i) {
    i.name = 'O' + idx-- + '_' + i.name;
    stack.push(i)
});

idx++;
var parent;
while (stack.length > 0) {
    var fld = stack.pop();
    if (fld.item) {
        console.log(fld.name);
        paths.push(getParentTree(fld));
        fld.item.reverse().forEach(function(i) {
            i.parent = fld;
            if (i.item) {
                i.name = 'O' + idx++ + '_' + i.name;
                stack.push(i);
            }
        });
    }
}

console.log('*****************');
console.log(paths.join('\n'));
fs.writeFileSync(tgtFile, paths.join('\r\n'));

function getParentTree(fld) {
    var path = fld.name.replace('/', '0x2f');
    while (fld.parent) {
        fld = fld.parent;
        path = fld.name.replace('/', '0x2f') + '/' + path;
    }
    return path;
}