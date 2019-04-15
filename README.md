# officeParser
A Node.js library to parse text out of any office file. 

## Install via npm
*Currently supports docx. Support for pptx and xlsx coming soon

```
npm install officeparser
```

----------

**Example**
```
var officeParser = require('officeparser');

officeParser.parseWord("/path/to/word.docx", function(data){
        // process data
        console.log(data)
})
```

----------

**github**
https://github.com/harshankur/officeParser
