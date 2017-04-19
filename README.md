# npmdoc-officegen

#### api documentation for  [officegen (v0.4.4)](https://github.com/Ziv-Barber/officegen#readme)  [![npm package](https://img.shields.io/npm/v/npmdoc-officegen.svg?style=flat-square)](https://www.npmjs.org/package/npmdoc-officegen) [![travis-ci.org build-status](https://api.travis-ci.org/npmdoc/node-npmdoc-officegen.svg)](https://travis-ci.org/npmdoc/node-npmdoc-officegen)

#### Office Open XML Generator using Node.js streams. Supporting Microsoft Office 2007 and later Word (docx), PowerPoint (pptx,ppsx) and Excel (xlsx). This module is for all frameworks and environments. No need for any commandline tool - this module is doing e

[![NPM](https://nodei.co/npm/officegen.png?downloads=true&downloadRank=true&stars=true)](https://www.npmjs.com/package/officegen)

- [https://npmdoc.github.io/node-npmdoc-officegen/build/apidoc.html](https://npmdoc.github.io/node-npmdoc-officegen/build/apidoc.html)

[![apidoc](https://npmdoc.github.io/node-npmdoc-officegen/build/screenCapture.buildCi.browser.%252Ftmp%252Fbuild%252Fapidoc.html.png)](https://npmdoc.github.io/node-npmdoc-officegen/build/apidoc.html)

![npmPackageListing](https://npmdoc.github.io/node-npmdoc-officegen/build/screenCapture.npmPackageListing.svg)

![npmPackageDependencyTree](https://npmdoc.github.io/node-npmdoc-officegen/build/screenCapture.npmPackageDependencyTree.svg)



# package.json

```json

{
    "author": {
        "name": "Ziv Barber",
        "url": "https://github.com/Ziv-Barber"
    },
    "bugs": {
        "url": "https://github.com/Ziv-Barber/officegen/issues"
    },
    "dependencies": {
        "archiver": "~1.3.0",
        "async": "^2.0.1",
        "fast-image-size": "^0.1.3",
        "jszip": "^2.5.0",
        "lodash": "^4.16.2",
        "readable-stream": "~2.2.2",
        "setimmediate": ">= 1.0.1",
        "xmlbuilder": "^3.0.0"
    },
    "description": "Office Open XML Generator using Node.js streams. Supporting Microsoft Office 2007 and later Word (docx), PowerPoint (pptx,ppsx) and Excel (xlsx). This module is for all frameworks and environments. No need for any commandline tool - this module is doing e",
    "devDependencies": {
        "adm-zip": "^0.4.7",
        "chai": "^3.5.0",
        "grunt": "^1.0.1",
        "grunt-contrib-jshint": "^1.0.0",
        "grunt-jsdoc": "^2.1.0",
        "jaguarjs-jsdoc": "^1.0.1",
        "mocha": "^3.0.2",
        "sinon": "^1.17.5",
        "time-grunt": "^1.4.0"
    },
    "directories": {
        "lib": "lib",
        "examples": "examples"
    },
    "dist": {
        "shasum": "c4009913e6374ac58aaf0e4b3ff99f50502581b4",
        "tarball": "https://registry.npmjs.org/officegen/-/officegen-0.4.4.tgz"
    },
    "engines": {
        "node": ">= 0.10.0"
    },
    "gitHead": "539bce3e971f3c80341cd29a6994dab5459fc838",
    "homepage": "https://github.com/Ziv-Barber/officegen#readme",
    "keywords": [
        "officegen",
        "office",
        "microsoft",
        "2007",
        "word",
        "powerpoint",
        "excel",
        "docx",
        "pptx",
        "ppsx",
        "xlsx",
        "charts",
        "Office Open XML",
        "stream",
        "generate",
        "create",
        "maker",
        "generator",
        "creator"
    ],
    "license": "MIT",
    "main": "index.js",
    "maintainers": [
        {
            "name": "zivbarber"
        }
    ],
    "name": "officegen",
    "optionalDependencies": {},
    "repository": {
        "type": "git",
        "url": "git://github.com/Ziv-Barber/officegen.git"
    },
    "scripts": {
        "test": "mocha test/test-docx.js test/test-xlsx.js test/test-pptx-nocharts.js test/test-pptx-charts.js"
    },
    "url": "https://github.com/Ziv-Barber/officegen",
    "version": "0.4.4"
}
```



# misc
- this document was created with [utility2](https://github.com/kaizhu256/node-utility2)
