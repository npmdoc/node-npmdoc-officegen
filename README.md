# api documentation for  [officegen (v0.4.4)](https://github.com/Ziv-Barber/officegen#readme)  [![npm package](https://img.shields.io/npm/v/npmdoc-officegen.svg?style=flat-square)](https://www.npmjs.org/package/npmdoc-officegen) [![travis-ci.org build-status](https://api.travis-ci.org/npmdoc/node-npmdoc-officegen.svg)](https://travis-ci.org/npmdoc/node-npmdoc-officegen)
#### Office Open XML Generator using Node.js streams. Supporting Microsoft Office 2007 and later Word (docx), PowerPoint (pptx,ppsx) and Excel (xlsx). This module is for all frameworks and environments. No need for any commandline tool - this module is doing e

[![NPM](https://nodei.co/npm/officegen.png?downloads=true)](https://www.npmjs.com/package/officegen)

[![apidoc](https://npmdoc.github.io/node-npmdoc-officegen/build/screenCapture.buildNpmdoc.browser._2Fhome_2Ftravis_2Fbuild_2Fnpmdoc_2Fnode-npmdoc-officegen_2Ftmp_2Fbuild_2Fapidoc.html.png)](https://npmdoc.github.io/node-npmdoc-officegen/build/apidoc.html)

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
            "name": "zivbarber",
            "email": "zivbarber@gmail.com"
        }
    ],
    "name": "officegen",
    "optionalDependencies": {},
    "readme": "ERROR: No README data found!",
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



# <a name="apidoc.tableOfContents"></a>[table of contents](#apidoc.tableOfContents)

#### [module officegen](#apidoc.module.officegen)
1.  [function <span class="apidocSignatureSpan">officegen.</span>docplug ( genobj, genPrivate, docType, defValuesFunc )](#apidoc.element.officegen.docplug)
1.  [function <span class="apidocSignatureSpan">officegen.</span>docx_p ( genobj, genPrivate, docType, dataContainer, extraSettings, options )](#apidoc.element.officegen.docx_p)
1.  [function <span class="apidocSignatureSpan">officegen.</span>docxplg_headfoot ( pluginsman )](#apidoc.element.officegen.docxplg_headfoot)
1.  [function <span class="apidocSignatureSpan">officegen.</span>officechart (chartInfo)](#apidoc.element.officegen.officechart)
1.  [function <span class="apidocSignatureSpan">officegen.</span>pptxplg_layouts ( pluginsman )](#apidoc.element.officegen.pptxplg_layouts)
1.  [function <span class="apidocSignatureSpan">officegen.</span>pptxplg_speakernotes ( pluginsman )](#apidoc.element.officegen.pptxplg_speakernotes)
1.  [function <span class="apidocSignatureSpan">officegen.</span>pptxplg_widescreen ( pluginsman )](#apidoc.element.officegen.pptxplg_widescreen)
1.  [function <span class="apidocSignatureSpan">officegen.</span>setVerboseMode ( new_state )](#apidoc.element.officegen.setVerboseMode)
1.  object <span class="apidocSignatureSpan">officegen.</span>charts
1.  object <span class="apidocSignatureSpan">officegen.</span>docType
1.  object <span class="apidocSignatureSpan">officegen.</span>docplug.prototype
1.  object <span class="apidocSignatureSpan">officegen.</span>docx_p.prototype
1.  object <span class="apidocSignatureSpan">officegen.</span>docxplg_headfoot.prototype
1.  object <span class="apidocSignatureSpan">officegen.</span>docxtable
1.  object <span class="apidocSignatureSpan">officegen.</span>genofficetable
1.  object <span class="apidocSignatureSpan">officegen.</span>msexcel_builder
1.  object <span class="apidocSignatureSpan">officegen.</span>msofficegen
1.  object <span class="apidocSignatureSpan">officegen.</span>openofficegen
1.  object <span class="apidocSignatureSpan">officegen.</span>plugins
1.  object <span class="apidocSignatureSpan">officegen.</span>pptxplg_layouts.prototype
1.  object <span class="apidocSignatureSpan">officegen.</span>pptxplg_speakernotes.prototype
1.  object <span class="apidocSignatureSpan">officegen.</span>pptxplg_widescreen.prototype
1.  object <span class="apidocSignatureSpan">officegen.</span>schema
1.  string <span class="apidocSignatureSpan">officegen.</span>version

#### [module officegen.charts](#apidoc.module.officegen.charts)
1.  [function <span class="apidocSignatureSpan">officegen.charts.</span>area (options)](#apidoc.element.officegen.charts.area)
1.  [function <span class="apidocSignatureSpan">officegen.charts.</span>bar (options)](#apidoc.element.officegen.charts.bar)
1.  [function <span class="apidocSignatureSpan">officegen.charts.</span>column (options)](#apidoc.element.officegen.charts.column)
1.  [function <span class="apidocSignatureSpan">officegen.charts.</span>doughnut (options)](#apidoc.element.officegen.charts.doughnut)
1.  [function <span class="apidocSignatureSpan">officegen.charts.</span>group-bar (options)](#apidoc.element.officegen.charts.group-bar)
1.  [function <span class="apidocSignatureSpan">officegen.charts.</span>line (options)](#apidoc.element.officegen.charts.line)
1.  [function <span class="apidocSignatureSpan">officegen.charts.</span>null (options)](#apidoc.element.officegen.charts.null)
1.  [function <span class="apidocSignatureSpan">officegen.charts.</span>pie (options)](#apidoc.element.officegen.charts.pie)
1.  [function <span class="apidocSignatureSpan">officegen.charts.</span>stacked-column (options)](#apidoc.element.officegen.charts.stacked-column)

#### [module officegen.docplug](#apidoc.module.officegen.docplug)
1.  [function <span class="apidocSignatureSpan">officegen.</span>docplug ( genobj, genPrivate, docType, defValuesFunc )](#apidoc.element.officegen.docplug.docplug)

#### [module officegen.docplug.prototype](#apidoc.module.officegen.docplug.prototype)
1.  [function <span class="apidocSignatureSpan">officegen.docplug.prototype.</span>emitEvent ( eventType, eventData )](#apidoc.element.officegen.docplug.prototype.emitEvent)
1.  [function <span class="apidocSignatureSpan">officegen.docplug.prototype.</span>getDataStorage ()](#apidoc.element.officegen.docplug.prototype.getDataStorage)
1.  [function <span class="apidocSignatureSpan">officegen.docplug.prototype.</span>getFeaturesStorage ()](#apidoc.element.officegen.docplug.prototype.getFeaturesStorage)
1.  [function <span class="apidocSignatureSpan">officegen.docplug.prototype.</span>getVerboseMode ( moduleName )](#apidoc.element.officegen.docplug.prototype.getVerboseMode)
1.  [function <span class="apidocSignatureSpan">officegen.docplug.prototype.</span>logIfVerbose ( message )](#apidoc.element.officegen.docplug.prototype.logIfVerbose)
1.  [function <span class="apidocSignatureSpan">officegen.docplug.prototype.</span>registerCallback ( eventType, cbFunc )](#apidoc.element.officegen.docplug.prototype.registerCallback)

#### [module officegen.docx_p](#apidoc.module.officegen.docx_p)
1.  [function <span class="apidocSignatureSpan">officegen.</span>docx_p ( genobj, genPrivate, docType, dataContainer, extraSettings, options )](#apidoc.element.officegen.docx_p.docx_p)

#### [module officegen.docx_p.prototype](#apidoc.module.officegen.docx_p.prototype)
1.  [function <span class="apidocSignatureSpan">officegen.docx_p.prototype.</span>addHorizontalLine ()](#apidoc.element.officegen.docx_p.prototype.addHorizontalLine)
1.  [function <span class="apidocSignatureSpan">officegen.docx_p.prototype.</span>addImage ( image_path, opt, image_format_type )](#apidoc.element.officegen.docx_p.prototype.addImage)
1.  [function <span class="apidocSignatureSpan">officegen.docx_p.prototype.</span>addLineBreak ()](#apidoc.element.officegen.docx_p.prototype.addLineBreak)
1.  [function <span class="apidocSignatureSpan">officegen.docx_p.prototype.</span>addText ( text_msg, opt, flag_data )](#apidoc.element.officegen.docx_p.prototype.addText)
1.  [function <span class="apidocSignatureSpan">officegen.docx_p.prototype.</span>endBookmark ()](#apidoc.element.officegen.docx_p.prototype.endBookmark)
1.  [function <span class="apidocSignatureSpan">officegen.docx_p.prototype.</span>startBookmark ( anchorName )](#apidoc.element.officegen.docx_p.prototype.startBookmark)

#### [module officegen.docxplg_headfoot](#apidoc.module.officegen.docxplg_headfoot)
1.  [function <span class="apidocSignatureSpan">officegen.</span>docxplg_headfoot ( pluginsman )](#apidoc.element.officegen.docxplg_headfoot.docxplg_headfoot)

#### [module officegen.docxplg_headfoot.prototype](#apidoc.module.officegen.docxplg_headfoot.prototype)
1.  [function <span class="apidocSignatureSpan">officegen.docxplg_headfoot.prototype.</span>beforeGen ( docObj )](#apidoc.element.officegen.docxplg_headfoot.prototype.beforeGen)
1.  [function <span class="apidocSignatureSpan">officegen.docxplg_headfoot.prototype.</span>cbEndNotes ( data )](#apidoc.element.officegen.docxplg_headfoot.prototype.cbEndNotes)
1.  [function <span class="apidocSignatureSpan">officegen.docxplg_headfoot.prototype.</span>cbFootNotes ( data )](#apidoc.element.officegen.docxplg_headfoot.prototype.cbFootNotes)
1.  [function <span class="apidocSignatureSpan">officegen.docxplg_headfoot.prototype.</span>extenddocxApi ( docObj )](#apidoc.element.officegen.docxplg_headfoot.prototype.extenddocxApi)

#### [module officegen.docxtable](#apidoc.module.officegen.docxtable)
1.  [function <span class="apidocSignatureSpan">officegen.docxtable.</span>_getBase (rowSpecs, colSpecs, opts)](#apidoc.element.officegen.docxtable._getBase)
1.  [function <span class="apidocSignatureSpan">officegen.docxtable.</span>_getCell (val, opts, tblOpts)](#apidoc.element.officegen.docxtable._getCell)
1.  [function <span class="apidocSignatureSpan">officegen.docxtable.</span>_getColSpecs (cols, opts)](#apidoc.element.officegen.docxtable._getColSpecs)
1.  [function <span class="apidocSignatureSpan">officegen.docxtable.</span>_getRow (cells, opts)](#apidoc.element.officegen.docxtable._getRow)
1.  [function <span class="apidocSignatureSpan">officegen.docxtable.</span>_tblGrid (opts)](#apidoc.element.officegen.docxtable._tblGrid)
1.  [function <span class="apidocSignatureSpan">officegen.docxtable.</span>getTable (rows, tblOpts)](#apidoc.element.officegen.docxtable.getTable)

#### [module officegen.genofficetable](#apidoc.module.officegen.genofficetable)
1.  [function <span class="apidocSignatureSpan">officegen.genofficetable.</span>_getBase (rowSpecs, colSpecs, options)](#apidoc.element.officegen.genofficetable._getBase)
1.  [function <span class="apidocSignatureSpan">officegen.genofficetable.</span>_getCell (val, options, idx, row_idx)](#apidoc.element.officegen.genofficetable._getCell)
1.  [function <span class="apidocSignatureSpan">officegen.genofficetable.</span>_getColSpecs (rows, options)](#apidoc.element.officegen.genofficetable._getColSpecs)
1.  [function <span class="apidocSignatureSpan">officegen.genofficetable.</span>_getRow (cells, options)](#apidoc.element.officegen.genofficetable._getRow)
1.  [function <span class="apidocSignatureSpan">officegen.genofficetable.</span>_tblGrid (idx, options)](#apidoc.element.officegen.genofficetable._tblGrid)
1.  [function <span class="apidocSignatureSpan">officegen.genofficetable.</span>getTable (rows, options)](#apidoc.element.officegen.genofficetable.getTable)

#### [module officegen.msexcel_builder](#apidoc.module.officegen.msexcel_builder)
1.  [function <span class="apidocSignatureSpan">officegen.msexcel_builder.</span>createWorkbook (fpath, fname)](#apidoc.element.officegen.msexcel_builder.createWorkbook)

#### [module officegen.msofficegen](#apidoc.module.officegen.msofficegen)
1.  [function <span class="apidocSignatureSpan">officegen.msofficegen.</span>makemsdoc ( genobj, new_type, options, gen_private, type_info )](#apidoc.element.officegen.msofficegen.makemsdoc)

#### [module officegen.officechart](#apidoc.module.officegen.officechart)
1.  [function <span class="apidocSignatureSpan">officegen.</span>officechart (chartInfo)](#apidoc.element.officegen.officechart.officechart)
1.  [function <span class="apidocSignatureSpan">officegen.officechart.</span>getChartBase (chartInfo)](#apidoc.element.officegen.officechart.getChartBase)

#### [module officegen.openofficegen](#apidoc.module.officegen.openofficegen)
1.  [function <span class="apidocSignatureSpan">officegen.openofficegen.</span>makeoodoc ( genobj, new_type, options, gen_private, type_info )](#apidoc.element.officegen.openofficegen.makeoodoc)

#### [module officegen.plugins](#apidoc.module.officegen.plugins)
1.  [function <span class="apidocSignatureSpan">officegen.plugins.</span>getDocTypeByName ( typeName )](#apidoc.element.officegen.plugins.getDocTypeByName)
1.  [function <span class="apidocSignatureSpan">officegen.plugins.</span>getPrototypeByName ( typeName )](#apidoc.element.officegen.plugins.getPrototypeByName)
1.  [function <span class="apidocSignatureSpan">officegen.plugins.</span>getVerboseMode ( docType, moduleName )](#apidoc.element.officegen.plugins.getVerboseMode)
1.  [function <span class="apidocSignatureSpan">officegen.plugins.</span>registerDocType ( typeName, createFunc, schema_data, docType, displayName )](#apidoc.element.officegen.plugins.registerDocType)
1.  [function <span class="apidocSignatureSpan">officegen.plugins.</span>registerParserType ( typeName, parserFunc, extra_data, displayName )](#apidoc.element.officegen.plugins.registerParserType)
1.  [function <span class="apidocSignatureSpan">officegen.plugins.</span>registerPrototype ( typeName, baseObj, displayName )](#apidoc.element.officegen.plugins.registerPrototype)

#### [module officegen.pptxplg_layouts](#apidoc.module.officegen.pptxplg_layouts)
1.  [function <span class="apidocSignatureSpan">officegen.</span>pptxplg_layouts ( pluginsman )](#apidoc.element.officegen.pptxplg_layouts.pptxplg_layouts)

#### [module officegen.pptxplg_layouts.prototype](#apidoc.module.officegen.pptxplg_layouts.prototype)
1.  [function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>beforeGen ( docObj )](#apidoc.element.officegen.pptxplg_layouts.prototype.beforeGen)
1.  [function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>cbMakePptxLayout10 ( data )](#apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout10)
1.  [function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>cbMakePptxLayout11 ( data )](#apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout11)
1.  [function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>cbMakePptxLayout2 ( data )](#apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout2)
1.  [function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>cbMakePptxLayout3 ( data )](#apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout3)
1.  [function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>cbMakePptxLayout4 ( data )](#apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout4)
1.  [function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>cbMakePptxLayout5 ( data )](#apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout5)
1.  [function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>cbMakePptxLayout6 ( data )](#apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout6)
1.  [function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>cbMakePptxLayout7 ( data )](#apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout7)
1.  [function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>cbMakePptxLayout8 ( data )](#apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout8)
1.  [function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>cbMakePptxLayout9 ( data )](#apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout9)
1.  [function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>extendPptxApi ( docObj )](#apidoc.element.officegen.pptxplg_layouts.prototype.extendPptxApi)
1.  [function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>extendPptxSlideApi ( docData )](#apidoc.element.officegen.pptxplg_layouts.prototype.extendPptxSlideApi)

#### [module officegen.pptxplg_speakernotes](#apidoc.module.officegen.pptxplg_speakernotes)
1.  [function <span class="apidocSignatureSpan">officegen.</span>pptxplg_speakernotes ( pluginsman )](#apidoc.element.officegen.pptxplg_speakernotes.pptxplg_speakernotes)

#### [module officegen.pptxplg_speakernotes.prototype](#apidoc.module.officegen.pptxplg_speakernotes.prototype)
1.  [function <span class="apidocSignatureSpan">officegen.pptxplg_speakernotes.prototype.</span>beforeGen ( docObj )](#apidoc.element.officegen.pptxplg_speakernotes.prototype.beforeGen)
1.  [function <span class="apidocSignatureSpan">officegen.pptxplg_speakernotes.prototype.</span>cbMakePptxNotesMasters ( data )](#apidoc.element.officegen.pptxplg_speakernotes.prototype.cbMakePptxNotesMasters)
1.  [function <span class="apidocSignatureSpan">officegen.pptxplg_speakernotes.prototype.</span>cbMakePptxSpkNote ( data )](#apidoc.element.officegen.pptxplg_speakernotes.prototype.cbMakePptxSpkNote)
1.  [function <span class="apidocSignatureSpan">officegen.pptxplg_speakernotes.prototype.</span>cbMakeTheme2 ( data )](#apidoc.element.officegen.pptxplg_speakernotes.prototype.cbMakeTheme2)
1.  [function <span class="apidocSignatureSpan">officegen.pptxplg_speakernotes.prototype.</span>extendPptxSlideApi ( docData )](#apidoc.element.officegen.pptxplg_speakernotes.prototype.extendPptxSlideApi)
1.  [function <span class="apidocSignatureSpan">officegen.pptxplg_speakernotes.prototype.</span>presentationGen ( docData )](#apidoc.element.officegen.pptxplg_speakernotes.prototype.presentationGen)

#### [module officegen.pptxplg_widescreen](#apidoc.module.officegen.pptxplg_widescreen)
1.  [function <span class="apidocSignatureSpan">officegen.</span>pptxplg_widescreen ( pluginsman )](#apidoc.element.officegen.pptxplg_widescreen.pptxplg_widescreen)

#### [module officegen.pptxplg_widescreen.prototype](#apidoc.module.officegen.pptxplg_widescreen.prototype)
1.  [function <span class="apidocSignatureSpan">officegen.pptxplg_widescreen.prototype.</span>extendPptxApi ( docObj )](#apidoc.element.officegen.pptxplg_widescreen.prototype.extendPptxApi)



# <a name="apidoc.module.officegen"></a>[module officegen](#apidoc.module.officegen)

#### <a name="apidoc.element.officegen.docplug"></a>[function <span class="apidocSignatureSpan">officegen.</span>docplug ( genobj, genPrivate, docType, defValuesFunc )](#apidoc.element.officegen.docplug)
- description and source-code
```javascript
function MakeDocPluginapi( genobj, genPrivate, docType, defValuesFunc ) {
	// Save everything because we'll need it later:
	this.docType = docType;
	this.genPrivate = genPrivate;
	this.ogPluginsApi = genPrivate.plugs;
	this.genobj = genobj;
	this.defValuesFunc = defValuesFunc;
	this.callbacksList = {};
	this.plugsList = [];

	// Here we can put anything specific to this document type BUT it's not data so if we'll clear the data inside the document we'
ll not need to clear anything here:
	this.genPrivate.features.type[this.docType] = this;

	// Prepare the object that we'll use to store data specific to this document type:
	this.genPrivate.type[this.docType] = {};
	if ( typeof defValuesFunc === 'function' ) {
		// Put the default data:
		defValuesFunc ( this );
	} // Endif.

	// Catch the clear document data event and connect it to us:
	genobj.on ( 'clearDocType', clearDocData ( this ) );
	return this;
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.docx_p"></a>[function <span class="apidocSignatureSpan">officegen.</span>docx_p ( genobj, genPrivate, docType, dataContainer, extraSettings, options )](#apidoc.element.officegen.docx_p)
- description and source-code
```javascript
function MakeDocxP( genobj, genPrivate, docType, dataContainer, extraSettings, options ) {
	// Save everything because we'll need it later:
	this.docType = docType;
	this.genPrivate = genPrivate;
	this.ogPluginsApi = genPrivate.plugs; // Generic officegen API for plugins.
	this.msPluginsApi = genPrivate.plugs.type.msoffice; // msoffice plugins API.
	this.genobj = genobj;
	this.data = [];
	this.extraSettings = extraSettings || {};
	this.options = options || {};

	this.mainPath = genPrivate.features.type.msoffice.main_path; // The "folder" name inside the document zip that all the specific
 resources of this document type are stored.
	this.mainPathFile = genPrivate.features.type.msoffice.main_path_file; // The name of the main real xml resource of this document
.
	this.relsMain = genPrivate.type.msoffice.rels_main; // Main rels file.
	this.relsApp = genPrivate.type.msoffice.rels_app; // Main rels file inside the specific document type "folder".
	this.filesList = genPrivate.type.msoffice.files_list; // Resources list xml.
	this.srcFilesList = genPrivate.type.msoffice.src_files_list; // For storing extra files inside the document zip.


	if ( dataContainer ) {
		dataContainer.push ( this );
	} // Endif.

	return this;
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.docxplg_headfoot"></a>[function <span class="apidocSignatureSpan">officegen.</span>docxplg_headfoot ( pluginsman )](#apidoc.element.officegen.docxplg_headfoot)
- description and source-code
```javascript
function MakeHeadfootPlugin( pluginsman ) {
	var funcThis = this;

	// You can change it if you want to support more types, since that the Word and Excel document generators also supporting very
similar plugins:
	if ( pluginsman.docType !== 'docx' ) {
		throw "[docx-headfoot] This plugin supporting only Word based documents.";
	} // Endif.

	this.ogPluginsApi = pluginsman.ogPluginsApi; // Generic officegen API for plugins.
	this.msPluginsApi = pluginsman.genPrivate.plugs.type.msoffice; // msoffice plugins API.
	this.pluginsman = pluginsman; // Document type specific plugins API.

	this.docxData = pluginsman.getDataStorage (); // Here you can store any temporary data needed for generating the document and depending
 on the data filled by the user.

	this.mainPath = pluginsman.genPrivate.features.type.msoffice.main_path; // The "folder" name inside the document zip that all the
 specific resources of this document type are stored.
	this.mainPathFile = pluginsman.genPrivate.features.type.msoffice.main_path_file; // The name of the main real xml resource of this
 document.
	this.relsMain = pluginsman.genPrivate.type.msoffice.rels_main; // Main rels file.
	this.relsApp = pluginsman.genPrivate.type.msoffice.rels_app; // Main rels file inside the specific document type "folder".
	this.filesList = pluginsman.genPrivate.type.msoffice.files_list; // Resources list xml.
	this.srcFilesList = pluginsman.genPrivate.type.msoffice.src_files_list; // For storing extra files inside the document zip.

	//
	// Catch events inside the document so we can extent it:
	//

	// We want to extend the main API of the docx document object:
	pluginsman.registerCallback ( 'makeDocApi', function ( docObj ) { funcThis.extenddocxApi ( docObj ); } );

	// This event tell us that the generator is about to start working:
	pluginsman.registerCallback ( 'beforeGen', function ( docObj ) { funcThis.beforeGen ( docObj ); } );

	return this;
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.officechart"></a>[function <span class="apidocSignatureSpan">officegen.</span>officechart (chartInfo)](#apidoc.element.officegen.officechart)
- description and source-code
```javascript
function OfficeChart(chartInfo) {

  if (chartInfo instanceof OfficeChart) {
    return chartInfo;
  }

  return {

    chartSpec: null, // Javascript object that represents the XML tree for the PowerPoint chart

    toXML: function () {
      return xmlBuilder.create(this.chartSpec, {version: '1.0', encoding: 'UTF-8', standalone: true}).end({ pretty: true, indent
: '  ', newline: '\n' });
    },

    toJSON: function () {
      return this.chartSpec;
    },

    getClass: function() { return "OfficeChart"; },

    ///
    /// @brief Create XML representation of chart object
    /// @param[chartInfo] object
    ///  {
    ///			title: 'eSurvey chart',
    ///			data:  [ // array of series
    ///				{
    ///					name: 'Income',
    ///					labels: ['2005', '2006', '2007', '2008', '2009'],
    ///					values: [23.5, 26.2, 30.1, 29.5, 24.6],
    ///         color: 'ff0000'
    ///				}
    ///			],
    ///     overlap:  "0",
    ///     gapWidth: "150"
    ///
    /// 	}

    initialize: function (chartInfo) {

      if (chartInfo.getClass &&  chartInfo.getClass()== 'OfficeChart') {
        return chartInfo;
      }

      // overlap ["50"] is handled as an option within the chartbase
      // gapWidth ["150"] is handled as an option within the chartbase
      // valAxisCrossAtMaxCategory [true|false] is handled as an option within the chart base
      // catAxisReverseOrder [true|false] is handled as an option within the chart base

      this.chartSpec = OfficeChart.getChartBase(chartInfo);  // get foundation XML for the chart type

      // Below are methods for handling options with more complex XML to mix in
      this.setData(chartInfo['data']);
      this.setTitle(chartInfo.title || chartInfo.name);
      this.setValAxisTitle(chartInfo.valAxisTitle);
      this.setCatAxisTitle(chartInfo.catAxisTitle);
      this.setValAxisNumFmt(chartInfo.valAxisNumFmt);
      this.setValAxisScale(chartInfo.valAxisMinValue, chartInfo.valAxisMaxValue);
      this.setTextSize(chartInfo.fontSize);
      this.mergeChartXml(chartInfo.xml);
      this.setValAxisMajorGridlines(chartInfo.valAxisMajorGridlines);
      this.setValAxisMinorGridlines(chartInfo.valAxisMinorGridlines);

      return this;
    },

    setTextSize: function(textSize) {
      if (textSize !== undefined) {
         var textRef = this._text(textSize);
        _.merge(this.chartSpec['c:chartSpace'], textRef);
      }
    },

    setTitle: function (title) {
      if (title !== undefined) {
        var titleRef = this._title(chartInfo.title || chartInfo.name);
        _.merge(this.chartSpec['c:chartSpace']['c:chart'], titleRef)
      }
    },

    setValAxisTitle: function (title) {
      if (title) {
        var titleRef = this._title(title);
        _.merge(this.chartSpec['c:chartSpace']['c:chart']['c:plotArea']['c:valAx'], titleRef);
      }
    },

    setCatAxisTitle: function (title) {
      if (title) {
        var titleRef = this._title(title);
        _.merge(this.chartSpec['c:chartSpace']['c:chart']['c:plotArea']['c:catAx'], titleRef);
      }
    },

    setValAxisNumFmt: function (format) {
      if (format !== undefined) {
        var numFmtRef = this._numFmt(format);
        _.merge(this.chartSpec['c:chartSpace']['c:chart']['c:plotArea']['c:valAx'], numFmtRef)
      }
    },

    setValAxisScale: function (min, max) {
      if (min!==undefined || max!== undefined) {
        var scalingRef = this._scaling(min, max);
        _.merge(this.chartSpec['c:chartSpace']['c:chart']['c:plotArea']['c:valAx'], scalingRef);
      }
    },

    mergeChartXml: function (xml) {
      if (xml !== undefined) {
        _.merge(this.chartSpec['c:chartSpace'], xml);
      }
    },

    setValAxisMajorGridlines: function(boolean) {
      if (boolean) {
        this.chartSpec['c:chartSpace']['c:chart']['c:plotArea']['c:valAx']['c:majorGridlines'] = {};
      }
    },
    setValAxisMinorGridlines: function(boolean) {
      if (boolean) {
        this.chartSpec[ ...
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.pptxplg_layouts"></a>[function <span class="apidocSignatureSpan">officegen.</span>pptxplg_layouts ( pluginsman )](#apidoc.element.officegen.pptxplg_layouts)
- description and source-code
```javascript
function MakeLayoutPlugin( pluginsman ) {
	var funcThis = this;

	this.ogPluginsApi = pluginsman.ogPluginsApi; // Generic officegen API for plugins.
	this.msPluginsApi = pluginsman.genPrivate.plugs.type.msoffice; // msoffice plugins API.
	this.pluginsman = pluginsman; // Document type specific plugins API.

	this.pptxData = pluginsman.getDataStorage (); // Here you can store any temporary data needed for generating the document and depending
 on the data filled by the user.

	this.mainPath = pluginsman.genPrivate.features.type.msoffice.main_path; // The "folder" name inside the document zip that all the
 specific resources of this document type are stored.
	this.mainPathFile = pluginsman.genPrivate.features.type.msoffice.main_path_file; // The name of the main real xml resource of this
 document.
	this.relsMain = pluginsman.genPrivate.type.msoffice.rels_main; // Main rels file.
	this.relsApp = pluginsman.genPrivate.type.msoffice.rels_app; // Main rels file inside the specific document type "folder".
	this.filesList = pluginsman.genPrivate.type.msoffice.files_list; // Resources list xml.
	this.srcFilesList = pluginsman.genPrivate.type.msoffice.src_files_list; // For storing extra files inside the document zip.

	//
	// Catch events inside the document so we can extent it:
	//

	// We want to extend the main API of the pptx document object:
	pluginsman.registerCallback ( 'makeDocApi', function ( docObj ) { funcThis.extendPptxApi ( docObj ); } );

	// We want to extend the slide object API:
	pluginsman.registerCallback ( 'newPage', function ( docData ) { funcThis.extendPptxSlideApi ( docData ); } );

	// This event tell us that the generator is about to start working:
	pluginsman.registerCallback ( 'beforeGen', function ( docObj ) { funcThis.beforeGen ( docObj ); } );

	return this;
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.pptxplg_speakernotes"></a>[function <span class="apidocSignatureSpan">officegen.</span>pptxplg_speakernotes ( pluginsman )](#apidoc.element.officegen.pptxplg_speakernotes)
- description and source-code
```javascript
function MakeSpknotesPlugin( pluginsman ) {
	var funcThis = this;

	if ( pluginsman.docType !== 'pptx' && pluginsman.docType !== 'ppsx' ) {
		throw "[pptx-speakernotes] This plugin supporting only PowerPoint based documents.";
	} // Endif.

	this.ogPluginsApi = pluginsman.ogPluginsApi; // Generic officegen API for plugins.
	this.msPluginsApi = pluginsman.genPrivate.plugs.type.msoffice; // msoffice plugins API.
	this.pluginsman = pluginsman; // Document type specific plugins API.

	this.pptxData = pluginsman.getDataStorage (); // Here you can store any temporary data needed for generating the document and depending
 on the data filled by the user.

	this.mainPath = pluginsman.genPrivate.features.type.msoffice.main_path; // The "folder" name inside the document zip that all the
 specific resources of this document type are stored.
	this.mainPathFile = pluginsman.genPrivate.features.type.msoffice.main_path_file; // The name of the main real xml resource of this
 document.
	this.relsMain = pluginsman.genPrivate.type.msoffice.rels_main; // Main rels file.
	this.relsApp = pluginsman.genPrivate.type.msoffice.rels_app; // Main rels file inside the specific document type "folder".
	this.filesList = pluginsman.genPrivate.type.msoffice.files_list; // Resources list xml.
	this.srcFilesList = pluginsman.genPrivate.type.msoffice.src_files_list; // For storing extra files inside the document zip.

	//
	// Catch events inside the document so we can extent it:
	//

	// We want to extend the slide object API:
	pluginsman.registerCallback ( 'newPage', function ( docData ) { funcThis.extendPptxSlideApi ( docData ); } );

	// This event tell us that the generator is about to start working:
	pluginsman.registerCallback ( 'beforeGen', function ( docObj ) { funcThis.beforeGen ( docObj ); } );

	// This event tell us to add our xml code to the main presentation file:
	pluginsman.registerCallback ( 'presentationGen', function ( docData ) { funcThis.presentationGen ( docData ); } );

	return this;
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.pptxplg_widescreen"></a>[function <span class="apidocSignatureSpan">officegen.</span>pptxplg_widescreen ( pluginsman )](#apidoc.element.officegen.pptxplg_widescreen)
- description and source-code
```javascript
function MakeWidescreenPlugin( pluginsman ) {
	var funcThis = this;

	if ( pluginsman.docType !== 'pptx' && pluginsman.docType !== 'ppsx' ) {
		throw "[pptx-widescreen] This plugin supporting only PowerPoint based documents.";
	} // Endif.

	this.pluginsman = pluginsman;
	this.pptxData = pluginsman.getDataStorage ();

	// We want to extend the main API of the pptx document object:
	pluginsman.registerCallback ( 'makeDocApi', function ( docObj ) { funcThis.extendPptxApi ( docObj ); } );
	return this;
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.setVerboseMode"></a>[function <span class="apidocSignatureSpan">officegen.</span>setVerboseMode ( new_state )](#apidoc.element.officegen.setVerboseMode)
- description and source-code
```javascript
function setVerboseMode( new_state ) {
	officegenGlobals.settings.verbose = new_state;
}
```
- example usage
```shell
...
'''

#### Debugging ####

If needed, you can activate some verbose messages (warning: this does not cover all part of the lib yet) with :

'''js
officegen.setVerboseMode(true);
'''


<a name="a6"></a>

## FAQ: ##
...
```



# <a name="apidoc.module.officegen.charts"></a>[module officegen.charts](#apidoc.module.officegen.charts)

#### <a name="apidoc.element.officegen.charts.area"></a>[function <span class="apidocSignatureSpan">officegen.charts.</span>area (options)](#apidoc.element.officegen.charts.area)
- description and source-code
```javascript
area = function (options) {
  options = options || {};
  return {
    "c:chartSpace": {
      "@xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "@xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "@xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "c:lang": {"@val": "en-US"},
      "c:chart": {
        "c:plotArea": {
          "c:layout": {},
          "c:areaChart": {
            "c:grouping": {"@val": "standard"},
            "#list": [
              {"c:axId": {"@val": "64451712"}},
              {"c:axId": {"@val": "64453248"}}
            ]
          },
          "c:catAx": {
            "c:axId": {"@val": "64451712"},
            "c:scaling": {
              "c:orientation": {"@val": options.catAxisReverseOrder ? "maxMin" : "minMax"}
            },
            "c:axPos": {"@val": "l"},
            "c:tickLblPos": {"@val": "nextTo"},
            "c:crossAx": {"@val": "64453248"},
            "c:crosses": {"@val": "autoZero"},
            "c:auto": {"@val": "1"},
            "c:lblAlgn": {"@val": "ctr"},
            "c:lblOffset": {"@val": "100"}
          },
          "c:valAx": {
            "c:axId": {"@val": "64453248"},
            "c:scaling": {
              "c:orientation": {"@val": "minMax"}
            },
            "c:axPos": {"@val": "b"},
//              "c:majorGridlines": {},
            "c:numFmt": {
              "@formatCode": "General",
              "@sourceLinked": "1"
            },
            "c:tickLblPos": {"@val": "nextTo"},
            "c:crossAx": {"@val": "64451712"},
            "c:crosses": {"@val": options.valAxisCrossAtMaxCategory ? "max" : "autoZero"},
            "c:crossBetween": {"@val": "between"}
          }
        },
        "c:legend": {
          "c:legendPos": {"@val": "r"},
          "c:layout": {}
        },
        "c:plotVisOnly": {"@val": "1"}
      },
      "c:txPr": {
        "a:bodyPr": {},
        "a:lstStyle": {},
        "a:p": {
          "a:pPr": {
            "a:defRPr": {"@sz": "1800"}
          },
          "a:endParaRPr": {"@lang": "en-US"}
        }
      },
      "c:externalData": {"@r:id": "rId1"}
    }
  }
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.charts.bar"></a>[function <span class="apidocSignatureSpan">officegen.charts.</span>bar (options)](#apidoc.element.officegen.charts.bar)
- description and source-code
```javascript
bar = function (options) {
  options = options || {};
  return {
    "c:chartSpace": {
      "@xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "@xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "@xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "c:lang": {"@val": "en-US"},
      "c:chart": {
        "c:plotArea": {
          "c:layout": {},
          "c:barChart": {
            "c:barDir": {"@val": "bar"},
            "c:grouping": {"@val": "clustered"},
            "#list": [
              {"c:axId": {"@val": "64451712"}},
              {"c:axId": {"@val": "64453248"}}
            ]
          },
          "c:catAx": {
            "c:axId": {"@val": "64451712"},
            "c:scaling": {
              "c:orientation": {"@val": options.catAxisReverseOrder ? "maxMin" : "minMax"}
            },
            "c:axPos": {"@val": "l"},
            "c:tickLblPos": {"@val": "nextTo"},
            "c:crossAx": {"@val": "64453248"},
            "c:crosses": {"@val": "autoZero"},
            "c:auto": {"@val": "1"},
            "c:lblAlgn": {"@val": "ctr"},
            "c:lblOffset": {"@val": "100"}
          },
          "c:valAx": {
            "c:axId": {"@val": "64453248"},
            "c:scaling": {
              "c:orientation": {"@val": "minMax"}
            },
            "c:axPos": {"@val": "b"},
//              "c:majorGridlines": {},
            "c:numFmt": {
              "@formatCode": "General",
              "@sourceLinked": "1"
            },
            "c:tickLblPos": {"@val": "nextTo"},
            "c:crossAx": {"@val": "64451712"},
            "c:crosses": {"@val": options.valAxisCrossAtMaxCategory ? "max" : "autoZero"},
            "c:crossBetween": {"@val": "between"}
          }
        },
        "c:legend": {
          "c:legendPos": {"@val": "r"},
          "c:layout": {}
        },
        "c:plotVisOnly": {"@val": "1"}
      },
      "c:txPr": {
        "a:bodyPr": {},
        "a:lstStyle": {},
        "a:p": {
          "a:pPr": {
            "a:defRPr": {"@sz": "1800"}
          },
          "a:endParaRPr": {"@lang": "en-US"}
        }
      },
      "c:externalData": {"@r:id": "rId1"}
    }
  };
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.charts.column"></a>[function <span class="apidocSignatureSpan">officegen.charts.</span>column (options)](#apidoc.element.officegen.charts.column)
- description and source-code
```javascript
column = function (options) {
  options = options || {};
  return {
    "c:chartSpace": {
      "@xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "@xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "@xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "c:lang": {"@val": "en-US"},
      "c:date1904": {"@val": "1"},
      "c:chart": {
        "c:plotArea": {
          "c:layout": {},
          "c:barChart": {
            "c:barDir": {"@val": "col"},
            "c:grouping": {"@val": "clustered"},
            "c:overlap": {"@val": options.overlap || "0"},
            "c:gapWidth": {"@val": options.gapWidth || "150"},

            "#list": [
              {"c:axId": {"@val": "64451712"}},
              {"c:axId": {"@val": "64453248"}}
            ]
          },
          "c:catAx": {
            "c:axId": {"@val": "64451712"},
            "c:scaling": {
              "c:orientation": {"@val": options.catAxisReverseOrder ? "maxMin" : "minMax"}
            },
            "c:axPos": {"@val": "l"},
            "c:tickLblPos": {"@val": "nextTo"},
            "c:crossAx": {"@val": "64453248"},
            "c:crosses": {"@val": "autoZero"},
            "c:auto": {"@val": "1"},
            "c:lblAlgn": {"@val": "ctr"},
            "c:lblOffset": {"@val": "100"}
          },
          "c:valAx": {
            "c:axId": {"@val": "64453248"},
            "c:scaling": {
              "c:orientation": {"@val": "minMax"}
            },
            "c:axPos": {"@val": "b"},
//              "c:majorGridlines": {},
            "c:numFmt": {
              "@formatCode": "General",
              "@sourceLinked": "1"
            },
            "c:tickLblPos": {"@val": "nextTo"},
            "c:crossAx": {"@val": "64451712"},
            "c:crosses": {"@val": options.valAxisCrossAtMaxCategory ? "max" : "autoZero"},
            "c:crossBetween": {"@val": "between"}
          }
        },
        "c:legend": {
          "c:legendPos": {"@val": "r"},
          "c:layout": {}
        },
        "c:plotVisOnly": {"@val": "1"}
      },
      "c:txPr": {
        "a:bodyPr": {},
        "a:lstStyle": {},
        "a:p": {
          "a:pPr": {
            "a:defRPr": {"@sz": "1800"}
          },
          "a:endParaRPr": {"@lang": "en-US"}
        }
      },
      "c:externalData": {"@r:id": "rId1"}
    }
  };
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.charts.doughnut"></a>[function <span class="apidocSignatureSpan">officegen.charts.</span>doughnut (options)](#apidoc.element.officegen.charts.doughnut)
- description and source-code
```javascript
doughnut = function (options) {
  options = options || {};
  return {
    "c:chartSpace": {
      "@xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "@xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "@xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "c:lang": {"@val": "en-US"},
      "c:chart": {
        "c:title": {
          "c:layout": {}
        },
        "c:plotArea": {
          "c:layout": {},
          "c:doughnutChart": {
            "c:varyColors": {"@val": "1"},
            "c:firstSliceAng": {"@val": "0"},
            "c:holeSize": {"@val": "75"},
            "#list": []
          }
        },

        "c:legend": {
          "c:legendPos": {"@val": "r"},
          "c:layout": {}
        },
        "c:plotVisOnly": {"@val": "1"}
      },
      "c:txPr": {
        "a:bodyPr": {},
        "a:lstStyle": {},
        "a:p": {
          "a:pPr": {
            "a:defRPr": {"@sz": "1800"}
          },
          "a:endParaRPr": {"@lang": "en-US"}
        }
      },
      "c:externalData": {"@r:id": "rId1"}
    }
  };
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.charts.group-bar"></a>[function <span class="apidocSignatureSpan">officegen.charts.</span>group-bar (options)](#apidoc.element.officegen.charts.group-bar)
- description and source-code
```javascript
group-bar = function (options) {
  options = options || {};
  return {
    "c:chartSpace": {
      "@xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "@xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "@xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "c:lang": {"@val": "en-US"},
      "c:date1904": {"@val": "1"},
      "c:chart": {
        "c:plotArea": {
          "c:layout": {},
          "c:barChart": {
            "c:barDir": {"@val": "bar"},
            "c:grouping": {"@val": "stacked"},
            "c:overlap": {"@val": options.overlap || "100"},
            "c:gapWidth": {"@val": options.gapWidth || "150"},
            "#list": [
              {"c:axId": {"@val": "64451712"}},
              {"c:axId": {"@val": "64453248"}}
            ]
          },
          "c:catAx": {
            "c:axId": {"@val": "64451712"},
            "c:scaling": {
              "c:orientation": {"@val": options.catAxisReverseOrder ? "maxMin" : "minMax"}
            },
            "c:axPos": {"@val": "l"},
            "c:tickLblPos": {"@val": "nextTo"},
            "c:crossAx": {"@val": "64453248"},
            "c:crosses": {"@val": "autoZero"},
            "c:auto": {"@val": "1"},
            "c:lblAlgn": {"@val": "ctr"},
            "c:lblOffset": {"@val": "100"}
          },
          "c:valAx": {
            "c:axId": {"@val": "64453248"},
            "c:scaling": {
              "c:orientation": {"@val": "minMax"}
            },
            "c:axPos": {"@val": "b"},
//              "c:majorGridlines": {},
            "c:numFmt": {
              "@formatCode": "General",
              "@sourceLinked": "1"
            },
            "c:tickLblPos": {"@val": "nextTo"},
            "c:crossAx": {"@val": "64451712"},
            "c:crosses": {"@val": options.valAxisCrossAtMaxCategory ? "max" : "autoZero"},
            "c:crossBetween": {"@val": "between"}
          }
        },
        "c:legend": {
          "c:legendPos": {"@val": "r"},
          "c:layout": {}
        },
        "c:plotVisOnly": {"@val": "1"}
      },
      "c:txPr": {
        "a:bodyPr": {},
        "a:lstStyle": {},
        "a:p": {
          "a:pPr": {
            "a:defRPr": {"@sz": "1800"}
          },
          "a:endParaRPr": {"@lang": "en-US"}
        }
      },
      "c:externalData": {"@r:id": "rId1"}
    }
  };
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.charts.line"></a>[function <span class="apidocSignatureSpan">officegen.charts.</span>line (options)](#apidoc.element.officegen.charts.line)
- description and source-code
```javascript
line = function (options) {
  options = options || {};
  return {
    "c:chartSpace": {
      "@xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "@xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "@xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "c:lang": {"@val": "en-US"},
      "c:chart": {
        "c:plotArea": {
          "c:layout": {},
          "c:lineChart": {
            "c:grouping": {"@val": "standard"},
            "#list": [
              {"c:axId": {"@val": "64451712"}},
              {"c:axId": {"@val": "64453248"}}
            ]
          },
          "c:catAx": {
            "c:axId": {"@val": "64451712"},
            "c:scaling": {
              "c:orientation": {"@val": options.catAxisReverseOrder ? "maxMin" : "minMax"}
            },
            "c:axPos": {"@val": "l"},
            "c:tickLblPos": {"@val": "nextTo"},
            "c:crossAx": {"@val": "64453248"},
            "c:crosses": {"@val": "autoZero"},
            "c:auto": {"@val": "1"},
            "c:lblAlgn": {"@val": "ctr"},
            "c:lblOffset": {"@val": "100"}
          },
          "c:valAx": {
            "c:axId": {"@val": "64453248"},
            "c:scaling": {
              "c:orientation": {"@val": "minMax"}
            },
            "c:axPos": {"@val": "b"},
//              "c:majorGridlines": {},
            "c:numFmt": {
              "@formatCode": "General",
              "@sourceLinked": "1"
            },
            "c:tickLblPos": {"@val": "nextTo"},
            "c:crossAx": {"@val": "64451712"},
            "c:crosses": {"@val": options.valAxisCrossAtMaxCategory ? "max" : "autoZero"},
            "c:crossBetween": {"@val": "between"}
          }
        },
        "c:legend": {
          "c:legendPos": {"@val": "r"},
          "c:layout": {}
        },
        "c:plotVisOnly": {"@val": "1"}
      },
      "c:txPr": {
        "a:bodyPr": {},
        "a:lstStyle": {},
        "a:p": {
          "a:pPr": {
            "a:defRPr": {"@sz": "1800"}
          },
          "a:endParaRPr": {"@lang": "en-US"}
        }
      },
      "c:externalData": {"@r:id": "rId1"}
    }
  }
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.charts.null"></a>[function <span class="apidocSignatureSpan">officegen.charts.</span>null (options)](#apidoc.element.officegen.charts.null)
- description and source-code
```javascript
null = function (options) {
  return {
    "c:chartSpace": {
      "@xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "@xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "@xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "c:lang": {"@val": "en-US"},
      "c:date1904": {"@val": "1"},
      "c:chart": {}
    }
  };
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.charts.pie"></a>[function <span class="apidocSignatureSpan">officegen.charts.</span>pie (options)](#apidoc.element.officegen.charts.pie)
- description and source-code
```javascript
pie = function (options) {
  options = options || {};
  return {
    "c:chartSpace": {
      "@xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "@xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "@xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "c:lang": {"@val": "en-US"},
      "c:chart": {
        "c:title": {
          "c:layout": {}
        },
        "c:plotArea": {
          "c:layout": {},
          "c:pieChart": {
            "c:varyColors": {"@val": "1"},
            "c:firstSliceAng": {"@val": "0"},
            "#list": []
          }
        },

        "c:legend": {
          "c:legendPos": {"@val": "r"},
          "c:layout": {}
        },
        "c:plotVisOnly": {"@val": "1"}
      },
      "c:txPr": {
        "a:bodyPr": {},
        "a:lstStyle": {},
        "a:p": {
          "a:pPr": {
            "a:defRPr": {"@sz": "1800"}
          },
          "a:endParaRPr": {"@lang": "en-US"}
        }
      },
      "c:externalData": {"@r:id": "rId1"}
    }
  }
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.charts.stacked-column"></a>[function <span class="apidocSignatureSpan">officegen.charts.</span>stacked-column (options)](#apidoc.element.officegen.charts.stacked-column)
- description and source-code
```javascript
stacked-column = function (options) {
  options = options || {};
  return  {
    "c:chartSpace": {
      "@xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "@xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "@xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "c:lang": { "@val": "en-US" },
      "c:date1904": { "@val": "1" },
      "c:chart": {
        "c:plotArea": {
          "c:layout": {},
          "c:barChart": {
            "c:barDir": { "@val": "col" },
            "c:grouping": { "@val": "stacked" },
            "c:overlap": { "@val": options.overlap || "100" },
            "c:gapWidth": { "@val": options.gapWidth || "25" },

            "#list": [
              {"c:axId": { "@val": "64451712" }},
              {"c:axId": { "@val": "64453248" }}
            ]
          },
          "c:catAx": {
            "c:axId": { "@val": "64451712" },
            "c:scaling": {
              "c:orientation": { "@val": options.catAxisReverseOrder ? "maxMin" : "minMax" }
            },
            "c:axPos": { "@val": "l" },
            "c:tickLblPos": { "@val": "nextTo" },
            "c:crossAx": { "@val": "64453248" },
            "c:crosses": { "@val": "autoZero" },
            "c:auto": { "@val": "1" },
            "c:lblAlgn": { "@val": "ctr" },
            "c:lblOffset": { "@val": "100" }
          },
          "c:valAx": {
            "c:axId": { "@val": "64453248" },
            "c:scaling": {
              "c:orientation": { "@val": "minMax" }
            },
            "c:axPos": { "@val": "b" },
//              "c:majorGridlines": {},
            "c:numFmt": {
              "@formatCode": "General",
              "@sourceLinked": "1"
            },
            "c:tickLblPos": { "@val": "nextTo" },
            "c:crossAx": { "@val": "64451712" },
            "c:crosses": { "@val": options.valAxisCrossAtMaxCategory ? "max" : "autoZero"},
            "c:crossBetween": { "@val": "between" }
          }
        },
        "c:legend": {
          "c:legendPos": { "@val": "r" },
          "c:layout": {}
        },
        "c:plotVisOnly": { "@val": "1" }
      },
      "c:txPr": {
        "a:bodyPr": {},
        "a:lstStyle": {},
        "a:p": {
          "a:pPr": {
            "a:defRPr": { "@sz": "1800" }
          },
          "a:endParaRPr": { "@lang": "en-US" }
        }
      },
      "c:externalData": { "@r:id": "rId1" }
    }
  };
}
```
- example usage
```shell
n/a
```



# <a name="apidoc.module.officegen.docplug"></a>[module officegen.docplug](#apidoc.module.officegen.docplug)

#### <a name="apidoc.element.officegen.docplug.docplug"></a>[function <span class="apidocSignatureSpan">officegen.</span>docplug ( genobj, genPrivate, docType, defValuesFunc )](#apidoc.element.officegen.docplug.docplug)
- description and source-code
```javascript
function MakeDocPluginapi( genobj, genPrivate, docType, defValuesFunc ) {
	// Save everything because we'll need it later:
	this.docType = docType;
	this.genPrivate = genPrivate;
	this.ogPluginsApi = genPrivate.plugs;
	this.genobj = genobj;
	this.defValuesFunc = defValuesFunc;
	this.callbacksList = {};
	this.plugsList = [];

	// Here we can put anything specific to this document type BUT it's not data so if we'll clear the data inside the document we'
ll not need to clear anything here:
	this.genPrivate.features.type[this.docType] = this;

	// Prepare the object that we'll use to store data specific to this document type:
	this.genPrivate.type[this.docType] = {};
	if ( typeof defValuesFunc === 'function' ) {
		// Put the default data:
		defValuesFunc ( this );
	} // Endif.

	// Catch the clear document data event and connect it to us:
	genobj.on ( 'clearDocType', clearDocData ( this ) );
	return this;
}
```
- example usage
```shell
n/a
```



# <a name="apidoc.module.officegen.docplug.prototype"></a>[module officegen.docplug.prototype](#apidoc.module.officegen.docplug.prototype)

#### <a name="apidoc.element.officegen.docplug.prototype.emitEvent"></a>[function <span class="apidocSignatureSpan">officegen.docplug.prototype.</span>emitEvent ( eventType, eventData )](#apidoc.element.officegen.docplug.prototype.emitEvent)
- description and source-code
```javascript
emitEvent = function ( eventType, eventData ) {
	var funcThis = this;

	// We'll do something only if we have this type of event:
	if ( this.callbacksList[eventType] ) {
		this.callbacksList[eventType].forEach ( function ( value ) {
			if ( typeof value === 'function' ) {
				value ( eventData, eventType, funcThis );
			} // Endif.
		});
	} // Endif.
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.docplug.prototype.getDataStorage"></a>[function <span class="apidocSignatureSpan">officegen.docplug.prototype.</span>getDataStorage ()](#apidoc.element.officegen.docplug.prototype.getDataStorage)
- description and source-code
```javascript
getDataStorage = function () {
	return this.genPrivate.type[this.docType];
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.docplug.prototype.getFeaturesStorage"></a>[function <span class="apidocSignatureSpan">officegen.docplug.prototype.</span>getFeaturesStorage ()](#apidoc.element.officegen.docplug.prototype.getFeaturesStorage)
- description and source-code
```javascript
getFeaturesStorage = function () {
	return this.genPrivate.features.type[this.docType];
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.docplug.prototype.getVerboseMode"></a>[function <span class="apidocSignatureSpan">officegen.docplug.prototype.</span>getVerboseMode ( moduleName )](#apidoc.element.officegen.docplug.prototype.getVerboseMode)
- description and source-code
```javascript
getVerboseMode = function ( moduleName ) {
	return this.ogPluginsApi.getVerboseMode ( this.docType, moduleName );
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.docplug.prototype.logIfVerbose"></a>[function <span class="apidocSignatureSpan">officegen.docplug.prototype.</span>logIfVerbose ( message )](#apidoc.element.officegen.docplug.prototype.logIfVerbose)
- description and source-code
```javascript
logIfVerbose = function ( message ) {
	if ( this.getVerboseMode () ) {
		console.log ( message );
	} // Endif.
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.docplug.prototype.registerCallback"></a>[function <span class="apidocSignatureSpan">officegen.docplug.prototype.</span>registerCallback ( eventType, cbFunc )](#apidoc.element.officegen.docplug.prototype.registerCallback)
- description and source-code
```javascript
registerCallback = function ( eventType, cbFunc ) {
	// First make sure that we have a list of callbacks for this event type:
	if ( !this.callbacksList[eventType] ) {
		this.callbacksList[eventType] = [];
	} // Endif.

	// Now we'll just push this callback:
	this.callbacksList[eventType].push ( cbFunc );
}
```
- example usage
```shell
n/a
```



# <a name="apidoc.module.officegen.docx_p"></a>[module officegen.docx_p](#apidoc.module.officegen.docx_p)

#### <a name="apidoc.element.officegen.docx_p.docx_p"></a>[function <span class="apidocSignatureSpan">officegen.</span>docx_p ( genobj, genPrivate, docType, dataContainer, extraSettings, options )](#apidoc.element.officegen.docx_p.docx_p)
- description and source-code
```javascript
function MakeDocxP( genobj, genPrivate, docType, dataContainer, extraSettings, options ) {
	// Save everything because we'll need it later:
	this.docType = docType;
	this.genPrivate = genPrivate;
	this.ogPluginsApi = genPrivate.plugs; // Generic officegen API for plugins.
	this.msPluginsApi = genPrivate.plugs.type.msoffice; // msoffice plugins API.
	this.genobj = genobj;
	this.data = [];
	this.extraSettings = extraSettings || {};
	this.options = options || {};

	this.mainPath = genPrivate.features.type.msoffice.main_path; // The "folder" name inside the document zip that all the specific
 resources of this document type are stored.
	this.mainPathFile = genPrivate.features.type.msoffice.main_path_file; // The name of the main real xml resource of this document
.
	this.relsMain = genPrivate.type.msoffice.rels_main; // Main rels file.
	this.relsApp = genPrivate.type.msoffice.rels_app; // Main rels file inside the specific document type "folder".
	this.filesList = genPrivate.type.msoffice.files_list; // Resources list xml.
	this.srcFilesList = genPrivate.type.msoffice.src_files_list; // For storing extra files inside the document zip.


	if ( dataContainer ) {
		dataContainer.push ( this );
	} // Endif.

	return this;
}
```
- example usage
```shell
n/a
```



# <a name="apidoc.module.officegen.docx_p.prototype"></a>[module officegen.docx_p.prototype](#apidoc.module.officegen.docx_p.prototype)

#### <a name="apidoc.element.officegen.docx_p.prototype.addHorizontalLine"></a>[function <span class="apidocSignatureSpan">officegen.docx_p.prototype.</span>addHorizontalLine ()](#apidoc.element.officegen.docx_p.prototype.addHorizontalLine)
- description and source-code
```javascript
addHorizontalLine = function () {
	var newP = this;
	newP.data[newP.data.length] = { 'horizontal_line': true };
}
```
- example usage
```shell
...
			case "text":
				newP.addText(data.val, data.opt);
				break;
			case "linebreak":
				newP.addLineBreak();
				break;
			case "horizontalline":
				newP.addHorizontalLine();
				break;
			case "image":
				// Improved by peizhuang in Aug 2016 (added data.opt):
				// data.imagetype been added by Ziv Barber in Aug 2016.
				newP.addImage(data.path, data.opt || {}, data.imagetype );
				break;
			case "pagebreak":
...
```

#### <a name="apidoc.element.officegen.docx_p.prototype.addImage"></a>[function <span class="apidocSignatureSpan">officegen.docx_p.prototype.</span>addImage ( image_path, opt, image_format_type )](#apidoc.element.officegen.docx_p.prototype.addImage)
- description and source-code
```javascript
addImage = function ( image_path, opt, image_format_type ) {
	var newP = this;
	var image_type = (typeof image_format_type == 'string') ? image_format_type : 'png';
	var defWidth = 320;
	var defHeight = 200;

	if ( typeof image_path == 'string' ) {
		var ret_data = fast_image_size ( image_path );
		if ( ret_data.type == 'unknown' ) {
			var image_ext = path.extname ( image_path );

			switch ( image_ext ) {
				case '.bmp':
					image_type = 'bmp';
					break;

				case '.gif':
					image_type = 'gif';
					break;

				case '.jpg':
				case '.jpeg':
					image_type = 'jpeg';
					break;

				case '.emf':
					image_type = 'emf';
					break;

				case '.tiff':
					image_type = 'tiff';
					break;
			} // End of switch.

		} else {
			if ( ret_data.width ) {
				defWidth = ret_data.width;
			} // Endif.

			if ( ret_data.height ) {
				defHeight = ret_data.height;
			} // Endif.

			image_type = ret_data.type;
			if ( image_type == 'jpg' ) {
				image_type = 'jpeg';
			} // Endif.
		} // Endif.
	} // Endif.

	var objNum = newP.data.length;
	newP.data[objNum] = { image: image_path, options: opt || {} };

	if ( !newP.data[objNum].options.cx && defWidth ) {
		newP.data[objNum].options.cx = defWidth;
	} // Endif.

	if ( !newP.data[objNum].options.cy && defHeight ) {
		newP.data[objNum].options.cy = defHeight;
	} // Endif.

	var image_id = newP.srcFilesList.indexOf ( image_path );
	var image_rel_id = -1;

	if ( image_id >= 0 ) {
		for ( var j = 0, total_size_j = newP.relsApp.length; j < total_size_j; j++ ) {
			if ( newP.relsApp[j].target == ('media/image' + (image_id + 1) + '.' + image_type) ) {
				image_rel_id = j + 1;
			} // Endif.
		} // Endif.

	} else
	{
		image_id = newP.srcFilesList.length;
		newP.srcFilesList[image_id] = image_path;
		newP.ogPluginsApi.intAddAnyResourceToParse ( newP.mainPath + '\\media\\image' + (image_id + 1) + '.' + image_type, (typeof image_path
 == 'string') ? 'file' : 'stream', image_path, null, false );
	} // Endif.

	if ( image_rel_id == -1 ) {
		image_rel_id = newP.relsApp.length + 1;

		newP.relsApp.push (
			{
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
				target: 'media/image' + (image_id + 1) + '.' + image_type,
				clear: 'data'
			}
		);
	} // Endif.

	if ((opt || {}).link) {

		var link_rel_id = newP.relsApp.length + 1;

		newP.relsApp.push({
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
			target: opt.link,
			targetMode: 'External'
		});

		newP.data[objNum].link_rel_id = link_rel_id;

	} // Endif.

	newP.data[objNum].image_id = image_id;
	newP.data[objNum].rel_id = image_rel_id;
}
```
- example usage
```shell
...
				break;
			case "horizontalline":
				newP.addHorizontalLine();
				break;
			case "image":
				// Improved by peizhuang in Aug 2016 (added data.opt):
				// data.imagetype been added by Ziv Barber in Aug 2016.
				newP.addImage(data.path, data.opt || {}, data.imagetype );
				break;
			case "pagebreak":
				newP = genobj.putPageBreak();
				break;
			case "table":
				newP = genobj.createTable(data.val, data.opt);
				break;
...
```

#### <a name="apidoc.element.officegen.docx_p.prototype.addLineBreak"></a>[function <span class="apidocSignatureSpan">officegen.docx_p.prototype.</span>addLineBreak ()](#apidoc.element.officegen.docx_p.prototype.addLineBreak)
- description and source-code
```javascript
addLineBreak = function () {
	var newP = this;
	newP.data[newP.data.length] = { 'line_break': true };
}
```
- example usage
```shell
...
	genobj.createJson = function ( data, newP ) {
		newP = newP || genobj.createP(data.lopt || {});
		switch(data.type) {
			case "text":
				newP.addText(data.val, data.opt);
				break;
			case "linebreak":
				newP.addLineBreak();
				break;
			case "horizontalline":
				newP.addHorizontalLine();
				break;
			case "image":
				// Improved by peizhuang in Aug 2016 (added data.opt):
				// data.imagetype been added by Ziv Barber in Aug 2016.
...
```

#### <a name="apidoc.element.officegen.docx_p.prototype.addText"></a>[function <span class="apidocSignatureSpan">officegen.docx_p.prototype.</span>addText ( text_msg, opt, flag_data )](#apidoc.element.officegen.docx_p.prototype.addText)
- description and source-code
```javascript
addText = function ( text_msg, opt, flag_data ) {
	var newP = this;
	var objNum = newP.data.length
	newP.data[objNum] = { text: text_msg, options: opt || {}, ext_data: flag_data };

	if ((opt || {}).link) {
		var link_rel_id = newP.relsApp.length + 1;

		newP.relsApp.push({
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
			target: opt.link,
			targetMode: 'External'
		})

		newP.data[objNum].link_rel_id = link_rel_id;
	}
}
```
- example usage
```shell
...
	 * @param {object} data ???.
	 * @param {object} newP ???.
	 */
	genobj.createJson = function ( data, newP ) {
		newP = newP || genobj.createP(data.lopt || {});
		switch(data.type) {
			case "text":
				newP.addText(data.val, data.opt);
				break;
			case "linebreak":
				newP.addLineBreak();
				break;
			case "horizontalline":
				newP.addHorizontalLine();
				break;
...
```

#### <a name="apidoc.element.officegen.docx_p.prototype.endBookmark"></a>[function <span class="apidocSignatureSpan">officegen.docx_p.prototype.</span>endBookmark ()](#apidoc.element.officegen.docx_p.prototype.endBookmark)
- description and source-code
```javascript
endBookmark = function () {
	var newP = this;
	newP.data[newP.data.length] = { 'bookmark_end': true };
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.docx_p.prototype.startBookmark"></a>[function <span class="apidocSignatureSpan">officegen.docx_p.prototype.</span>startBookmark ( anchorName )](#apidoc.element.officegen.docx_p.prototype.startBookmark)
- description and source-code
```javascript
startBookmark = function ( anchorName ) {
	var newP = this;
	newP.data[newP.data.length] = { 'bookmark_start': anchorName };
}
```
- example usage
```shell
n/a
```



# <a name="apidoc.module.officegen.docxplg_headfoot"></a>[module officegen.docxplg_headfoot](#apidoc.module.officegen.docxplg_headfoot)

#### <a name="apidoc.element.officegen.docxplg_headfoot.docxplg_headfoot"></a>[function <span class="apidocSignatureSpan">officegen.</span>docxplg_headfoot ( pluginsman )](#apidoc.element.officegen.docxplg_headfoot.docxplg_headfoot)
- description and source-code
```javascript
function MakeHeadfootPlugin( pluginsman ) {
	var funcThis = this;

	// You can change it if you want to support more types, since that the Word and Excel document generators also supporting very
similar plugins:
	if ( pluginsman.docType !== 'docx' ) {
		throw "[docx-headfoot] This plugin supporting only Word based documents.";
	} // Endif.

	this.ogPluginsApi = pluginsman.ogPluginsApi; // Generic officegen API for plugins.
	this.msPluginsApi = pluginsman.genPrivate.plugs.type.msoffice; // msoffice plugins API.
	this.pluginsman = pluginsman; // Document type specific plugins API.

	this.docxData = pluginsman.getDataStorage (); // Here you can store any temporary data needed for generating the document and depending
 on the data filled by the user.

	this.mainPath = pluginsman.genPrivate.features.type.msoffice.main_path; // The "folder" name inside the document zip that all the
 specific resources of this document type are stored.
	this.mainPathFile = pluginsman.genPrivate.features.type.msoffice.main_path_file; // The name of the main real xml resource of this
 document.
	this.relsMain = pluginsman.genPrivate.type.msoffice.rels_main; // Main rels file.
	this.relsApp = pluginsman.genPrivate.type.msoffice.rels_app; // Main rels file inside the specific document type "folder".
	this.filesList = pluginsman.genPrivate.type.msoffice.files_list; // Resources list xml.
	this.srcFilesList = pluginsman.genPrivate.type.msoffice.src_files_list; // For storing extra files inside the document zip.

	//
	// Catch events inside the document so we can extent it:
	//

	// We want to extend the main API of the docx document object:
	pluginsman.registerCallback ( 'makeDocApi', function ( docObj ) { funcThis.extenddocxApi ( docObj ); } );

	// This event tell us that the generator is about to start working:
	pluginsman.registerCallback ( 'beforeGen', function ( docObj ) { funcThis.beforeGen ( docObj ); } );

	return this;
}
```
- example usage
```shell
n/a
```



# <a name="apidoc.module.officegen.docxplg_headfoot.prototype"></a>[module officegen.docxplg_headfoot.prototype](#apidoc.module.officegen.docxplg_headfoot.prototype)

#### <a name="apidoc.element.officegen.docxplg_headfoot.prototype.beforeGen"></a>[function <span class="apidocSignatureSpan">officegen.docxplg_headfoot.prototype.</span>beforeGen ( docObj )](#apidoc.element.officegen.docxplg_headfoot.prototype.beforeGen)
- description and source-code
```javascript
beforeGen = function ( docObj ) {
	var funcThis = this;

	var isNeedHeadFoot = false;
	if ( funcThis.docxData.docHeader || funcThis.docxData.docHeadereven || funcThis.docxData.docHeaderfirst || funcThis.docxData.docFooter
 || funcThis.docxData.docFootereven || funcThis.docxData.docFooterfirst ) {
		isNeedHeadFoot = true;
	} // Endif.

	if ( isNeedHeadFoot ) {
		// Create the endnotes resource:
		this.ogPluginsApi.intAddAnyResourceToParse ( this.mainPath + '\\endnotes.xml', 'buffer', null, function ( data ) { return funcThis
.cbEndNotes ( data ) }, false );

		// Create the footnotes resource:
		this.ogPluginsApi.intAddAnyResourceToParse ( this.mainPath + '\\footnotes.xml', 'buffer', null, function ( data ) { return funcThis
.cbFootNotes ( data ) }, false );

		// Create header1.xml:
		this.ogPluginsApi.intAddAnyResourceToParse ( this.mainPath + '\\header1.xml', 'buffer', funcThis.pluginsman.genobj.getHeader ( '
even' ), funcThis.pluginsman.genobj.cbMakeDocxDocument, false );

		// Create header2.xml:
		this.ogPluginsApi.intAddAnyResourceToParse ( this.mainPath + '\\header2.xml', 'buffer', funcThis.pluginsman.genobj.getHeader (),
funcThis.pluginsman.genobj.cbMakeDocxDocument, false );

		// Create header3.xml:
		this.ogPluginsApi.intAddAnyResourceToParse ( this.mainPath + '\\header3.xml', 'buffer', funcThis.pluginsman.genobj.getHeader ( '
first' ), funcThis.pluginsman.genobj.cbMakeDocxDocument, false );

		// Create footer1.xml:
		this.ogPluginsApi.intAddAnyResourceToParse ( this.mainPath + '\\footer1.xml', 'buffer', funcThis.pluginsman.genobj.getFooter ( '
even' ), funcThis.pluginsman.genobj.cbMakeDocxDocument, false );

		// Create footer2.xml:
		this.ogPluginsApi.intAddAnyResourceToParse ( this.mainPath + '\\footer2.xml', 'buffer', funcThis.pluginsman.genobj.getFooter (),
funcThis.pluginsman.genobj.cbMakeDocxDocument, false );

		// Create footer3.xml:
		this.ogPluginsApi.intAddAnyResourceToParse ( this.mainPath + '\\footer3.xml', 'buffer', funcThis.pluginsman.genobj.getFooter ( '
first' ), funcThis.pluginsman.genobj.cbMakeDocxDocument, false );

		// Add a rel entry to the main docx rels:

		this.relsApp.push ({
				target: 'header1.xml',
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
				clear: 'generate' // Placing 'generate' here means that officegen will destroy this entry in the files list after finishing
to generate the document.
		});
		this.docxData.header1RelId = this.relsApp.length;

		this.relsApp.push ({
				target: 'header2.xml',
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
				clear: 'generate' // Placing 'generate' here means that officegen will destroy this entry in the files list after finishing
to generate the document.
		});
		this.docxData.header2RelId = this.relsApp.length;

		this.relsApp.push ({
				target: 'header3.xml',
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
				clear: 'generate' // Placing 'generate' here means that officegen will destroy this entry in the files list after finishing
to generate the document.
		});
		this.docxData.header3RelId = this.relsApp.length;

		this.relsApp.push ({
				target: 'footer1.xml',
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
				clear: 'generate' // Placing 'generate' here means that officegen will destroy this entry in the files list after finishing
to generate the document.
		});
		this.docxData.footer1RelId = this.relsApp.length;

		this.relsApp.push ({
				target: 'footer2.xml',
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
				clear: 'generate' // Placing 'generate' here means that officegen will destroy this entry in the files list after finishing
to generate the document.
		});
		this.docxData.footer2RelId = this.relsApp.length;

		this.relsApp.push ({
				target: 'footer3.xml',
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
				clear: 'generat ...
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.docxplg_headfoot.prototype.cbEndNotes"></a>[function <span class="apidocSignatureSpan">officegen.docxplg_headfoot.prototype.</span>cbEndNotes ( data )](#apidoc.element.officegen.docxplg_headfoot.prototype.cbEndNotes)
- description and source-code
```javascript
cbEndNotes = function ( data ) {
	return this.msPluginsApi.cbMakeMsOfficeBasicXml ( data ) + '<w:endnotes xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility
/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships
" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://
schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http
://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><w:endnote
 w:type="separator" w:id="-1"><w:p w:rsidR="002E67D7" w:rsidRDefault="002E67D7" w:rsidP="009F2180"><w:pPr><w:spacing w:after="0"
w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id
="0"><w:p w:rsidR="002E67D7" w:rsidRDefault="002E67D7" w:rsidP="009F2180"><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="
auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:endnote></w:endnotes>';
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.docxplg_headfoot.prototype.cbFootNotes"></a>[function <span class="apidocSignatureSpan">officegen.docxplg_headfoot.prototype.</span>cbFootNotes ( data )](#apidoc.element.officegen.docxplg_headfoot.prototype.cbFootNotes)
- description and source-code
```javascript
cbFootNotes = function ( data ) {
	return this.msPluginsApi.cbMakeMsOfficeBasicXml ( data ) + '<w:footnotes xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility
/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships
" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://
schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http
://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><w:footnote
 w:type="separator" w:id="-1"><w:p w:rsidR="002E67D7" w:rsidRDefault="002E67D7" w:rsidP="009F2180"><w:pPr><w:spacing w:after="0"
w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w
:id="0"><w:p w:rsidR="002E67D7" w:rsidRDefault="002E67D7" w:rsidP="009F2180"><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule
="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote></w:footnotes>';
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.docxplg_headfoot.prototype.extenddocxApi"></a>[function <span class="apidocSignatureSpan">officegen.docxplg_headfoot.prototype.</span>extenddocxApi ( docObj )](#apidoc.element.officegen.docxplg_headfoot.prototype.extenddocxApi)
- description and source-code
```javascript
extenddocxApi = function ( docObj ) {
	var funcThis = this;

	docObj.getHeader = function ( objSubType ) {
		if ( objSubType ) {
			if ( objSubType !== 'even' && objSubType !== 'first' ) {
				throw 'objSubType can be either "even", "first" or false value';
			} // Endif.
		} // Endif.

		objSubType = 'docHeader' + (objSubType || '');
		if ( funcThis.docxData[objSubType] ) {
			return funcThis.docxData[objSubType];
		} // Endif.

		funcThis.docxData[objSubType] = new MakeHeadFootObj ( funcThis, 'hdr', 'Header' );
		return funcThis.docxData[objSubType];
	};

	docObj.getFooter = function ( objSubType ) {
		if ( objSubType ) {
			if ( objSubType !== 'even' && objSubType !== 'first' ) {
				throw 'objSubType can be either "even", "first" or false value';
			} // Endif.
		} // Endif.

		objSubType = 'docFooter' + (objSubType || '');
		if ( funcThis.docxData[objSubType] ) {
			return funcThis.docxData[objSubType];
		} // Endif.

		funcThis.docxData[objSubType] = new MakeHeadFootObj ( funcThis, 'ftr', 'Footer' );
		return funcThis.docxData[objSubType];
	};
}
```
- example usage
```shell
n/a
```



# <a name="apidoc.module.officegen.docxtable"></a>[module officegen.docxtable](#apidoc.module.officegen.docxtable)

#### <a name="apidoc.element.officegen.docxtable._getBase"></a>[function <span class="apidocSignatureSpan">officegen.docxtable.</span>_getBase (rowSpecs, colSpecs, opts)](#apidoc.element.officegen.docxtable._getBase)
- description and source-code
```javascript
_getBase = function (rowSpecs, colSpecs, opts) {
  var self = this;

  var baseTable = {
    "w:tbl": {
      "w:tblPr": {
        "w:tblStyle": {
          "@w:val": "a3"
        },
        "w:tblW": {
          "@w:w": "0",
          "@w:type": "auto"
        },
        "w:tblLook": {
          "@w:val": "04A0",
          "@w:firstRow": "1",
          "@w:lastRow": "0",
          "@w:firstColumn": "1",
          "@w:lastColumn": "0",
          "@w:noHBand": "0",
          "@w:noVBand": "1"
        }
      },
      "w:tblGrid": {
        "#list": colSpecs
      },
      "#list": [rowSpecs]
    }
  };
  if (opts.borders) {
    baseTable["w:tbl"]["w:tblPr"]["w:tblBorders"] = {
      "w:top": {
        "@w:val": "single",
        "@w:sz": "12",
        "@w:space": "0",
        "@w:color": "000000"
      },
      "w:bottom": {
        "@w:val": "single",
        "@w:sz": "12",
        "@w:space": "0",
        "@w:color": "000000"
      },
      "w:left": {
        "@w:val": "single",
        "@w:sz": "12",
        "@w:space": "0",
        "@w:color": "000000"
      },
      "w:right": {
        "@w:val": "single",
        "@w:sz": "12",
        "@w:space": "0",
        "@w:color": "000000"
      },
      "w:insideH": {
        "@w:val": "single",
        "@w:sz": "12",
        "@w:space": "0",
        "@w:color": "000000"
      },
      "w:insideV": {
        "@w:val": "single",
        "@w:sz": "12",
        "@w:space": "0",
        "@w:color": "000000"
      }
    };
  }
  return baseTable;
}
```
- example usage
```shell
...

  // assume passed in an array of row objects
  getTable: function(rows, tblOpts) {
tblOpts = tblOpts || {};

var self = this;

return self._getBase(
  rows.map(function(row) {
    return self._getRow(
      row.map(function(cell) {
        cell = cell || {};
        if (typeof cell === 'string' || typeof cell === 'number') {
          var val = cell;
          cell = {
...
```

#### <a name="apidoc.element.officegen.docxtable._getCell"></a>[function <span class="apidocSignatureSpan">officegen.docxtable.</span>_getCell (val, opts, tblOpts)](#apidoc.element.officegen.docxtable._getCell)
- description and source-code
```javascript
_getCell = function (val, opts, tblOpts) {
  opts = opts || {};
  // var b = {};

  // if (opts.b) {
  //   b = {
  //     "w:tc": {
  //       "w:p": {
  //         "w:r": {
  //           "w:rPr": {
  //             "w:b": {}
  //           }
  //         }
  //       }
  //     }
  //   }
  // }

  var cellObj = {
    "w:tc": {
      "w:tcPr": {
        "w:tcW": {
          "@w:w": opts.cellColWidth || tblOpts.tableColWidth || "0",
          "@w:type": "dxa"
        },
        "w:gridSpan": {
          "@w:val" : opts.gridSpan || "1"
        },
        "w:vAlign": {
          "@w:val": opts.vAlign || "top"
        },
        "w:shd": {
          "@w:val": "clear",
          "@w:color": "auto",
          "@w:fill": opts.shd && opts.shd.fill || "",
          "@w:themeFill": opts.shd && opts.shd.themeFill || "",
          "@w:themeFillTint": opts.shd && opts.shd.themeFillTint || ""
        }
      },
      "w:p": {
        "@w:rsidR": "00995B51",
        "@w:rsidRPr": "00722E63",
        "@w:rsidRDefault": "00995B51",
        "w:pPr": {
          "w:jc": {
            "@w:val": opts.align || tblOpts.tableAlign || "center"
          },
          // "w:rPr": {
          //   "w:rFonts": {
          //     "@w:asciiTheme": "majorEastAsia",
          //     "@w:eastAsiaTheme": "majorEastAsia",
          //     "@w:hAnsiTheme": "majorEastAsia"
          //   },
          //   // "w:b": {},
          //   "w:sz": {
          //     "@w:val": "24"
          //   },
          //   "w:szCs": {
          //     "@w:val": "24"
          //   }
          // }
        },
        "w:r": {
          "@w:rsidRPr": "00722E63",
          "w:rPr": {
            "w:rFonts": {
              "@w:ascii": opts.fontFamily || tblOpts.tableFontFamily || "宋体",
              "@w:hAnsi": opts.fontFamily || tblOpts.tableFontFamily || "宋体"
            },
            "w:color": {
              "@w:val": opts.color || tblOpts.tableColor || "000"
            },
            "w:b": {},
            "w:sz": {
              "@w:val": opts.sz || tblOpts.sz || "24"
            },
            "w:szCs": {
              "@w:val": opts.sz || tblOpts.sz || "24"
            }
          },
          "w:t": val
        }
      }
    }
  };

  if (!opts.b) {
    delete cellObj["w:tc"]["w:p"]["w:r"]["w:rPr"]["w:b"];
  }

  return cellObj;
}
```
- example usage
```shell
...
        if (typeof cell === 'string' || typeof cell === 'number') {
          var val = cell;
          cell = {
            val: val
          };
        }

        return self._getCell(cell.val, cell.opts, tblOpts);
      }),
      tblOpts
    );
  }),
  self._getColSpecs(rows, tblOpts),
  tblOpts
);
...
```

#### <a name="apidoc.element.officegen.docxtable._getColSpecs"></a>[function <span class="apidocSignatureSpan">officegen.docxtable.</span>_getColSpecs (cols, opts)](#apidoc.element.officegen.docxtable._getColSpecs)
- description and source-code
```javascript
_getColSpecs = function (cols, opts) {
  var self = this;
  return cols[0].map(function(val, idx) {
    return self._tblGrid(opts);
  });
}
```
- example usage
```shell
...
          }

          return self._getCell(cell.val, cell.opts, tblOpts);
        }),
        tblOpts
      );
    }),
    self._getColSpecs(rows, tblOpts),
    tblOpts
  );
},

_getBase: function(rowSpecs, colSpecs, opts) {
  var self = this;
...
```

#### <a name="apidoc.element.officegen.docxtable._getRow"></a>[function <span class="apidocSignatureSpan">officegen.docxtable.</span>_getRow (cells, opts)](#apidoc.element.officegen.docxtable._getRow)
- description and source-code
```javascript
_getRow = function (cells, opts) {
  return {
    "w:tr": {
      "@w:rsidR": "00995B51",
      "@w:rsidTr": "007F1D13",
      "#list": [cells] // populate this with an array of table cell objects
    }
  };
}
```
- example usage
```shell
...
  getTable: function(rows, tblOpts) {
tblOpts = tblOpts || {};

var self = this;

return self._getBase(
  rows.map(function(row) {
    return self._getRow(
      row.map(function(cell) {
        cell = cell || {};
        if (typeof cell === 'string' || typeof cell === 'number') {
          var val = cell;
          cell = {
            val: val
          };
...
```

#### <a name="apidoc.element.officegen.docxtable._tblGrid"></a>[function <span class="apidocSignatureSpan">officegen.docxtable.</span>_tblGrid (opts)](#apidoc.element.officegen.docxtable._tblGrid)
- description and source-code
```javascript
_tblGrid = function (opts) {
  return {
    "w:gridCol": {
      "@w:w": opts.tableColWidth || "1"
    }
  };
}
```
- example usage
```shell
...
  }
  return baseTable;
},

_getColSpecs: function(cols, opts) {
  var self = this;
  return cols[0].map(function(val, idx) {
    return self._tblGrid(opts);
  });
},

// TODO
_tblGrid: function(opts) {
  return {
    "w:gridCol": {
...
```

#### <a name="apidoc.element.officegen.docxtable.getTable"></a>[function <span class="apidocSignatureSpan">officegen.docxtable.</span>getTable (rows, tblOpts)](#apidoc.element.officegen.docxtable.getTable)
- description and source-code
```javascript
getTable = function (rows, tblOpts) {
  tblOpts = tblOpts || {};

  var self = this;

  return self._getBase(
    rows.map(function(row) {
      return self._getRow(
        row.map(function(cell) {
          cell = cell || {};
          if (typeof cell === 'string' || typeof cell === 'number') {
            var val = cell;
            cell = {
              val: val
            };
          }

          return self._getCell(cell.val, cell.opts, tblOpts);
        }),
        tblOpts
      );
    }),
    self._getColSpecs(rows, tblOpts),
    tblOpts
  );
}
```
- example usage
```shell
...
			outString += '</w:p>';
		} // Endif.

		// Work on all the stored paragraphs inside this document:
		for ( var i = 0, total_size = objs_list.length; i < total_size; i++ ) {

			if (objs_list[i] && objs_list[i].type === 'table') {
				var table_obj = docxTable.getTable(objs_list[i].data, objs_list[i].options);
          		var table_xml = xmlBuilder.create(table_obj,{version: '1.0', encoding: 'UTF-8', standalone: true}).toString({ pretty
: true, indent: '  ', newline: '\n' });
	          	outString += table_xml;
	          	continue;
			}

			outString += '<w:p w:rsidR="00A77427" w:rsidRDefault="007F1D13">';
			var pPrData = '';
...
```



# <a name="apidoc.module.officegen.genofficetable"></a>[module officegen.genofficetable](#apidoc.module.officegen.genofficetable)

#### <a name="apidoc.element.officegen.genofficetable._getBase"></a>[function <span class="apidocSignatureSpan">officegen.genofficetable.</span>_getBase (rowSpecs, colSpecs, options)](#apidoc.element.officegen.genofficetable._getBase)
- description and source-code
```javascript
_getBase = function (rowSpecs, colSpecs, options) {
  var self = this;

  return {
    "p:graphicFrame": {
      "p:nvGraphicFramePr": {
        "p:cNvPr": {
          "@id": "6",
          "@name": "Table 5"
        },
        "p:cNvGraphicFramePr": {
          "a:graphicFrameLocks": {
            "@noGrp": "1"
          }
        },
        "p:nvPr": {
          "p:extLst": {
            "p:ext": {
              "@uri": "{D42A27DB-BD31-4B8C-83A1-F6EECF244321}",
              "p14:modId": {
                "@xmlns:p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
                "@val": "1579011935"
              }
            }
          }
        }
      },
      "p:xfrm": {
        "a:off": {
          "@x": options.x || "1524000",
          "@y": options.y || "1397000"
        },
        "a:ext": {
          "@cx": options.cx || "6096000",
          "@cy": options.cy || "1483360"
        }
      },
      "a:graphic": {
        "a:graphicData": {
          "@uri": "http://schemas.openxmlformats.org/drawingml/2006/table",
          "a:tbl": {
            "a:tblPr": {
              "@firstRow": "1",
              "@bandRow": "1",
              "a:tableStyleId":options.tabstyle//"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"
            },
            "a:tblGrid": {
              "#list": colSpecs
            },

            "#list": [rowSpecs]  // replace this with  an array of table row objects
          }
        }
      }
    }
  }
}
```
- example usage
```shell
...

  // assume passed in an array of row objects
  getTable: function(rows, tblOpts) {
tblOpts = tblOpts || {};

var self = this;

return self._getBase(
  rows.map(function(row) {
    return self._getRow(
      row.map(function(cell) {
        cell = cell || {};
        if (typeof cell === 'string' || typeof cell === 'number') {
          var val = cell;
          cell = {
...
```

#### <a name="apidoc.element.officegen.genofficetable._getCell"></a>[function <span class="apidocSignatureSpan">officegen.genofficetable.</span>_getCell (val, options, idx, row_idx)](#apidoc.element.officegen.genofficetable._getCell)
- description and source-code
```javascript
_getCell = function (val, options, idx, row_idx) {
  var font_size = options.font_size || 14;
  var font_face = options.font_face || "Times New Roman";
  return {
    "a:tc": {
      "a:txBody": {
        "a:bodyPr": {},
        "a:lstStyle": {},

        "a:p": {"a:pPr":{"@algn":options.align?(options.align[idx]?options.align[idx]:'ctr'):'ctr'},
          "a:r": {
            "a:rPr": {
              "@lang": "en-US",
              "@sz": ""+font_size*100,
              "@dirty": "0",
              "@smtClean": "0",
				"@b":options.bold?(options.bold[row_idx]?(options.bold[row_idx][idx]?options.bold[row_idx][idx]:"0"):"0"):"0",
              "@i":options.italics?(options.italics[row_idx]?(options.italics[row_idx][idx]?options.italics[row_idx][idx]:"0"):"
0"):"0",
              "a:latin": {
                "@typeface": font_face
              },
              "a:cs": {
                "@typeface": font_face
              }
            },
            "a:t": val  // this is the cell value
          },
          "a:endParaRPr": {
            "@lang": "en-US",
            "@sz": ""+font_size*100,
            "@dirty": "0",
            "a:latin": {
              "@typeface": font_face
            },
            "a:cs": {
              "@typeface": font_face
            }
          }
        }
      },
      "a:tcPr": {}
    }
  }
}
```
- example usage
```shell
...
        if (typeof cell === 'string' || typeof cell === 'number') {
          var val = cell;
          cell = {
            val: val
          };
        }

        return self._getCell(cell.val, cell.opts, tblOpts);
      }),
      tblOpts
    );
  }),
  self._getColSpecs(rows, tblOpts),
  tblOpts
);
...
```

#### <a name="apidoc.element.officegen.genofficetable._getColSpecs"></a>[function <span class="apidocSignatureSpan">officegen.genofficetable.</span>_getColSpecs (rows, options)](#apidoc.element.officegen.genofficetable._getColSpecs)
- description and source-code
```javascript
_getColSpecs = function (rows, options) {
  var self = this;
  return rows[0].map(function(val,idx) {
    return self._tblGrid(idx, options);
  })
}
```
- example usage
```shell
...
          }

          return self._getCell(cell.val, cell.opts, tblOpts);
        }),
        tblOpts
      );
    }),
    self._getColSpecs(rows, tblOpts),
    tblOpts
  );
},

_getBase: function(rowSpecs, colSpecs, opts) {
  var self = this;
...
```

#### <a name="apidoc.element.officegen.genofficetable._getRow"></a>[function <span class="apidocSignatureSpan">officegen.genofficetable.</span>_getRow (cells, options)](#apidoc.element.officegen.genofficetable._getRow)
- description and source-code
```javascript
_getRow = function (cells, options) {
  return {
    "a:tr": {
      "@h": options.rowHeight ||"0", //|| "370840",
      "#list": [cells] // populate this with an array of table cell objects
    }
  }
}
```
- example usage
```shell
...
  getTable: function(rows, tblOpts) {
tblOpts = tblOpts || {};

var self = this;

return self._getBase(
  rows.map(function(row) {
    return self._getRow(
      row.map(function(cell) {
        cell = cell || {};
        if (typeof cell === 'string' || typeof cell === 'number') {
          var val = cell;
          cell = {
            val: val
          };
...
```

#### <a name="apidoc.element.officegen.genofficetable._tblGrid"></a>[function <span class="apidocSignatureSpan">officegen.genofficetable.</span>_tblGrid (idx, options)](#apidoc.element.officegen.genofficetable._tblGrid)
- description and source-code
```javascript
_tblGrid = function (idx, options) {
  return {
    "a:gridCol": {
      "@w": (options.columnWidths ? options.columnWidths[idx] : options.columnWidth||"0" ) //|| "2048000"
    }
  };
}
```
- example usage
```shell
...
  }
  return baseTable;
},

_getColSpecs: function(cols, opts) {
  var self = this;
  return cols[0].map(function(val, idx) {
    return self._tblGrid(opts);
  });
},

// TODO
_tblGrid: function(opts) {
  return {
    "w:gridCol": {
...
```

#### <a name="apidoc.element.officegen.genofficetable.getTable"></a>[function <span class="apidocSignatureSpan">officegen.genofficetable.</span>getTable (rows, options)](#apidoc.element.officegen.genofficetable.getTable)
- description and source-code
```javascript
getTable = function (rows, options) {
  var options = options || {};
	options.tabstyle=options.tabstyle?options.tabstyle:"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"
  if (options.columnWidth === undefined) {
    options.columnWidth = 8 / (rows[0].length) * EMU
  }
  var self = this;

  return self._getBase(
      rows.map(function(row,row_idx) {
        return self._getRow(
            row.map(function(val,idx) {
              return self._getCell(val, options,idx,row_idx);
            }),
            options
        );
      }),
      self._getColSpecs(rows, options),
      options
  )
}
```
- example usage
```shell
...
			outString += '</w:p>';
		} // Endif.

		// Work on all the stored paragraphs inside this document:
		for ( var i = 0, total_size = objs_list.length; i < total_size; i++ ) {

			if (objs_list[i] && objs_list[i].type === 'table') {
				var table_obj = docxTable.getTable(objs_list[i].data, objs_list[i].options);
          		var table_xml = xmlBuilder.create(table_obj,{version: '1.0', encoding: 'UTF-8', standalone: true}).toString({ pretty
: true, indent: '  ', newline: '\n' });
	          	outString += table_xml;
	          	continue;
			}

			outString += '<w:p w:rsidR="00A77427" w:rsidRDefault="007F1D13">';
			var pPrData = '';
...
```



# <a name="apidoc.module.officegen.msexcel_builder"></a>[module officegen.msexcel_builder](#apidoc.module.officegen.msexcel_builder)

#### <a name="apidoc.element.officegen.msexcel_builder.createWorkbook"></a>[function <span class="apidocSignatureSpan">officegen.msexcel_builder.</span>createWorkbook (fpath, fname)](#apidoc.element.officegen.msexcel_builder.createWorkbook)
- description and source-code
```javascript
createWorkbook = function (fpath, fname) {
  return new Workbook(fpath, fname);
}
```
- example usage
```shell
...
			gen_private.pages[pageNumber].data[objNumber].renderType = 'column';
			gen_private.pages[pageNumber].data[objNumber].title = data['title'];
			gen_private.pages[pageNumber].data[objNumber].data = data['data'];
			gen_private.pages[pageNumber].data[objNumber].options = data;   // include generic options like x,y,cx,cy, etc.
			//data['renderType'] = renderType;

			// First, generate a temporatory excel file for storing the chart's data
			var workbook = excelbuilder.createWorkbook( genobj.options.tempDir, 'sample' + GLOBAL_CHART_COUNT + '.xlsx');

			// Create a new worksheet with 10 columns and 12 rows
			// number of columns: data['data'].length+1 -> equaly number of series
			// number of rows: data['data'][0].values.length+1
			var sheet1 = workbook.createSheet('Sheet1', data['data'].length+1, data['data'][0].values.length+1);
			var headerrow = 1;
...
```



# <a name="apidoc.module.officegen.msofficegen"></a>[module officegen.msofficegen](#apidoc.module.officegen.msofficegen)

#### <a name="apidoc.element.officegen.msofficegen.makemsdoc"></a>[function <span class="apidocSignatureSpan">officegen.msofficegen.</span>makemsdoc ( genobj, new_type, options, gen_private, type_info )](#apidoc.element.officegen.msofficegen.makemsdoc)
- description and source-code
```javascript
function makemsdoc( genobj, new_type, options, gen_private, type_info ) {
	/**
	 * Generate string of the current date and time.
	 *
	 * This method generating a string with the current date and time in Office XML format.
	 *
	 * @return String of the current date and time in Office XML format.
	 */
	function getCurDateTimeForOffice () {
		var date = new Date ();

		var year = date.getFullYear ();
		var month = date.getMonth () + 1;
		var day = date.getDate ();
		var hour = date.getHours ();
		var min = date.getMinutes ();
		var sec = date.getSeconds ();

		month = (month < 10 ? "0" : "") + month;
		day = (day < 10 ? "0" : "") + day;
		hour = (hour < 10 ? "0" : "") + hour;
		min = (min < 10 ? "0" : "") + min;
		sec = (sec < 10 ? "0" : "") + sec;

		return year + "-" + month + "-" + day + "T" + hour + ":" + min + ":" + sec + 'Z';
	}

	/**
	 * Compact the given array.
	 *
	 * This function compacting the given array.
	 *
	 * @param {object} arr The array to compact.
	 */
	function compactArray ( arr ) {
		var len = arr.length, i, outArray;

		for ( i = 0; i < len; i++ ) {
			if ( arr[i] ) {
				arr.push ( arr[i] );
			} // Endif.
		} // End of for loop.

		arr.splice ( 0 , len ); // Cut the array and leave only the non-empty values.
	}

	/**
	 * Create the main files list resource.
	 *
	 * @param {object} data Ignored by this callback function.
	 * @return Text string.
	 */
	function cbMakeMainFilesList ( data ) {
		var outString = gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml ( data );
		outString += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';

		for ( var i = 0, total_size = gen_private.type.msoffice.files_list.length; i < total_size; i++ ) {
			if ( typeof gen_private.type.msoffice.files_list[i] != 'undefined' ) {
				if ( gen_private.type.msoffice.files_list[i].ext )
				{
					
					// Fixed by @author:vtloc @date:2014Jan09
					// Reason: if we write out duplicate extension, office will raise error
					//
					if( outString.indexOf('<Default Extension="'+gen_private.type.msoffice.files_list[i].ext) == -1 )
					{ // check to make sure we don't write duplicate extension tag
						outString += '<Default Extension="' + gen_private.type.msoffice.files_list[i].ext + '" ContentType="' + gen_private.type.msoffice
.files_list[i].type + '"/>';
					}
				} else
				{
					outString += '<Override PartName="' + gen_private.type.msoffice.files_list[i].name + '" ContentType="' + gen_private.type.msoffice
.files_list[i].type + '"/>';
				} // Endif.
			} // Endif.
		} // End of for loop.

		outString += '</Types>\n';
		return outString;
	}

	/**
	 * ???.
	 *
	 * @param {object} data Ignored by this callback function.
	 * @return Text string.
	 */
	function cbMakeTheme ( data ) {
		if ( genobj.theme ) {
			// console.log ( genobj.theme );
			return genobj.theme;
		} // Endif.

		return gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml ( data ) + '<a:theme xmlns:a="http://schemas.openxmlformats.org/
drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="
000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr
 val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3
><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5
><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></
a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface
=""/><a:font script="Jpan" typeface="MS P????"/><a:font script="Hang" typeface="?? ??"/><a:font script="Hans" typeface="??"/><a:
font script="Hant" typeface="????"/><a:font script="Arab" typeface="Times New Roman"/><a:font script ...
```
- example usage
```shell
n/a
```



# <a name="apidoc.module.officegen.officechart"></a>[module officegen.officechart](#apidoc.module.officegen.officechart)

#### <a name="apidoc.element.officegen.officechart.officechart"></a>[function <span class="apidocSignatureSpan">officegen.</span>officechart (chartInfo)](#apidoc.element.officegen.officechart.officechart)
- description and source-code
```javascript
function OfficeChart(chartInfo) {

  if (chartInfo instanceof OfficeChart) {
    return chartInfo;
  }

  return {

    chartSpec: null, // Javascript object that represents the XML tree for the PowerPoint chart

    toXML: function () {
      return xmlBuilder.create(this.chartSpec, {version: '1.0', encoding: 'UTF-8', standalone: true}).end({ pretty: true, indent
: '  ', newline: '\n' });
    },

    toJSON: function () {
      return this.chartSpec;
    },

    getClass: function() { return "OfficeChart"; },

    ///
    /// @brief Create XML representation of chart object
    /// @param[chartInfo] object
    ///  {
    ///			title: 'eSurvey chart',
    ///			data:  [ // array of series
    ///				{
    ///					name: 'Income',
    ///					labels: ['2005', '2006', '2007', '2008', '2009'],
    ///					values: [23.5, 26.2, 30.1, 29.5, 24.6],
    ///         color: 'ff0000'
    ///				}
    ///			],
    ///     overlap:  "0",
    ///     gapWidth: "150"
    ///
    /// 	}

    initialize: function (chartInfo) {

      if (chartInfo.getClass &&  chartInfo.getClass()== 'OfficeChart') {
        return chartInfo;
      }

      // overlap ["50"] is handled as an option within the chartbase
      // gapWidth ["150"] is handled as an option within the chartbase
      // valAxisCrossAtMaxCategory [true|false] is handled as an option within the chart base
      // catAxisReverseOrder [true|false] is handled as an option within the chart base

      this.chartSpec = OfficeChart.getChartBase(chartInfo);  // get foundation XML for the chart type

      // Below are methods for handling options with more complex XML to mix in
      this.setData(chartInfo['data']);
      this.setTitle(chartInfo.title || chartInfo.name);
      this.setValAxisTitle(chartInfo.valAxisTitle);
      this.setCatAxisTitle(chartInfo.catAxisTitle);
      this.setValAxisNumFmt(chartInfo.valAxisNumFmt);
      this.setValAxisScale(chartInfo.valAxisMinValue, chartInfo.valAxisMaxValue);
      this.setTextSize(chartInfo.fontSize);
      this.mergeChartXml(chartInfo.xml);
      this.setValAxisMajorGridlines(chartInfo.valAxisMajorGridlines);
      this.setValAxisMinorGridlines(chartInfo.valAxisMinorGridlines);

      return this;
    },

    setTextSize: function(textSize) {
      if (textSize !== undefined) {
         var textRef = this._text(textSize);
        _.merge(this.chartSpec['c:chartSpace'], textRef);
      }
    },

    setTitle: function (title) {
      if (title !== undefined) {
        var titleRef = this._title(chartInfo.title || chartInfo.name);
        _.merge(this.chartSpec['c:chartSpace']['c:chart'], titleRef)
      }
    },

    setValAxisTitle: function (title) {
      if (title) {
        var titleRef = this._title(title);
        _.merge(this.chartSpec['c:chartSpace']['c:chart']['c:plotArea']['c:valAx'], titleRef);
      }
    },

    setCatAxisTitle: function (title) {
      if (title) {
        var titleRef = this._title(title);
        _.merge(this.chartSpec['c:chartSpace']['c:chart']['c:plotArea']['c:catAx'], titleRef);
      }
    },

    setValAxisNumFmt: function (format) {
      if (format !== undefined) {
        var numFmtRef = this._numFmt(format);
        _.merge(this.chartSpec['c:chartSpace']['c:chart']['c:plotArea']['c:valAx'], numFmtRef)
      }
    },

    setValAxisScale: function (min, max) {
      if (min!==undefined || max!== undefined) {
        var scalingRef = this._scaling(min, max);
        _.merge(this.chartSpec['c:chartSpace']['c:chart']['c:plotArea']['c:valAx'], scalingRef);
      }
    },

    mergeChartXml: function (xml) {
      if (xml !== undefined) {
        _.merge(this.chartSpec['c:chartSpace'], xml);
      }
    },

    setValAxisMajorGridlines: function(boolean) {
      if (boolean) {
        this.chartSpec['c:chartSpace']['c:chart']['c:plotArea']['c:valAx']['c:majorGridlines'] = {};
      }
    },
    setValAxisMinorGridlines: function(boolean) {
      if (boolean) {
        this.chartSpec[ ...
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.officechart.getChartBase"></a>[function <span class="apidocSignatureSpan">officegen.officechart.</span>getChartBase (chartInfo)](#apidoc.element.officegen.officechart.getChartBase)
- description and source-code
```javascript
getChartBase = function (chartInfo) {

  var chartBase;

  if (typeof chartInfo == 'string') {
    chartBase = _chartSpecs[chartInfo]();
  }
  else if (typeof chartInfo.renderType == 'string') {
    chartBase = _chartSpecs[chartInfo.renderType](chartInfo);
  }
  else if (chartInfo.xml) {
    chartBase = chartInfo.xml
  }
  else {
    throw new Error("invalid chart type")
  }
  // return deep copy so can create multiple charts from same base within one PowerPoint document
  return JSON.parse(JSON.stringify(chartBase));
}
```
- example usage
```shell
...
}

// overlap ["50"] is handled as an option within the chartbase
// gapWidth ["150"] is handled as an option within the chartbase
// valAxisCrossAtMaxCategory [true|false] is handled as an option within the chart base
// catAxisReverseOrder [true|false] is handled as an option within the chart base

this.chartSpec = OfficeChart.getChartBase(chartInfo);  // get foundation XML for the chart type

// Below are methods for handling options with more complex XML to mix in
this.setData(chartInfo['data']);
this.setTitle(chartInfo.title || chartInfo.name);
this.setValAxisTitle(chartInfo.valAxisTitle);
this.setCatAxisTitle(chartInfo.catAxisTitle);
this.setValAxisNumFmt(chartInfo.valAxisNumFmt);
...
```



# <a name="apidoc.module.officegen.openofficegen"></a>[module officegen.openofficegen](#apidoc.module.officegen.openofficegen)

#### <a name="apidoc.element.officegen.openofficegen.makeoodoc"></a>[function <span class="apidocSignatureSpan">officegen.openofficegen.</span>makeoodoc ( genobj, new_type, options, gen_private, type_info )](#apidoc.element.officegen.openofficegen.makeoodoc)
- description and source-code
```javascript
function makeoodoc( genobj, new_type, options, gen_private, type_info ) {
	<span class="apidocCodeCommentSpan">/**
	 * Get the string that opening every Office XML type.
	 * <br /><br />
	 *
	 * Every OpenOffice XML resource will have this header at the begining of the file.
	 *
	 * @param {object} data Ignored by this callback function.
	 * @return Text string.
	 */
</span>	gen_private.plugs.type.openoffice.makeOpenOfficeBasicXml = function ( data ) {
		return '<?xml version="1.0" encoding="UTF-8"?>\n';
	};

	// Basic API for plugins:

	gen_private.plugs.type.openoffice = {};

	/**
	 * Create the mimetpe resource.
	 * <br /><br />
	 *
	 * Every OpenOffice based document must have this resource.
	 *
	 * @param {object} data Ignored by this callback function.
	 * @return Text string.
	 */
	function makeOpenOfficeMimeType ( data ) {
		return 'application/vnd.oasis.opendocument.' + gen_private.mixed.res_data.mimeType;
	}

	/**
	 * Generate the manifest XML resource.
	 * <br /><br />
	 *
	 * Tbis function creating the manifest resource.
	 *
	 * @param {object} data Array filled with all the resources information.
	 * @return Text string.
	 */
	gen_private.plugs.type.openoffice.makeManifest = function ( data ) {
		var outString = gen_private.plugs.type.openoffice.makeOpenOfficeBasicXml ( data );
		outString += '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">\n
';

		// Add all the rels records inside the data array:
		for ( var i = 0, total_size = data.length; i < total_size; i++ ) {
			if ( typeof data[i] != 'undefined' ) {
				outString += ' <manifest:file-entry manifest:media-type="' + data[i].type + '" manifest:full-path="' + data[i].target + '"/>\n';
			} // Endif.
		} // End of for loop.

		outString += '</manifest:manifest>\n';
		return outString;
	};

	/**
	 * Prepare the officegen object to OpenOffice documents.
	 * <br /><br />
	 *
	 * Every plugin that implementing gemenrating OpenOffice document must call this method to initialize
	 * the common stuff.
	 *
	 * @param {object} mimeType The mime type of this document.
	 * @param {object} ext_opt Optional settings (unused right now).
	 */
	gen_private.plugs.type.openoffice.makeOfficeGenerator = function ( mimeType, ext_opt ) {
		gen_private.mixed.res_data.mimeType = mimeType;
		gen_private.mixed.files_list = [];

		gen_private.mixed.files_list.push (
			{
				name: 'content.xml',
				type: 'text/xml'
			},
			{
				name: 'settings.xml',
				type: 'text/xml'
			},
			{
				name: 'styles.xml',
				type: 'text/xml'
			},
			{
				name: 'manifest.rdf',
				type: 'application/rdf+xml'
			},
			{
				name: 'meta.xml',
				type: 'text/xml'
			}
		);

		gen_private.plugs.intAddAnyResourceToParse ( 'mimetype', 'buffer', null, makeOpenOfficeMimeType, true );
		gen_private.plugs.intAddAnyResourceToParse ( 'META-INF\\manifest.xml', 'buffer', gen_private.mixed.rels_main, gen_private.plugs
.type.openoffice.cbMakeRels, true );
		// gen_private.plugs.intAddAnyResourceToParse ( 'settings.xml', 'buffer', null, make???, true );
		// gen_private.plugs.intAddAnyResourceToParse ( 'styles.xml', 'buffer', null, make???, true );
		// gen_private.plugs.intAddAnyResourceToParse ( 'manifest.rdf', 'buffer', null, make???, true );
		// gen_private.plugs.intAddAnyResourceToParse ( 'meta.xml', 'buffer', null, make???, true );
	};

	return this;
}
```
- example usage
```shell
n/a
```



# <a name="apidoc.module.officegen.plugins"></a>[module officegen.plugins](#apidoc.module.officegen.plugins)

#### <a name="apidoc.element.officegen.plugins.getDocTypeByName"></a>[function <span class="apidocSignatureSpan">officegen.plugins.</span>getDocTypeByName ( typeName )](#apidoc.element.officegen.plugins.getDocTypeByName)
- description and source-code
```javascript
getDocTypeByName = function ( typeName ) {
		return officegenGlobals.types[typeName];
	}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.plugins.getPrototypeByName"></a>[function <span class="apidocSignatureSpan">officegen.plugins.</span>getPrototypeByName ( typeName )](#apidoc.element.officegen.plugins.getPrototypeByName)
- description and source-code
```javascript
function getPrototypeByName( typeName ) {
		return officegenGlobals.docPrototypes[typeName];
	}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.plugins.getVerboseMode"></a>[function <span class="apidocSignatureSpan">officegen.plugins.</span>getVerboseMode ( docType, moduleName )](#apidoc.element.officegen.plugins.getVerboseMode)
- description and source-code
```javascript
getVerboseMode = function ( docType, moduleName ) {
		if ( !docType && !moduleName ) {
			return officegenGlobals.settings.verbose ? true : false;
		} // Endif.

		var verboseFlag;

		if ( officegenGlobals.settings.verbose && typeof officegenGlobals.settings.verbose === 'object' ) {
			if ( docType && officegenGlobals.settings.verbose.docType && typeof officegenGlobals.settings.verbose.docType === 'object' &&
officegenGlobals.settings.verbose.docType.indexOf ) {
				verboseFlag = (officegenGlobals.settings.verbose.docType.indexOf ( docType ) >= 0) ? true : false;
			} // Endif.

			if ( verboseFlag !== false && moduleName && officegenGlobals.settings.verbose.moduleName && typeof officegenGlobals.settings.
verbose.moduleName === 'object' && officegenGlobals.settings.verbose.moduleName.indexOf ) {
				verboseFlag = (officegenGlobals.settings.verbose.moduleName.indexOf ( moduleName ) >= 0) ? true : false;
			} // Endif.
		} // Endif.

		return verboseFlag ? true : false;
	}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.plugins.registerDocType"></a>[function <span class="apidocSignatureSpan">officegen.plugins.</span>registerDocType ( typeName, createFunc, schema_data, docType, displayName )](#apidoc.element.officegen.plugins.registerDocType)
- description and source-code
```javascript
registerDocType = function ( typeName, createFunc, schema_data, docType, displayName ) {
		officegenGlobals.types[typeName] = {};
		officegenGlobals.types[typeName].createFunc = createFunc;
		officegenGlobals.types[typeName].schema_data = schema_data;
		officegenGlobals.types[typeName].type = docType;
		officegenGlobals.types[typeName].display = displayName;
	}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.plugins.registerParserType"></a>[function <span class="apidocSignatureSpan">officegen.plugins.</span>registerParserType ( typeName, parserFunc, extra_data, displayName )](#apidoc.element.officegen.plugins.registerParserType)
- description and source-code
```javascript
registerParserType = function ( typeName, parserFunc, extra_data, displayName ) {
		officegenGlobals.resParserTypes[typeName] = {};
		officegenGlobals.resParserTypes[typeName].parserFunc = parserFunc;
		officegenGlobals.resParserTypes[typeName].extra_data = extra_data;
		officegenGlobals.resParserTypes[typeName].display = displayName;
	}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.plugins.registerPrototype"></a>[function <span class="apidocSignatureSpan">officegen.plugins.</span>registerPrototype ( typeName, baseObj, displayName )](#apidoc.element.officegen.plugins.registerPrototype)
- description and source-code
```javascript
registerPrototype = function ( typeName, baseObj, displayName ) {
		officegenGlobals.docPrototypes[typeName] = {};
		officegenGlobals.docPrototypes[typeName].baseObj = baseObj;
		officegenGlobals.docPrototypes[typeName].display = displayName;
	}
```
- example usage
```shell
n/a
```



# <a name="apidoc.module.officegen.pptxplg_layouts"></a>[module officegen.pptxplg_layouts](#apidoc.module.officegen.pptxplg_layouts)

#### <a name="apidoc.element.officegen.pptxplg_layouts.pptxplg_layouts"></a>[function <span class="apidocSignatureSpan">officegen.</span>pptxplg_layouts ( pluginsman )](#apidoc.element.officegen.pptxplg_layouts.pptxplg_layouts)
- description and source-code
```javascript
function MakeLayoutPlugin( pluginsman ) {
	var funcThis = this;

	this.ogPluginsApi = pluginsman.ogPluginsApi; // Generic officegen API for plugins.
	this.msPluginsApi = pluginsman.genPrivate.plugs.type.msoffice; // msoffice plugins API.
	this.pluginsman = pluginsman; // Document type specific plugins API.

	this.pptxData = pluginsman.getDataStorage (); // Here you can store any temporary data needed for generating the document and depending
 on the data filled by the user.

	this.mainPath = pluginsman.genPrivate.features.type.msoffice.main_path; // The "folder" name inside the document zip that all the
 specific resources of this document type are stored.
	this.mainPathFile = pluginsman.genPrivate.features.type.msoffice.main_path_file; // The name of the main real xml resource of this
 document.
	this.relsMain = pluginsman.genPrivate.type.msoffice.rels_main; // Main rels file.
	this.relsApp = pluginsman.genPrivate.type.msoffice.rels_app; // Main rels file inside the specific document type "folder".
	this.filesList = pluginsman.genPrivate.type.msoffice.files_list; // Resources list xml.
	this.srcFilesList = pluginsman.genPrivate.type.msoffice.src_files_list; // For storing extra files inside the document zip.

	//
	// Catch events inside the document so we can extent it:
	//

	// We want to extend the main API of the pptx document object:
	pluginsman.registerCallback ( 'makeDocApi', function ( docObj ) { funcThis.extendPptxApi ( docObj ); } );

	// We want to extend the slide object API:
	pluginsman.registerCallback ( 'newPage', function ( docData ) { funcThis.extendPptxSlideApi ( docData ); } );

	// This event tell us that the generator is about to start working:
	pluginsman.registerCallback ( 'beforeGen', function ( docObj ) { funcThis.beforeGen ( docObj ); } );

	return this;
}
```
- example usage
```shell
n/a
```



# <a name="apidoc.module.officegen.pptxplg_layouts.prototype"></a>[module officegen.pptxplg_layouts.prototype](#apidoc.module.officegen.pptxplg_layouts.prototype)

#### <a name="apidoc.element.officegen.pptxplg_layouts.prototype.beforeGen"></a>[function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>beforeGen ( docObj )](#apidoc.element.officegen.pptxplg_layouts.prototype.beforeGen)
- description and source-code
```javascript
beforeGen = function ( docObj ) {
	var funcThis = this;

	if ( !this.pptxData.slideLayouts ) {
		this.pptxData.slideLayouts = [];
	} // Endif.

	var slideLayouts = this.pptxData.slideLayouts;

	function addLayout ( curId ) {
		slideLayouts.push (
			{
				execCb: function ( data ) { return funcThis['cbMakePptxLayout' + curId] ( data ); },
				cbData: {}
			}
		);
	}

	// Add all the office additional standard layouts:
	for ( var i = 0; i < 10; i++ ) {
		addLayout ( i + 2 );
	} // End of for loop.

	// Work on all the registered layout so we'll add them into the output document:
	var curId = 0;
	for ( var item in slideLayouts ) {
		// Make sure that we'll know the right rel ID of this new slide resource:
		slideLayouts[item].relIdMaster = funcThis.pluginsman.genobj.slideMasterRels.length + 1;

		// Add it to the rels of the slides master:
		funcThis.pluginsman.genobj.slideMasterRels.push ({
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
			target: '../slideLayouts/slideLayout' + (curId + 2) + '.xml',
			clear: 'generate'
		});

		// Add this slide layout resource + the rels resource of it:
		funcThis.ogPluginsApi.intAddAnyResourceToParse ( funcThis.mainPath + '\\slideLayouts\\slideLayout' + (curId + 2) + '.xml', 'buffer
', slideLayouts[item].cbData, slideLayouts[item].execCb, false );
		funcThis.ogPluginsApi.intAddAnyResourceToParse ( funcThis.mainPath + '\\slideLayouts\\_rels\\slideLayout' + (curId + 2) + '.xml
.rels', 'buffer', [
			{
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
				target: '../slideMasters/slideMaster1.xml'
			}
		], funcThis.pluginsman.genPrivate.plugs.type.msoffice.cbMakeRels, false );

		// Add it also to the main list of resources in the document:
		funcThis.filesList.push ({
			name: '/ppt/slideLayouts/slideLayout' + (curId + 2) + '.xml',
			type: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
			clear: 'generate'
		});

		curId++;
	} // End of for loop.
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout10"></a>[function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>cbMakePptxLayout10 ( data )](#apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout10)
- description and source-code
```javascript
cbMakePptxLayout10 = function ( data ) {
	if ( !data || typeof data !== 'object' ) {
		data = {};
	} // Endif.

	return this.msPluginsApi.cbMakeMsOfficeBasicXml ( data ) + '<p:sld' + (data.isRealSlide ? '' : 'Layout') + ' xmlns:a="http://schemas
.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="
http://schemas.openxmlformats.org/presentationml/2006/main"' + (data.isRealSlide ? '' : ' type="vertTx" preserve="1"') + '><p:cSld
 name="Title and Vertical Text"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr
><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr
><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p
:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master title style</a:t></
a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Vertical Text Placeholder 2"/><p:cNvSpPr
><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" orient="vert" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr
 vert="eaVert"/><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master text styles</
a:t></a:r></a:p><a:p><a:pPr lvl="1"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Second level</a:t></a:r></a:p><a:p><a:pPr lvl="
2"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Third level</a:t></a:r></a:p><a:p><a:pPr lvl="3"/><a:r><a:rPr lang="en-US" smtClean
="0"/><a:t>Fourth level</a:t></a:r></a:p><a:p><a:pPr lvl="4"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Fifth level</a:t></a:r
><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Date Placeholder 3"/><p:cNvSpPr><a:spLocks
 noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt" sz="half" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle
/><a:p><a:fld id="{A76116CE-C4A3-4A05-B2D7-7C2E9A889C0F}" type="datetimeFigureOut"><a:rPr lang="en-US" smtClean="0"/><a:t>1/30/2017
</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="5" name="Footer Placeholder 4"/><
p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody
><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="6" name="Slide Number
 Placeholder 5"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr
><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{B1393E5F-521B-4CAD-9D3A-AE923D912DCE}" type="slidenum"><a:rPr lang="
en-US" smtClean="0"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><
a:masterClrMapping/></p:clrMapOvr></p:sld' + (data.isRealSlide ? '' : 'Layout') + '>';
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout11"></a>[function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>cbMakePptxLayout11 ( data )](#apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout11)
- description and source-code
```javascript
cbMakePptxLayout11 = function ( data ) {
	if ( !data || typeof data !== 'object' ) {
		data = {};
	} // Endif.

	return this.msPluginsApi.cbMakeMsOfficeBasicXml ( data ) + '<p:sld' + (data.isRealSlide ? '' : 'Layout') + ' xmlns:a="http://schemas
.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="
http://schemas.openxmlformats.org/presentationml/2006/main"' + (data.isRealSlide ? '' : ' type="vertTitleAndTx" preserve="1"') + '><
p:cSld name="Vertical Title and Text"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr
><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr
><p:cNvPr id="2" name="Vertical Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title" orient="vert"/></
p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="6629400" y="274638"/><a:ext cx="2057400" cy="5851525"/></a:xfrm></p:spPr><p:txBody><
a:bodyPr vert="eaVert"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master title style</a:t></a:
r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Vertical Text Placeholder 2"/><p:cNvSpPr
><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" orient="vert" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x
="457200" y="274638"/><a:ext cx="6019800" cy="5851525"/></a:xfrm></p:spPr><p:txBody><a:bodyPr vert="eaVert"/><a:lstStyle/><a:p><
a:pPr lvl="0"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master text styles</a:t></a:r></a:p><a:p><a:pPr lvl="1"/><
a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Second level</a:t></a:r></a:p><a:p><a:pPr lvl="2"/><a:r><a:rPr lang="en-US" smtClean="
0"/><a:t>Third level</a:t></a:r></a:p><a:p><a:pPr lvl="3"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Fourth level</a:t></a:r></
a:p><a:p><a:pPr lvl="4"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Fifth level</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></
p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Date Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><
p:ph type="dt" sz="half" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{A76116CE-C4A3
-4A05-B2D7-7C2E9A889C0F}" type="datetimeFigureOut"><a:rPr lang="en-US" smtClean="0"/><a:t>1/30/2017</a:t></a:fld><a:endParaRPr lang
="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="5" name="Footer Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></
p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a
:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="6" name="Slide Number Placeholder 5"/><p:cNvSpPr><
a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr
/><a:lstStyle/><a:p><a:fld id="{B1393E5F-521B-4CAD-9D3A-AE923D912DCE}" type="slidenum"><a:rPr lang="en-US" smtClean="0"/><a:t>‹#›</
a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr
></p:sld' + (data.isRealSlide ? '' : 'Layout') + '>';
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout2"></a>[function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>cbMakePptxLayout2 ( data )](#apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout2)
- description and source-code
```javascript
cbMakePptxLayout2 = function ( data ) {
	if ( !data || typeof data !== 'object' ) {
		data = {};
	} // Endif.

	// You can place here the title:
	var ph1 = '<a:p><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master title style</a:t></a:r><a:endParaRPr lang="en-
US"/></a:p>';
	if ( data.ph1 && data.slide && data.ph1.length ) {
		ph1 = this.pluginsman.genobj.cMakePptxOutTextP ( '', data.ph1, {}, data.slide );
	} // Endif.

	// More data:
	var ph2 = '<a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master text styles</a:t></a:r></a:p><
a:p><a:pPr lvl="1"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Second level</a:t></a:r></a:p><a:p><a:pPr lvl="2"/><a:r><a:rPr lang
="en-US" smtClean="0"/><a:t>Third level</a:t></a:r></a:p><a:p><a:pPr lvl="3"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Fourth
 level</a:t></a:r></a:p><a:p><a:pPr lvl="4"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Fifth level</a:t></a:r><a:endParaRPr lang
="en-US"/></a:p>';
	if ( data.ph2 && data.slide && data.ph2.length ) {
		ph2 = this.pluginsman.genobj.cMakePptxOutTextP ( '', data.ph2, {}, data.slide );
	} // Endif.

	var ph3 = this.pluginsman.genobj.createFieldText ( 'DATE_TIME', 1, data.useDate );

	var footFull = '';
	var ft = '';
	var curElNum = 4;

	if ( data.isDate || !data.isRealSlide ) {
		footFull += '<p:sp><p:nvSpPr><p:cNvPr id="' + curElNum + '" name="Date Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr
><p:nvPr><p:ph type="dt" sz="half" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{A76116CE
-C4A3-4A05-B2D7-7C2E9A889C0F}" type="datetimeFigureOut"><a:rPr lang="en-US" smtClean="0"/><a:t>' + ph3 + '</a:t></a:fld><a:endParaRPr
 lang="en-US"/></a:p></p:txBody></p:sp>';
		curElNum++;
	} // Endif.

	if ( data.isFooter && data.ft && data.slide && data.ft.length ) {
		ft = this.pluginsman.genobj.cMakePptxOutTextP ( '', data.ft, {}, data.slide );
		footFull += '<p:sp><p:nvSpPr><p:cNvPr id="' + curElNum + '" name="Footer Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:
cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>' + ft + '</
p:txBody></p:sp>';
		curElNum++;

	} else if ( !data.isRealSlide ) {
		footFull += '<p:sp><p:nvSpPr><p:cNvPr id="' + curElNum + '" name="Footer Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:
cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr
 lang="en-US"/></a:p></p:txBody></p:sp>';
		curElNum++;
	} // Endif.

	if ( data.isSlideNum || !data.isRealSlide ) {
		footFull += '<p:sp><p:nvSpPr><p:cNvPr id="' + curElNum + '" name="Slide Number Placeholder 5"/><p:cNvSpPr><a:spLocks noGrp="1"/></
p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p
><a:fld id="{B1393E5F-521B-4CAD-9D3A-AE923D912DCE}" type="slidenum"><a:rPr lang="en-US" smtClean="0"/><a:t>‹#›</a:t></a:fld><a:endParaRPr
 lang="en-US"/></a:p></p:txBody></p:sp>';
		curElNum++;
	} // Endif.

	return this.msPluginsApi.cbMakeMsOfficeBasicXml ( data ) + '<p:sld' + (data.isRealSlide ? '' : 'Layout') + ' xmlns:a="http://schemas
.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="
http://schemas.openxmlformats.org/presentationml/2006/main"' + (data.isRealSlide ? '' : ' type="secHead" preserve="1"') + '><p:cSld
 name="Section Header"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><
a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr
 id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:
xfrm><a:off x="722313" y="4406900"/><a:ext cx="7772400" cy="1362075"/></a:xfrm></p:spPr><p:txBody><a:bodyPr anchor="t"/><a:lstStyle
><a:lvl1pPr algn="l"><a:defRPr sz ...
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout3"></a>[function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>cbMakePptxLayout3 ( data )](#apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout3)
- description and source-code
```javascript
cbMakePptxLayout3 = function ( data ) {
	if ( !data || typeof data !== 'object' ) {
		data = {};
	} // Endif.

	// You can place here the title:
	var ph1 = '<a:p><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master title style</a:t></a:r><a:endParaRPr lang="en-
US"/></a:p>';
	if ( data.ph1 && data.slide && data.ph1.length ) {
		ph1 = this.pluginsman.genobj.cMakePptxOutTextP ( '', data.ph1, {}, data.slide );
	} // Endif.

	// More data:
	var ph2 = '<a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master text styles</a:t></a:r></a:p>';
	if ( data.ph2 && data.slide && data.ph2.length ) {
		ph2 = this.pluginsman.genobj.cMakePptxOutTextP ( '', data.ph2, {}, data.slide );
	} // Endif.

	var ph3 = this.pluginsman.genobj.createFieldText ( 'DATE_TIME', 1, data.useDate );

	var footFull = '';
	var ft = '';
	var curElNum = 4;

	if ( data.isDate || !data.isRealSlide ) {
		footFull += '<p:sp><p:nvSpPr><p:cNvPr id="' + curElNum + '" name="Date Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr
><p:nvPr><p:ph type="dt" sz="half" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{A76116CE
-C4A3-4A05-B2D7-7C2E9A889C0F}" type="datetimeFigureOut"><a:rPr lang="en-US" smtClean="0"/><a:t>' + ph3 + '</a:t></a:fld><a:endParaRPr
 lang="en-US"/></a:p></p:txBody></p:sp>';
		curElNum++;
	} // Endif.

	if ( data.isFooter && data.ft && data.slide && data.ft.length ) {
		ft = this.pluginsman.genobj.cMakePptxOutTextP ( '', data.ft, {}, data.slide );
		footFull += '<p:sp><p:nvSpPr><p:cNvPr id="' + curElNum + '" name="Footer Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:
cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>' + ft + '</
p:txBody></p:sp>';
		curElNum++;

	} else if ( !data.isRealSlide ) {
		footFull += '<p:sp><p:nvSpPr><p:cNvPr id="' + curElNum + '" name="Footer Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:
cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr
 lang="en-US"/></a:p></p:txBody></p:sp>';
		curElNum++;
	} // Endif.

	if ( data.isSlideNum || !data.isRealSlide ) {
		footFull += '<p:sp><p:nvSpPr><p:cNvPr id="' + curElNum + '" name="Slide Number Placeholder 5"/><p:cNvSpPr><a:spLocks noGrp="1"/></
p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p
><a:fld id="{B1393E5F-521B-4CAD-9D3A-AE923D912DCE}" type="slidenum"><a:rPr lang="en-US" smtClean="0"/><a:t>‹#›</a:t></a:fld><a:endParaRPr
 lang="en-US"/></a:p></p:txBody></p:sp>';
		curElNum++;
	} // Endif.

	return this.msPluginsApi.cbMakeMsOfficeBasicXml ( data ) + '<p:sld' + (data.isRealSlide ? '' : 'Layout') + ' xmlns:a="http://schemas
.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="
http://schemas.openxmlformats.org/presentationml/2006/main"' + (data.isRealSlide ? '' : ' type="secHead" preserve="1"') + '><p:cSld
 name="Section Header"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><
a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr
 id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:
xfrm><a:off x="722313" y="4406900"/><a:ext cx="7772400" cy="1362075"/></a:xfrm></p:spPr><p:txBody><a:bodyPr anchor="t"/><a:lstStyle
><a:lvl1pPr algn="l"><a:defRPr sz="4000" b="1" cap="all"/></a:lvl1pPr></a:lstStyle>' + ph1 + '</p:txBody></p:sp><p:sp><p:nvSpPr><
p:cNvPr id="3" name="Text Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr
></p:nvSpPr><p:spPr><a:xfrm><a:off x="722313" y="2906713"/><a:ext cx="7772400" cy="1500187"/></a:xfrm></p:spPr><p:txBody><a:bodyPr
 anchor="b"/><a:lstStyle><a:lvl1pPr marL="0" indent ...
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout4"></a>[function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>cbMakePptxLayout4 ( data )](#apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout4)
- description and source-code
```javascript
cbMakePptxLayout4 = function ( data ) {
	if ( !data || typeof data !== 'object' ) {
		data = {};
	} // Endif.

	return this.msPluginsApi.cbMakeMsOfficeBasicXml ( data ) + '<p:sld' + (data.isRealSlide ? '' : 'Layout') + ' xmlns:a="http://schemas
.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="
http://schemas.openxmlformats.org/presentationml/2006/main"' + (data.isRealSlide ? '' : ' type="twoObj" preserve="1"') + '><p:cSld
 name="Two Content"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:
off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr
 id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p
:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master title style</a:t></a:r><a:
endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Content Placeholder 2"/><p:cNvSpPr><a:spLocks
 noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph sz="half" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="1600200"/><a:ext
 cx="4038600" cy="4525963"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="2800"/></a:lvl1pPr><a:lvl2pPr
><a:defRPr sz="2400"/></a:lvl2pPr><a:lvl3pPr><a:defRPr sz="2000"/></a:lvl3pPr><a:lvl4pPr><a:defRPr sz="1800"/></a:lvl4pPr><a:lvl5pPr
><a:defRPr sz="1800"/></a:lvl5pPr><a:lvl6pPr><a:defRPr sz="1800"/></a:lvl6pPr><a:lvl7pPr><a:defRPr sz="1800"/></a:lvl7pPr><a:lvl8pPr
><a:defRPr sz="1800"/></a:lvl8pPr><a:lvl9pPr><a:defRPr sz="1800"/></a:lvl9pPr></a:lstStyle><a:p><a:pPr lvl="0"/><a:r><a:rPr lang
="en-US" smtClean="0"/><a:t>Click to edit Master text styles</a:t></a:r></a:p><a:p><a:pPr lvl="1"/><a:r><a:rPr lang="en-US" smtClean
="0"/><a:t>Second level</a:t></a:r></a:p><a:p><a:pPr lvl="2"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Third level</a:t></a:r
></a:p><a:p><a:pPr lvl="3"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Fourth level</a:t></a:r></a:p><a:p><a:pPr lvl="4"/><a:r><
a:rPr lang="en-US" smtClean="0"/><a:t>Fifth level</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr
><p:cNvPr id="4" name="Content Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph sz="half" idx="2"/></p
:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="4648200" y="1600200"/><a:ext cx="4038600" cy="4525963"/></a:xfrm></p:spPr><p:txBody><
a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="2800"/></a:lvl1pPr><a:lvl2pPr><a:defRPr sz="2400"/></a:lvl2pPr><a:lvl3pPr><a:defRPr
 sz="2000"/></a:lvl3pPr><a:lvl4pPr><a:defRPr sz="1800"/></a:lvl4pPr><a:lvl5pPr><a:defRPr sz="1800"/></a:lvl5pPr><a:lvl6pPr><a:defRPr
 sz="1800"/></a:lvl6pPr><a:lvl7pPr><a:defRPr sz="1800"/></a:lvl7pPr><a:lvl8pPr><a:defRPr sz="1800"/></a:lvl8pPr><a:lvl9pPr><a:defRPr
 sz="1800"/></a:lvl9pPr></a:lstStyle><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master text
styles</a:t></a:r></a:p><a:p><a:pPr lvl="1"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Second level</a:t></a:r></a:p><a:p><a:pPr
 lvl="2"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Third level</a:t></a:r></a:p><a:p><a:pPr lvl="3"/><a:r><a:rPr lang="en-US"
smtClean="0"/><a:t>Fourth level</a:t></a:r></a:p><a:p><a:pPr lvl="4"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Fifth level</a
:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="5" name="Date Placeholder 4"/><p:cNvSpPr
><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt" sz="half" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><
a:lstStyle/><a:p><a:fld id="{A76116CE-C4A3-4A05-B2D7-7C2E9A889C0F}" type="datetimeFigureOut"><a:rPr lang="en-US" smtClean="0"/><
a:t>1/30/2017</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="6" name="Footer Placeholder
 5"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr ...
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout5"></a>[function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>cbMakePptxLayout5 ( data )](#apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout5)
- description and source-code
```javascript
cbMakePptxLayout5 = function ( data ) {
	if ( !data || typeof data !== 'object' ) {
		data = {};
	} // Endif.

	return this.msPluginsApi.cbMakeMsOfficeBasicXml ( data ) + '<p:sld' + (data.isRealSlide ? '' : 'Layout') + ' xmlns:a="http://schemas
.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="
http://schemas.openxmlformats.org/presentationml/2006/main"' + (data.isRealSlide ? '' : ' type="twoTxTwoObj" preserve="1"') + '><
p:cSld name="Comparison"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm
><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p
:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr
/><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr/></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr lang="en-US" smtClean="0"/><a
:t>Click to edit Master title style</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="
3" name="Text Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr
><p:spPr><a:xfrm><a:off x="457200" y="1535113"/><a:ext cx="4040188" cy="639762"/></a:xfrm></p:spPr><p:txBody><a:bodyPr anchor="b
"/><a:lstStyle><a:lvl1pPr marL="0" indent="0"><a:buNone/><a:defRPr sz="2400" b="1"/></a:lvl1pPr><a:lvl2pPr marL="457200" indent="
0"><a:buNone/><a:defRPr sz="2000" b="1"/></a:lvl2pPr><a:lvl3pPr marL="914400" indent="0"><a:buNone/><a:defRPr sz="1800" b="1"/></
a:lvl3pPr><a:lvl4pPr marL="1371600" indent="0"><a:buNone/><a:defRPr sz="1600" b="1"/></a:lvl4pPr><a:lvl5pPr marL="1828800" indent
="0"><a:buNone/><a:defRPr sz="1600" b="1"/></a:lvl5pPr><a:lvl6pPr marL="2286000" indent="0"><a:buNone/><a:defRPr sz="1600" b="1"/></
a:lvl6pPr><a:lvl7pPr marL="2743200" indent="0"><a:buNone/><a:defRPr sz="1600" b="1"/></a:lvl7pPr><a:lvl8pPr marL="3200400" indent
="0"><a:buNone/><a:defRPr sz="1600" b="1"/></a:lvl8pPr><a:lvl9pPr marL="3657600" indent="0"><a:buNone/><a:defRPr sz="1600" b="1"/></
a:lvl9pPr></a:lstStyle><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master text styles</a:t></
a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Content Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr
><p:nvPr><p:ph sz="half" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="2174875"/><a:ext cx="4040188" cy="3951288
"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="2400"/></a:lvl1pPr><a:lvl2pPr><a:defRPr sz="2000"/></
a:lvl2pPr><a:lvl3pPr><a:defRPr sz="1800"/></a:lvl3pPr><a:lvl4pPr><a:defRPr sz="1600"/></a:lvl4pPr><a:lvl5pPr><a:defRPr sz="1600"/></
a:lvl5pPr><a:lvl6pPr><a:defRPr sz="1600"/></a:lvl6pPr><a:lvl7pPr><a:defRPr sz="1600"/></a:lvl7pPr><a:lvl8pPr><a:defRPr sz="1600"/></
a:lvl8pPr><a:lvl9pPr><a:defRPr sz="1600"/></a:lvl9pPr></a:lstStyle><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US" smtClean="0"/><
a:t>Click to edit Master text styles</a:t></a:r></a:p><a:p><a:pPr lvl="1"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Second level
</a:t></a:r></a:p><a:p><a:pPr lvl="2"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Third level</a:t></a:r></a:p><a:p><a:pPr lvl="
3"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Fourth level</a:t></a:r></a:p><a:p><a:pPr lvl="4"/><a:r><a:rPr lang="en-US" smtClean
="0"/><a:t>Fifth level</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="5" name="Text
 Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr
><p:spPr><a:xfrm><a:off x="4645025" y="1535113"/><a:ext cx="4041775" cy="639762"/></a:xfrm></p:spPr><p:txBody><a:bodyPr anchor="
b"/><a:lstStyle><a:lvl1pPr marL="0" indent="0"><a:buNone/><a:defRPr sz="2400" b="1"/></a:lvl1pPr><a:lvl2pPr marL="457200" indent
="0"><a:buNone/><a:defRPr sz="2000" b="1"/></a:lvl2pPr><a:lvl3pPr marL="91440 ...
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout6"></a>[function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>cbMakePptxLayout6 ( data )](#apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout6)
- description and source-code
```javascript
cbMakePptxLayout6 = function ( data ) {
	if ( !data || typeof data !== 'object' ) {
		data = {};
	} // Endif.

	return this.msPluginsApi.cbMakeMsOfficeBasicXml ( data ) + '<p:sld' + (data.isRealSlide ? '' : 'Layout') + ' xmlns:a="http://schemas
.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="
http://schemas.openxmlformats.org/presentationml/2006/main"' + (data.isRealSlide ? '' : ' type="titleOnly" preserve="1"') + '><p
:cSld name="Title Only"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm
><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p
:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr
/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master title style</a:t></a:r
><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Date Placeholder 2"/><p:cNvSpPr><a:spLocks
 noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt" sz="half" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle
/><a:p><a:fld id="{A76116CE-C4A3-4A05-B2D7-7C2E9A889C0F}" type="datetimeFigureOut"><a:rPr lang="en-US" smtClean="0"/><a:t>1/30/2017
</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Footer Placeholder 3"/><
p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody
><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="5" name="Slide Number
 Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr
><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{B1393E5F-521B-4CAD-9D3A-AE923D912DCE}" type="slidenum"><a:rPr lang="
en-US" smtClean="0"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><
a:masterClrMapping/></p:clrMapOvr></p:sld' + (data.isRealSlide ? '' : 'Layout') + '>';
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout7"></a>[function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>cbMakePptxLayout7 ( data )](#apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout7)
- description and source-code
```javascript
cbMakePptxLayout7 = function ( data ) {
	if ( !data || typeof data !== 'object' ) {
		data = {};
	} // Endif.

	return this.msPluginsApi.cbMakeMsOfficeBasicXml ( data ) + '<p:sld' + (data.isRealSlide ? '' : 'Layout') + ' xmlns:a="http://schemas
.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="
http://schemas.openxmlformats.org/presentationml/2006/main"' + (data.isRealSlide ? '' : ' type="blank" preserve="1"') + '><p:cSld
 name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="
0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="
2" name="Date Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt" sz="half" idx="10"/></p:nvPr></
p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{A76116CE-C4A3-4A05-B2D7-7C2E9A889C0F}" type="datetimeFigureOut
"><a:rPr lang="en-US" smtClean="0"/><a:t>1/30/2017</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr
><p:cNvPr id="3" name="Footer Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter"
idx="11"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp
><p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Number Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type
="sldNum" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{B1393E5F-521B-4CAD
-9D3A-AE923D912DCE}" type="slidenum"><a:rPr lang="en-US" smtClean="0"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></
p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld' + (data.isRealSlide ? '' : 'Layout') + '>';
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout8"></a>[function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>cbMakePptxLayout8 ( data )](#apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout8)
- description and source-code
```javascript
cbMakePptxLayout8 = function ( data ) {
	if ( !data || typeof data !== 'object' ) {
		data = {};
	} // Endif.

	return this.msPluginsApi.cbMakeMsOfficeBasicXml ( data ) + '<p:sld' + (data.isRealSlide ? '' : 'Layout') + ' xmlns:a="http://schemas
.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="
http://schemas.openxmlformats.org/presentationml/2006/main"' + (data.isRealSlide ? '' : ' type="objTx" preserve="1"') + '><p:cSld
 name="Content with Caption"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a
:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr
><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p
:spPr><a:xfrm><a:off x="457200" y="273050"/><a:ext cx="3008313" cy="1162050"/></a:xfrm></p:spPr><p:txBody><a:bodyPr anchor="b"/><
a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="2000" b="1"/></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr lang="en-US" smtClean="0"/><
a:t>Click to edit Master title style</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="
3" name="Content Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr
><a:xfrm><a:off x="3575050" y="273050"/><a:ext cx="5111750" cy="5853113"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle><a:
lvl1pPr><a:defRPr sz="3200"/></a:lvl1pPr><a:lvl2pPr><a:defRPr sz="2800"/></a:lvl2pPr><a:lvl3pPr><a:defRPr sz="2400"/></a:lvl3pPr
><a:lvl4pPr><a:defRPr sz="2000"/></a:lvl4pPr><a:lvl5pPr><a:defRPr sz="2000"/></a:lvl5pPr><a:lvl6pPr><a:defRPr sz="2000"/></a:lvl6pPr
><a:lvl7pPr><a:defRPr sz="2000"/></a:lvl7pPr><a:lvl8pPr><a:defRPr sz="2000"/></a:lvl8pPr><a:lvl9pPr><a:defRPr sz="2000"/></a:lvl9pPr
></a:lstStyle><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master text styles</a:t></a:r></a:p
><a:p><a:pPr lvl="1"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Second level</a:t></a:r></a:p><a:p><a:pPr lvl="2"/><a:r><a:rPr
 lang="en-US" smtClean="0"/><a:t>Third level</a:t></a:r></a:p><a:p><a:pPr lvl="3"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Fourth
 level</a:t></a:r></a:p><a:p><a:pPr lvl="4"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Fifth level</a:t></a:r><a:endParaRPr lang
="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Text Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></
p:cNvSpPr><p:nvPr><p:ph type="body" sz="half" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="1435100"/><a:ext
 cx="3008313" cy="4691063"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr marL="0" indent="0"><a:buNone/><a:defRPr
 sz="1400"/></a:lvl1pPr><a:lvl2pPr marL="457200" indent="0"><a:buNone/><a:defRPr sz="1200"/></a:lvl2pPr><a:lvl3pPr marL="914400"
indent="0"><a:buNone/><a:defRPr sz="1000"/></a:lvl3pPr><a:lvl4pPr marL="1371600" indent="0"><a:buNone/><a:defRPr sz="900"/></a:lvl4pPr
><a:lvl5pPr marL="1828800" indent="0"><a:buNone/><a:defRPr sz="900"/></a:lvl5pPr><a:lvl6pPr marL="2286000" indent="0"><a:buNone/><
a:defRPr sz="900"/></a:lvl6pPr><a:lvl7pPr marL="2743200" indent="0"><a:buNone/><a:defRPr sz="900"/></a:lvl7pPr><a:lvl8pPr marL="
3200400" indent="0"><a:buNone/><a:defRPr sz="900"/></a:lvl8pPr><a:lvl9pPr marL="3657600" indent="0"><a:buNone/><a:defRPr sz="900
"/></a:lvl9pPr></a:lstStyle><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master text styles</a
:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="5" name="Date Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p
:cNvSpPr><p:nvPr><p:ph type="dt" sz="half" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld
id="{A76116CE-C4A3-4A05-B2D7-7C2E9A889C0F}" type="datetimeFigureOut"><a:rPr lang="en-US" smtClean="0"/><a:t>1/30/2017</a:t></a:fld
><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="6" name="F ...
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout9"></a>[function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>cbMakePptxLayout9 ( data )](#apidoc.element.officegen.pptxplg_layouts.prototype.cbMakePptxLayout9)
- description and source-code
```javascript
cbMakePptxLayout9 = function ( data ) {
	if ( !data || typeof data !== 'object' ) {
		data = {};
	} // Endif.

	return this.msPluginsApi.cbMakeMsOfficeBasicXml ( data ) + '<p:sld' + (data.isRealSlide ? '' : 'Layout') + ' xmlns:a="http://schemas
.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="
http://schemas.openxmlformats.org/presentationml/2006/main"' + (data.isRealSlide ? '' : ' type="picTx" preserve="1"') + '><p:cSld
 name="Picture with Caption"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a
:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr
><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p
:spPr><a:xfrm><a:off x="1792288" y="4800600"/><a:ext cx="5486400" cy="566738"/></a:xfrm></p:spPr><p:txBody><a:bodyPr anchor="b"/><
a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="2000" b="1"/></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr lang="en-US" smtClean="0"/><
a:t>Click to edit Master title style</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="
3" name="Picture Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="pic" idx="1"/></p:nvPr></p:nvSpPr
><p:spPr><a:xfrm><a:off x="1792288" y="612775"/><a:ext cx="5486400" cy="4114800"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle
><a:lvl1pPr marL="0" indent="0"><a:buNone/><a:defRPr sz="3200"/></a:lvl1pPr><a:lvl2pPr marL="457200" indent="0"><a:buNone/><a:defRPr
 sz="2800"/></a:lvl2pPr><a:lvl3pPr marL="914400" indent="0"><a:buNone/><a:defRPr sz="2400"/></a:lvl3pPr><a:lvl4pPr marL="1371600
" indent="0"><a:buNone/><a:defRPr sz="2000"/></a:lvl4pPr><a:lvl5pPr marL="1828800" indent="0"><a:buNone/><a:defRPr sz="2000"/></
a:lvl5pPr><a:lvl6pPr marL="2286000" indent="0"><a:buNone/><a:defRPr sz="2000"/></a:lvl6pPr><a:lvl7pPr marL="2743200" indent="0"><
a:buNone/><a:defRPr sz="2000"/></a:lvl7pPr><a:lvl8pPr marL="3200400" indent="0"><a:buNone/><a:defRPr sz="2000"/></a:lvl8pPr><a:lvl9pPr
 marL="3657600" indent="0"><a:buNone/><a:defRPr sz="2000"/></a:lvl9pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody
></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Text Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type
="body" sz="half" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="1792288" y="5367338"/><a:ext cx="5486400" cy="804862"/></
a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr marL="0" indent="0"><a:buNone/><a:defRPr sz="1400"/></a:lvl1pPr><a:lvl2pPr
 marL="457200" indent="0"><a:buNone/><a:defRPr sz="1200"/></a:lvl2pPr><a:lvl3pPr marL="914400" indent="0"><a:buNone/><a:defRPr sz
="1000"/></a:lvl3pPr><a:lvl4pPr marL="1371600" indent="0"><a:buNone/><a:defRPr sz="900"/></a:lvl4pPr><a:lvl5pPr marL="1828800" indent
="0"><a:buNone/><a:defRPr sz="900"/></a:lvl5pPr><a:lvl6pPr marL="2286000" indent="0"><a:buNone/><a:defRPr sz="900"/></a:lvl6pPr><
a:lvl7pPr marL="2743200" indent="0"><a:buNone/><a:defRPr sz="900"/></a:lvl7pPr><a:lvl8pPr marL="3200400" indent="0"><a:buNone/><
a:defRPr sz="900"/></a:lvl8pPr><a:lvl9pPr marL="3657600" indent="0"><a:buNone/><a:defRPr sz="900"/></a:lvl9pPr></a:lstStyle><a:p
><a:pPr lvl="0"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master text styles</a:t></a:r></a:p></p:txBody></p:sp
><p:sp><p:nvSpPr><p:cNvPr id="5" name="Date Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt"
sz="half" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{A76116CE-C4A3-4A05-B2D7-7C2E9A889C0F
}" type="datetimeFigureOut"><a:rPr lang="en-US" smtClean="0"/><a:t>1/30/2017</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:
txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="6" name="Footer Placeholder 5"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><
p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr><p:spPr/>< ...
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.pptxplg_layouts.prototype.extendPptxApi"></a>[function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>extendPptxApi ( docObj )](#apidoc.element.officegen.pptxplg_layouts.prototype.extendPptxApi)
- description and source-code
```javascript
extendPptxApi = function ( docObj ) {
	var funcThis = this;

	/**
	 * Create a new title layout based slide.
	 * @param {*} title Optional title string.
	 * @param {*} subTitle Optional sub-title string.
	 * @return The new slide object.
	 */
	docObj.makeTitleSlide = function ( title, subTitle ) {
		return docObj.makeNewSlide ({
			useLayout: 'title',
			title: title,
			subTitle: subTitle
		});
	};

	/**
	 * Create a new obj layout based slide.
	 * @param {*} title Optional title string.
	 * @param {*} objData Optional obj data.
	 * @return The new slide object.
	 */
	docObj.makeObjSlide = function ( title, objData ) {
		return docObj.makeNewSlide ({
			useLayout: 'obj',
			title: title,
			objData: objData
		});
	};

	/**
	 * Create a new secHead layout based slide.
	 * @param {*} title Optional title string.
	 * @param {*} subTitle Optional sub-title string.
	 * @return The new slide object.
	 */
	docObj.makeSecHeadSlide = function ( title, subTitle ) {
		return docObj.makeNewSlide ({
			useLayout: 'secHead',
			title: title,
			subTitle: subTitle
		});
	};

	/**
	 * Create a new custom layout.
	 * @param {string} layoutName New layout name.
	 * @return The new layout object.
	 */
	docObj.makeNewLayout = function ( layoutName ) {
		if ( !funcThis.pptxData.slideLayouts ) {
			funcThis.pptxData.slideLayouts = [];
		} // Endif.

		if ( !layoutName || typeof layoutName !== 'string' ) {
			throw 'Invalid layout name.';
		} // Endif.

		var slideLayouts = funcThis.pptxData.slideLayouts;

		var newLayout = {
			slide: {
				getPageNumber: function () { return 1; },
				layoutName: layoutName
			},
			data: []
		}

		// Register a new layout object:
		slideLayouts.push (
			{
				// We are using the slide resource generator to make the new layout since that the format is almost the same:
				execCb: function ( data ) { return funcThis.pluginsman.genobj.cbMakePptxSlide ( data ); },
				cbData: newLayout
			}
		);

		// (we will add here API to add objects into the layout, but first we must move the slide API into a new file)
		// BMK_TODO:

		return newLayout;
	};
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.pptxplg_layouts.prototype.extendPptxSlideApi"></a>[function <span class="apidocSignatureSpan">officegen.pptxplg_layouts.prototype.</span>extendPptxSlideApi ( docData )](#apidoc.element.officegen.pptxplg_layouts.prototype.extendPptxSlideApi)
- description and source-code
```javascript
extendPptxSlideApi = function ( docData ) {
	var funcThis = this;
	var newSlide = docData.page; // The new slide's API.
	var slideData = docData.pageData; // Place here data related to the slide.
	var slideOptions = docData.slideOptions; // Slide creation options.

	// We'll add API only if the user requested it:
	if ( slideOptions && typeof slideOptions === 'object' && slideOptions.useLayout ) {
		// Support for the standard office layout1:
		if ( slideOptions.useLayout === 'title' || slideOptions.useLayout === 'obj' || slideOptions.useLayout === 'secHead' ) {
			var resCb = this.pluginsman.genobj.cbMakePptxLayout1;
			if ( slideOptions.useLayout === 'obj' ) {
				resCb = function ( data ) {
					return funcThis.cbMakePptxLayout2 ( data );
				};
			} // Endif.

			if ( slideOptions.useLayout === 'secHead' ) {
				resCb = function ( data ) {
					return funcThis.cbMakePptxLayout3 ( data );
				};
			} // Endif.

			newSlide.useLayout = {
				mkResCb: resCb,
				slide: newSlide,
				ph1: [],
				ph2: [],
				ft: [],
				useDate: new Date (),
				isRealSlide: true,
				isFooter: false,
				isSlideNum: true,
				isDate: true
			};

			newSlide.setTitle = function ( titleStr ) {
				if ( typeof titleStr === 'string' ) {
					newSlide.useLayout.ph1 = [
						{ text: titleStr }
					];

				} else if ( typeof titleStr === 'object' ) {
					newSlide.useLayout.ph1 = titleStr;
				} // Endif.
			};

			if ( slideOptions.useLayout === 'obj' ) {
				newSlide.setObjData = function ( objData ) {
					if ( typeof objData === 'string' ) {
						newSlide.useLayout.ph2 = [
							{ text: objData }
						];

					} else if ( typeof objData === 'object' ) {
						newSlide.useLayout.ph2 = objData;
					} // Endif.
				};

			} else {
				newSlide.setSubTitle = function ( titleStr ) {
					if ( typeof titleStr === 'string' ) {
						newSlide.useLayout.ph2 = [
							{ text: titleStr }
						];

					} else if ( typeof titleStr === 'object' ) {
						newSlide.useLayout.ph2 = titleStr;
					} // Endif.
				};
			} // Endif.

			newSlide.setFooter = function ( footerStr ) {
				newSlide.useLayout.isFooter = true;

				if ( typeof footerStr === 'string' ) {
					newSlide.useLayout.ft = [
						{ text: footerStr }
					];

				} else if ( typeof footerStr === 'object' ) {
					newSlide.useLayout.ft = footerStr;
				} // Endif.
			};

			if ( slideOptions.title ) {
				newSlide.setTitle ( slideOptions.title );
			} // Endif.

			if ( slideOptions.useLayout === 'obj' ) {
				if ( slideOptions.objData ) {
					newSlide.setObjData ( slideOptions.objData );
				} // Endif.

			} else {
				if ( slideOptions.subTitle ) {
					newSlide.setSubTitle ( slideOptions.subTitle );
				} // Endif.
			} // Endif.

			if ( slideOptions.footer ) {
				newSlide.setFooter ( slideOptions.footer );
			} // Endif.

		} else {
			throw ( 'Invalid office layout.' );
		} // Endif.
	} // Endif.
}
```
- example usage
```shell
n/a
```



# <a name="apidoc.module.officegen.pptxplg_speakernotes"></a>[module officegen.pptxplg_speakernotes](#apidoc.module.officegen.pptxplg_speakernotes)

#### <a name="apidoc.element.officegen.pptxplg_speakernotes.pptxplg_speakernotes"></a>[function <span class="apidocSignatureSpan">officegen.</span>pptxplg_speakernotes ( pluginsman )](#apidoc.element.officegen.pptxplg_speakernotes.pptxplg_speakernotes)
- description and source-code
```javascript
function MakeSpknotesPlugin( pluginsman ) {
	var funcThis = this;

	if ( pluginsman.docType !== 'pptx' && pluginsman.docType !== 'ppsx' ) {
		throw "[pptx-speakernotes] This plugin supporting only PowerPoint based documents.";
	} // Endif.

	this.ogPluginsApi = pluginsman.ogPluginsApi; // Generic officegen API for plugins.
	this.msPluginsApi = pluginsman.genPrivate.plugs.type.msoffice; // msoffice plugins API.
	this.pluginsman = pluginsman; // Document type specific plugins API.

	this.pptxData = pluginsman.getDataStorage (); // Here you can store any temporary data needed for generating the document and depending
 on the data filled by the user.

	this.mainPath = pluginsman.genPrivate.features.type.msoffice.main_path; // The "folder" name inside the document zip that all the
 specific resources of this document type are stored.
	this.mainPathFile = pluginsman.genPrivate.features.type.msoffice.main_path_file; // The name of the main real xml resource of this
 document.
	this.relsMain = pluginsman.genPrivate.type.msoffice.rels_main; // Main rels file.
	this.relsApp = pluginsman.genPrivate.type.msoffice.rels_app; // Main rels file inside the specific document type "folder".
	this.filesList = pluginsman.genPrivate.type.msoffice.files_list; // Resources list xml.
	this.srcFilesList = pluginsman.genPrivate.type.msoffice.src_files_list; // For storing extra files inside the document zip.

	//
	// Catch events inside the document so we can extent it:
	//

	// We want to extend the slide object API:
	pluginsman.registerCallback ( 'newPage', function ( docData ) { funcThis.extendPptxSlideApi ( docData ); } );

	// This event tell us that the generator is about to start working:
	pluginsman.registerCallback ( 'beforeGen', function ( docObj ) { funcThis.beforeGen ( docObj ); } );

	// This event tell us to add our xml code to the main presentation file:
	pluginsman.registerCallback ( 'presentationGen', function ( docData ) { funcThis.presentationGen ( docData ); } );

	return this;
}
```
- example usage
```shell
n/a
```



# <a name="apidoc.module.officegen.pptxplg_speakernotes.prototype"></a>[module officegen.pptxplg_speakernotes.prototype](#apidoc.module.officegen.pptxplg_speakernotes.prototype)

#### <a name="apidoc.element.officegen.pptxplg_speakernotes.prototype.beforeGen"></a>[function <span class="apidocSignatureSpan">officegen.pptxplg_speakernotes.prototype.</span>beforeGen ( docObj )](#apidoc.element.officegen.pptxplg_speakernotes.prototype.beforeGen)
- description and source-code
```javascript
beforeGen = function ( docObj ) {
	var funcThis = this;

	var notesCount = 0;
	this.pluginsman.genPrivate.pages.forEach ( function ( value, indexId ) {
		if ( value.speakerNoteMsg && typeof value.speakerNoteMsg === 'object' && value.speakerNoteMsg.length ) {
			// Add each speaker note:
			funcThis.ogPluginsApi.intAddAnyResourceToParse ( funcThis.mainPath + '\\notesSlides\\notesSlide' + (notesCount + 1) + '.xml', '
buffer', { slide: value, slideNum: indexId + 1 }, function ( data ) { return funcThis.cbMakePptxSpkNote ( data ); }, false );
			funcThis.ogPluginsApi.intAddAnyResourceToParse ( funcThis.mainPath + '\\notesSlides\\_rels\\notesSlide' + (notesCount + 1) + '.
xml.rels', 'buffer', [
				{
					type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster',
					target: '../notesMasters/notesMaster1.xml'
				},
				{
					type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
					target: '../slides/slide' + (indexId + 1) + '.xml'
				}
			], funcThis.pluginsman.genPrivate.plugs.type.msoffice.cbMakeRels, false );

			// We'll connect the slide to this note:
			value.rels.push ({
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
				target: '../notesSlides/notesSlide' + (notesCount + 1) + '.xml'
			});

			// Add this speaker note to the list of files in the document:
			funcThis.filesList.push ({
				name: '/' + funcThis.mainPath + '/notesSlides/notesSlide' + (notesCount + 1) + '.xml',
				type: 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
				clear: 'generate' // Placing 'generate' here means that officegen will destroy this entry in the files list after finishing
to generate the document.
			});

			notesCount++;
		} // Endif.
	});

	this.pptxData.haveNotes = notesCount;
	if ( notesCount ) {
		// Speaker notes - master:
		this.ogPluginsApi.intAddAnyResourceToParse ( this.mainPath + '\\notesMasters\\notesMaster1.xml', 'buffer', null, function ( data
 ) { return funcThis.cbMakePptxNotesMasters ( data ); }, false );
		this.ogPluginsApi.intAddAnyResourceToParse ( this.mainPath + '\\notesMasters\\_rels\\notesMaster1.xml.rels', 'buffer', [
			{
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
				target: '../theme/theme2.xml'
			}
		], this.pluginsman.genPrivate.plugs.type.msoffice.cbMakeRels, false );

		// Add a rel entry to the main pptx rels:
		this.relsApp.push ({
			target: 'notesMasters/notesMaster1.xml',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster',
			clear: 'generate' // Placing 'generate' here means that officegen will destroy this entry in the files list after finishing to
 generate the document.
		},
		{
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
			target: 'theme/theme2.xml',
			clear: 'generate'
		});

		funcThis.ogPluginsApi.intAddAnyResourceToParse ( funcThis.mainPath + '\\theme\\theme2.xml', 'buffer', null, function ( data ) {
return funcThis.cbMakeTheme2 ( data ); }, false );

		// Add the notes master to the list of files in the document:
		this.filesList.push ({
			name: '/' + this.mainPath + '/notesMasters/notesMaster1.xml',
			type: 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml',
			clear: 'generate' // Placing 'generate' here means that officegen will destroy this entry in the files list after finishing to
 generate the document.
		},
		{
			name: '/' + this.mainPath + '/theme/theme2.xml',
			type: 'application/vnd.openxmlformats-officedocument.theme+xml',
			clear: 'generate' // Placing 'generate' here means that officegen will destroy this entry in the files list after finishing to
 generate the document.
		});
	} // Endif.
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.pptxplg_speakernotes.prototype.cbMakePptxNotesMasters"></a>[function <span class="apidocSignatureSpan">officegen.pptxplg_speakernotes.prototype.</span>cbMakePptxNotesMasters ( data )](#apidoc.element.officegen.pptxplg_speakernotes.prototype.cbMakePptxNotesMasters)
- description and source-code
```javascript
cbMakePptxNotesMasters = function ( data ) {
	return this.msPluginsApi.cbMakeMsOfficeBasicXml ( data ) + '<p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml
/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats
.org/presentationml/2006/main"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr
><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:
chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Header Placeholder 1"/><p:
cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="hdr" sz="quarter"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0
" y="0"/><a:ext cx="2971800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert
="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200"/></a
:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Date Placeholder
 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="
3884613" y="0"/><a:ext cx="2971800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr
 vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200"/></
a:lvl1pPr></a:lstStyle><a:p><a:fld id="{AA6EBAA8-ADDE-4290-A9BE-D64DC1002B1F}" type="datetimeFigureOut"><a:rPr lang="en-US" smtClean
="0"/><a:t>2/3/2017</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide
 Image Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx
="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="1143000" y="685800"/><a:ext cx="4572000" cy="3429000"/></a:xfrm><a:prstGeom
prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr
><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><
a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="5" name="Notes Placeholder 4"/><p:cNvSpPr><a:spLocks
 noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="
4343400"/><a:ext cx="5486400" cy="4114800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr
 vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr
lvl="0"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master text styles</a:t></a:r></a:p><a:p><a:pPr lvl="1"/><a:r
><a:rPr lang="en-US" smtClean="0"/><a:t>Second level</a:t></a:r></a:p><a:p><a:pPr lvl="2"/><a:r><a:rPr lang="en-US" smtClean="0"/><
a:t>Third level</a:t></a:r></a:p><a:p><a:pPr lvl="3"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Fourth level</a:t></a:r></a:p><
a:p><a:pPr lvl="4"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Fifth level</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody
></p:sp><p:sp><p:nvSpPr><p:cNvPr id="6" name="Footer Placeholder 5"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph
type="ftr" sz="quarter" idx="4"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="8685213"/><a:ext cx="2971800" cy="457200"/></
a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440
" bIns="45720" rtlCol="0" anchor="b"/><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr
 lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="7" name="Slide Number Placeholder 6"/><p:cNvSpPr><a:spLocks
noGrp="1"/></p:cNvSpPr><p:nvP ...
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.pptxplg_speakernotes.prototype.cbMakePptxSpkNote"></a>[function <span class="apidocSignatureSpan">officegen.pptxplg_speakernotes.prototype.</span>cbMakePptxSpkNote ( data )](#apidoc.element.officegen.pptxplg_speakernotes.prototype.cbMakePptxSpkNote)
- description and source-code
```javascript
cbMakePptxSpkNote = function ( data ) {
	var dataOut = this.msPluginsApi.cbMakeMsOfficeBasicXml ( data ) + '<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml
/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats
.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:
grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:
sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p
:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder
 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr
><a:normAutofit/></a:bodyPr><a:lstStyle/>';

	// Generate the speaker note lines for this slide:
	dataOut += this.pluginsman.genobj.cMakePptxOutTextP ( '', data.slide.speakerNoteMsg, {}, data.slide.slide );

	dataOut += '</p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Number Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1
"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><
a:p><a:fld id="{E165F6F6-A11E-4DE0-B2DB-03E2E9A5946D}" type="slidenum"><a:rPr lang="en-US" smtClean="0"/><a:t>' + data.slideNum + '</
a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr
></p:notes>';
	return dataOut;
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.pptxplg_speakernotes.prototype.cbMakeTheme2"></a>[function <span class="apidocSignatureSpan">officegen.pptxplg_speakernotes.prototype.</span>cbMakeTheme2 ( data )](#apidoc.element.officegen.pptxplg_speakernotes.prototype.cbMakeTheme2)
- description and source-code
```javascript
cbMakeTheme2 = function ( data ) {
	return this.msPluginsApi.cbMakeMsOfficeBasicXml ( data ) + '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/
main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1
><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></
a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="
9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><
a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink
></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font
 script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="
Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font
 script="Thai" typeface="Angsana New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font
script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script
="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font
 script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface
="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha
"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="
Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface
="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:majorFont
><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font
 script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab
" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Cordia New"/><a:font script="Ethi" typeface
="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh
"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a
:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt"
typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script
="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font
script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font
script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><
a:font script="Uigh" typeface="Microsoft Uighur"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill
><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val
="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod
 val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a
:schemeClr></a:gs>< ...
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.pptxplg_speakernotes.prototype.extendPptxSlideApi"></a>[function <span class="apidocSignatureSpan">officegen.pptxplg_speakernotes.prototype.</span>extendPptxSlideApi ( docData )](#apidoc.element.officegen.pptxplg_speakernotes.prototype.extendPptxSlideApi)
- description and source-code
```javascript
extendPptxSlideApi = function ( docData ) {
	var funcThis = this;
	var newSlide = docData.page; // The new slide's API.
	var slideData = docData.pageData; // Place here data related to the slide.

	/**
	 * Set the speaker notes value.
	 * @param {*} messageText - speaker note message.
	 * @param {*} isAppend - is true to append the message as another line.
	 */
	newSlide.setSpeakerNote = function ( messageText, isAppend ) {
		if ( !messageText ) {
			slideData.speakerNoteMsg = [];

		} else if ( typeof messageText === 'string' ) {
			if ( isAppend ) {
				if ( !slideData.speakerNoteMsg || typeof slideData.speakerNoteMsg !== 'object' ) {
					slideData.speakerNoteMsg = [];

				} else if ( slideData.speakerNoteMsg.length ) {
					slideData.speakerNoteMsg.push ( { text: '\n' + messageText } );
					return;
				} // Endif.

				slideData.speakerNoteMsg.push ( { text: messageText } );

			} else {
				slideData.speakerNoteMsg = [
					{ text: messageText }
				];
			} // Endif.

		} else if ( typeof messageText === 'object' ) {
			slideData.speakerNoteMsg = messageText;
		} // Endif.
	};
}
```
- example usage
```shell
n/a
```

#### <a name="apidoc.element.officegen.pptxplg_speakernotes.prototype.presentationGen"></a>[function <span class="apidocSignatureSpan">officegen.pptxplg_speakernotes.prototype.</span>presentationGen ( docData )](#apidoc.element.officegen.pptxplg_speakernotes.prototype.presentationGen)
- description and source-code
```javascript
presentationGen = function ( docData ) {
	var funcThis = this;

	// Check if we have speaker notes:
	if ( this.pptxData.haveNotes ) {
		docData.data += '<p:notesMasterIdLst><p:notesMasterId r:id="rId' + (this.pluginsman.genPrivate.pages.length + 2) + '"/></p:notesMasterIdLst
>'
	} // Endif.
}
```
- example usage
```shell
n/a
```



# <a name="apidoc.module.officegen.pptxplg_widescreen"></a>[module officegen.pptxplg_widescreen](#apidoc.module.officegen.pptxplg_widescreen)

#### <a name="apidoc.element.officegen.pptxplg_widescreen.pptxplg_widescreen"></a>[function <span class="apidocSignatureSpan">officegen.</span>pptxplg_widescreen ( pluginsman )](#apidoc.element.officegen.pptxplg_widescreen.pptxplg_widescreen)
- description and source-code
```javascript
function MakeWidescreenPlugin( pluginsman ) {
	var funcThis = this;

	if ( pluginsman.docType !== 'pptx' && pluginsman.docType !== 'ppsx' ) {
		throw "[pptx-widescreen] This plugin supporting only PowerPoint based documents.";
	} // Endif.

	this.pluginsman = pluginsman;
	this.pptxData = pluginsman.getDataStorage ();

	// We want to extend the main API of the pptx document object:
	pluginsman.registerCallback ( 'makeDocApi', function ( docObj ) { funcThis.extendPptxApi ( docObj ); } );
	return this;
}
```
- example usage
```shell
n/a
```



# <a name="apidoc.module.officegen.pptxplg_widescreen.prototype"></a>[module officegen.pptxplg_widescreen.prototype](#apidoc.module.officegen.pptxplg_widescreen.prototype)

#### <a name="apidoc.element.officegen.pptxplg_widescreen.prototype.extendPptxApi"></a>[function <span class="apidocSignatureSpan">officegen.pptxplg_widescreen.prototype.</span>extendPptxApi ( docObj )](#apidoc.element.officegen.pptxplg_widescreen.prototype.extendPptxApi)
- description and source-code
```javascript
extendPptxApi = function ( docObj ) {
	var funcThis = this;

	docObj.setSlideSize = function ( cx, cy, sizeType ) {
		if ( cx && typeof cx === 'number' ) {
			funcThis.pptxData.pptWidth = cx * funcThis.pptxData.EMUS_PER_PT;
		} // Endif.

		if ( cy && typeof cy === 'number' ) {
			funcThis.pptxData.pptHeight = cy * funcThis.pptxData.EMUS_PER_PT;
		} // Endif.

		/*
		Supported types:

		'35mm'
		'A3'
		'A4'
		'B4ISO'
		'B4JIS'
		'B5ISO'
		'B5JIS'
		'banner'
		'custom'
		'hagakiCard'
		'ledger'
		'letter'
		'overhead'
		'screen16x10'
		'screen16x9'
		'screen4x3'
		*/

		if ( sizeType && typeof sizeType === 'string' ) {
			funcThis.pptxData.pptType = sizeType;
		} // Endif.
	};

	docObj.setWidescreen = function ( wide ) {
		funcThis.pptxData.pptWidth 	= wide ? (960 * funcThis.pptxData.EMUS_PER_PT) 		: funcThis.pptxData.pptWidth;
		funcThis.pptxData.pptType 	= wide ? 'screen16x9' 	: 'screen4x3';
	};
}
```
- example usage
```shell
n/a
```



# misc
- this document was created with [utility2](https://github.com/kaizhu256/node-utility2)
