(function webpackUniversalModuleDefinition(root, factory) {
	if(typeof exports === 'object' && typeof module === 'object')
		module.exports = factory(require("JSZip"));
	else if(typeof define === 'function' && define.amd)
		define(["JSZip"], factory);
	else if(typeof exports === 'object')
		exports["docx"] = factory(require("JSZip"));
	else
		root["docx"] = factory(root["JSZip"]);
})(window, function(__WEBPACK_EXTERNAL_MODULE_jszip__) {
return /******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = "./src/docx-preview.ts");
/******/ })
/************************************************************************/
/******/ ({

/***/ "./src/document-parser.ts":
/*!********************************!*\
  !*** ./src/document-parser.ts ***!
  \********************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var dom_1 = __webpack_require__(/*! ./dom/dom */ "./src/dom/dom.ts");
var utils = __webpack_require__(/*! ./utils */ "./src/utils.ts");
var document_1 = __webpack_require__(/*! ./document */ "./src/document.ts");
var common_1 = __webpack_require__(/*! ./dom/common */ "./src/dom/common.ts");
var common_2 = __webpack_require__(/*! ./parser/common */ "./src/parser/common.ts");
var paragraph_1 = __webpack_require__(/*! ./parser/paragraph */ "./src/parser/paragraph.ts");
var section_1 = __webpack_require__(/*! ./parser/section */ "./src/parser/section.ts");
var text_1 = __webpack_require__(/*! ./elements/text */ "./src/elements/text.ts");
var symbol_1 = __webpack_require__(/*! ./elements/symbol */ "./src/elements/symbol.ts");
var break_1 = __webpack_require__(/*! ./elements/break */ "./src/elements/break.ts");
var tab_1 = __webpack_require__(/*! ./elements/tab */ "./src/elements/tab.ts");
var hyperlink_1 = __webpack_require__(/*! ./elements/hyperlink */ "./src/elements/hyperlink.ts");
var run_1 = __webpack_require__(/*! ./elements/run */ "./src/elements/run.ts");
var bookmark_1 = __webpack_require__(/*! ./elements/bookmark */ "./src/elements/bookmark.ts");
var cell_1 = __webpack_require__(/*! ./elements/cell */ "./src/elements/cell.ts");
var table_1 = __webpack_require__(/*! ./elements/table */ "./src/elements/table.ts");
var row_1 = __webpack_require__(/*! ./elements/row */ "./src/elements/row.ts");
var paragraph_2 = __webpack_require__(/*! ./elements/paragraph */ "./src/elements/paragraph.ts");
var image_1 = __webpack_require__(/*! ./elements/image */ "./src/elements/image.ts");
var drawing_1 = __webpack_require__(/*! ./elements/drawing */ "./src/elements/drawing.ts");
var xml_serialize_1 = __webpack_require__(/*! ./parser/xml-serialize */ "./src/parser/xml-serialize.ts");
var section_2 = __webpack_require__(/*! ./elements/section */ "./src/elements/section.ts");
exports.autos = {
    shd: "white",
    color: "black",
    highlight: "transparent"
};
var DocumentParser = (function () {
    function DocumentParser() {
        this.skipDeclaration = true;
        this.ignoreWidth = false;
        this.debug = false;
    }
    DocumentParser.prototype.parseContentTypeFile = function (xmlString) {
        var xoverrides = xml.parse(xmlString, this.skipDeclaration);
        var parts = new Map();
        xml.elements(xoverrides).filter(function (element) { return element.tagName === 'Override'; }).forEach(function (overrideElement) {
            var elementContentType = xml.stringAttr(overrideElement, "ContentType");
            if (Object.values(document_1.PartType).includes(elementContentType)) {
                var part = Object.entries(document_1.PartType).find(function (_a) {
                    var _ = _a[0], contentType = _a[1];
                    return contentType === elementContentType;
                })[0];
                parts.set(document_1.PartType[part], xml.stringAttr(overrideElement, "PartName"));
            }
        });
        return parts;
    };
    DocumentParser.prototype.parseDocumentRelationsFile = function (xmlString) {
        var xrels = xml.parse(xmlString, this.skipDeclaration);
        return xml.elements(xrels).map(function (c) { return ({
            id: xml.stringAttr(c, "Id"),
            type: values.valueOfRelType(c),
            target: xml.stringAttr(c, "Target"),
        }); });
    };
    DocumentParser.prototype.parseFontTable = function (xmlString) {
        var xfonts = xml.parse(xmlString, this.skipDeclaration);
        return xml.elements(xfonts).map(function (c) { return ({
            name: xml.stringAttr(c, "name"),
            fontKey: xml.elementStringAttr(c, "embedRegular", "fontKey"),
            refId: xml.elementStringAttr(c, "embedRegular", "id")
        }); });
    };
    DocumentParser.prototype.parseDocumentFile = function (xmlString) {
        var _this = this;
        var result = {
            type: dom_1.DomType.Document,
            children: [],
            style: {},
            props: null
        };
        var xbody = xml.byTagName(xml.parse(xmlString, this.skipDeclaration), "body");
        var section = new section_2.Section();
        xml.foreach(xbody, function (elem) {
            switch (elem.localName) {
                case "p":
                    var paragraph = _this.parseParagraph(elem);
                    if (paragraph.props.sectionProps) {
                        section.props = paragraph.props.sectionProps;
                        result.children.push(section);
                        section = new section_2.Section();
                    }
                    section.children.push(paragraph);
                    _this.checkAndMergeConsecutivePragraphBorder(section.children);
                    break;
                case "tbl":
                    section.children.push(_this.parseTable(elem));
                    break;
                case "sectPr":
                    section.props = section_1.parseSectionProperties(elem);
                    result.children.push(section);
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.checkAndMergeConsecutivePragraphBorder = function (children) {
        if (children.length > 1 && children[children.length - 1] instanceof paragraph_2.Paragraph && children[children.length - 2] instanceof paragraph_2.Paragraph) {
            var p1 = children[children.length - 2];
            var p2 = children[children.length - 1];
            if (p2.style && p1.style &&
                p2.style['bdr-left'] === p1.style['bdr-left'] &&
                p2.style['bdr-right'] === p1.style['bdr-right'] &&
                p2.style['bdr-top'] === p1.style['bdr-top'] &&
                p2.style['bdr-bottom'] === p1.style['bdr-bottom'] &&
                p2.style['bdr-between'] === p1.style['bdr-between']) {
                p2.style['border-top'] = p1.style['bdr-between'];
                p1.style['border-bottom'] = '';
            }
        }
    };
    DocumentParser.prototype.parseHeaderOrFooter = function (xmlString) {
        var _this = this;
        var result = {
            type: dom_1.DomType.HeaderOrFooter,
            children: [],
            style: {},
            props: null
        };
        var xbody = xml.parse(xmlString, this.skipDeclaration);
        xml.foreach(xbody, function (elem) {
            switch (elem.localName) {
                case "p":
                    result.children.push(_this.parseParagraph(elem));
                    _this.checkAndMergeConsecutivePragraphBorder(result.children);
                    break;
                case "tbl":
                    result.children.push(_this.parseTable(elem));
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseStylesFile = function (xmlString) {
        var _this = this;
        var result = [];
        var xstyles = xml.parse(xmlString, this.skipDeclaration);
        xml.foreach(xstyles, function (n) {
            switch (n.localName) {
                case "style":
                    result.push(_this.parseStyle(n));
                    break;
                case "docDefaults":
                    result.push(_this.parseDefaultStyles(n));
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseDefaultStyles = function (node) {
        var _this = this;
        var result = {
            id: null,
            name: null,
            target: null,
            basedOn: null,
            styles: []
        };
        xml.foreach(node, function (c) {
            switch (c.localName) {
                case "rPrDefault":
                    var rPr = xml.byTagName(c, "rPr");
                    if (rPr)
                        result.styles.push({
                            target: "span",
                            values: _this.parseDefaultProperties(rPr, {})
                        });
                    break;
                case "pPrDefault":
                    var pPr = xml.byTagName(c, "pPr");
                    if (pPr)
                        result.styles.push({
                            target: "p",
                            values: _this.parseDefaultProperties(pPr, {})
                        });
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseCommonProperties = function (elem, props) {
        if (elem.namespaceURI != common_1.ns.wordml)
            return;
        switch (elem.localName) {
            case "color":
                props.color = common_2.colorAttr(elem, elem.namespaceURI, "val");
                break;
            case "sz":
                props.fontSize = common_2.lengthAttr(elem, elem.namespaceURI, "val", common_2.LengthUsage.FontSize);
                break;
        }
    };
    DocumentParser.prototype.parseStyle = function (node) {
        var _this = this;
        var result = {
            id: xml.className(node, "styleId"),
            isDefault: xml.boolAttr(node, "default"),
            name: null,
            target: null,
            basedOn: null,
            styles: [],
            linked: null
        };
        switch (xml.stringAttr(node, "type")) {
            case "paragraph":
                result.target = "p";
                break;
            case "table":
                result.target = "table";
                break;
            case "character":
                result.target = "span";
                break;
        }
        xml.foreach(node, function (n) {
            switch (n.localName) {
                case "basedOn":
                    result.basedOn = xml.className(n, "val");
                    break;
                case "name":
                    result.name = xml.stringAttr(n, "val");
                    break;
                case "link":
                    result.linked = xml.className(n, "val");
                    break;
                case "aliases":
                    result.aliases = xml.stringAttr(n, "val").split(",");
                    break;
                case "pPr":
                    result.styles.push({
                        target: "p",
                        values: _this.parseDefaultProperties(n, {})
                    });
                    break;
                case "rPr":
                    result.styles.push({
                        target: "span",
                        values: _this.parseDefaultProperties(n, {})
                    });
                    break;
                case "tblPr":
                case "tcPr":
                    result.styles.push({
                        target: "td",
                        values: _this.parseDefaultProperties(n, {})
                    });
                    break;
                case "tblStylePr":
                    for (var _i = 0, _a = _this.parseTableStyle(n); _i < _a.length; _i++) {
                        var s = _a[_i];
                        result.styles.push(s);
                    }
                    break;
                case "rsid":
                case "qFormat":
                case "hidden":
                case "semiHidden":
                case "unhideWhenUsed":
                case "autoRedefine":
                case "uiPriority":
                    break;
                default:
                    _this.debug && console.warn("DOCX: Unknown style element: " + n.localName);
            }
        });
        return result;
    };
    DocumentParser.prototype.parseTableStyle = function (node) {
        var _this = this;
        var result = [];
        var type = xml.stringAttr(node, "type");
        var selector = "";
        switch (type) {
            case "firstRow":
                selector = "tr.first-row td";
                break;
            case "lastRow":
                selector = "tr.last-row td";
                break;
            case "firstCol":
                selector = "td.first-col";
                break;
            case "lastCol":
                selector = "td.last-col";
                break;
            case "band1Vert":
                selector = "td.odd-col";
                break;
            case "band2Vert":
                selector = "td.even-col";
                break;
            case "band1Horz":
                selector = "tr.odd-row";
                break;
            case "band2Horz":
                selector = "tr.even-row";
                break;
            default: return [];
        }
        xml.foreach(node, function (n) {
            switch (n.localName) {
                case "pPr":
                    result.push({
                        target: selector + " p",
                        values: _this.parseDefaultProperties(n, {})
                    });
                    break;
                case "rPr":
                    result.push({
                        target: selector + " span",
                        values: _this.parseDefaultProperties(n, {})
                    });
                    break;
                case "tblPr":
                case "tcPr":
                    result.push({
                        target: selector,
                        values: _this.parseDefaultProperties(n, {})
                    });
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseNumberingFile = function (xmlString) {
        var _this = this;
        var result = [];
        var xnums = xml.parse(xmlString, this.skipDeclaration);
        var mapping = {};
        var bullets = [];
        xml.foreach(xnums, function (n) {
            switch (n.localName) {
                case "abstractNum":
                    _this.parseAbstractNumbering(n, bullets)
                        .forEach(function (x) { return result.push(x); });
                    break;
                case "numPicBullet":
                    bullets.push(_this.parseNumberingPicBullet(n));
                    break;
                case "num":
                    var numId = xml.stringAttr(n, "numId");
                    var abstractNumId = xml.elementStringAttr(n, "abstractNumId", "val");
                    mapping[abstractNumId] = numId;
                    break;
            }
        });
        result.forEach(function (x) { return x.id = mapping[x.id]; });
        return result;
    };
    DocumentParser.prototype.parseNumberingPicBullet = function (elem) {
        var pict = xml.byTagName(elem, "pict");
        var shape = pict && xml.byTagName(pict, "shape");
        var imagedata = shape && xml.byTagName(shape, "imagedata");
        return imagedata ? {
            id: xml.intAttr(elem, "numPicBulletId"),
            src: xml.stringAttr(imagedata, "id"),
            style: xml.stringAttr(shape, "style")
        } : null;
    };
    DocumentParser.prototype.parseAbstractNumbering = function (node, bullets) {
        var _this = this;
        var result = [];
        var id = xml.stringAttr(node, "abstractNumId");
        xml.foreach(node, function (n) {
            switch (n.localName) {
                case "lvl":
                    result.push(_this.parseNumberingLevel(id, n, bullets));
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseNumberingLevel = function (id, node, bullets) {
        var _this = this;
        var result = {
            id: id,
            level: xml.intAttr(node, "ilvl"),
            style: {}
        };
        xml.foreach(node, function (n) {
            switch (n.localName) {
                case "pPr":
                    _this.parseDefaultProperties(n, result.style);
                    delete result.style["text-indent"];
                    break;
                case "lvlPicBulletId":
                    var id = xml.intAttr(n, "val");
                    result.bullet = bullets.filter(function (x) { return x.id == id; })[0];
                    break;
                case "lvlText":
                    result.levelText = xml.stringAttr(n, "val");
                    break;
                case "numFmt":
                    result.format = xml.stringAttr(n, "val");
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseParagraph = function (node) {
        var _this = this;
        var result = new paragraph_2.Paragraph();
        xml.foreach(node, function (c) {
            switch (c.localName) {
                case "r":
                    result.children.push(_this.parseRun(c));
                    break;
                case "hyperlink":
                    result.children.push(_this.parseHyperlink(c));
                    break;
                case "bookmarkStart":
                    result.children.push(xml_serialize_1.deserialize(node, new bookmark_1.Bookmark()));
                    break;
                case "pPr":
                    _this.parseParagraphProperties(c, result);
                    _this.parseCommonProperties(c, result.props);
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseParagraphProperties = function (elem, paragraph) {
        var _this = this;
        this.parseDefaultProperties(elem, paragraph.style = {}, null, function (c) {
            if (paragraph_1.parseParagraphProperties(c, paragraph.props))
                return true;
            switch (c.localName) {
                case "pStyle":
                    utils.addElementClass(paragraph, xml.className(c, "val"));
                    break;
                case "cnfStyle":
                    utils.addElementClass(paragraph, values.classNameOfCnfStyle(c));
                    break;
                case "framePr":
                    _this.parseFrame(c, paragraph);
                    break;
                case "rPr":
                    _this.parseDefaultProperties(c, paragraph.defaultRunStyle);
                    break;
                default:
                    return false;
            }
            return true;
        });
    };
    DocumentParser.prototype.parseFrame = function (node, paragraph) {
        var dropCap = xml.stringAttr(node, "dropCap");
        if (dropCap == "drop")
            paragraph.style["float"] = "left";
    };
    DocumentParser.prototype.parseHyperlink = function (node) {
        var _this = this;
        var result = xml_serialize_1.deserialize(node, new hyperlink_1.Hyperlink());
        xml.foreach(node, function (c) {
            switch (c.localName) {
                case "r":
                    result.children.push(_this.parseRun(c));
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseRun = function (node) {
        var _this = this;
        var result = new run_1.Run();
        xml.foreach(node, function (c) {
            switch (c.localName) {
                case "t":
                    result.children.push(xml_serialize_1.deserialize(c, new text_1.Text()));
                    break;
                case "fldChar":
                    result.fldCharType = xml.stringAttr(c, "fldCharType");
                    break;
                case "br":
                    result.children.push(xml_serialize_1.deserialize(c, new break_1.Break()));
                    break;
                case "lastRenderedPageBreak":
                    var br = new break_1.Break();
                    br.break = "page";
                    result.children.push(br);
                    break;
                case "sym":
                    result.children.push(xml_serialize_1.deserialize(c, new symbol_1.Symbol()));
                    break;
                case "tab":
                    result.children.push(new tab_1.Tab());
                    break;
                case "instrText":
                    result.instrText = c.textContent;
                    break;
                case "drawing":
                    var d = _this.parseDrawing(c);
                    if (d)
                        result.children = [d];
                    break;
                case "rPr":
                    _this.parseRunProperties(c, result);
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseRunProperties = function (elem, run) {
        this.parseDefaultProperties(elem, run.style = {}, null, function (c) {
            switch (c.localName) {
                case "rStyle":
                    run.className = xml.className(c, "val");
                    break;
                case "vertAlign":
                    run.props.verticalAlignment = xml.stringAttr(c, "val");
                    break;
                default:
                    return false;
            }
            return true;
        });
    };
    DocumentParser.prototype.parseDrawing = function (node) {
        for (var _i = 0, _a = xml.elements(node); _i < _a.length; _i++) {
            var n = _a[_i];
            switch (n.localName) {
                case "inline":
                case "anchor":
                    return this.parseDrawingWrapper(n);
            }
        }
    };
    DocumentParser.prototype.parseDrawingWrapper = function (node) {
        var result = new drawing_1.Drawing();
        var isAnchor = node.localName == "anchor";
        var wrapType = null;
        var simplePos = xml.boolAttr(node, "simplePos");
        var posX = { relative: "page", align: null, offset: "0" };
        var posY = { relative: "page", align: null, offset: "0" };
        for (var _i = 0, _a = xml.elements(node); _i < _a.length; _i++) {
            var n = _a[_i];
            switch (n.localName) {
                case "simplePos":
                    if (simplePos) {
                        posX.offset = xml.sizeAttr(n, "x", SizeType.Emu);
                        posY.offset = xml.sizeAttr(n, "y", SizeType.Emu);
                    }
                    break;
                case "extent":
                    result.style["width"] = xml.sizeAttr(n, "cx", SizeType.Emu);
                    result.style["height"] = xml.sizeAttr(n, "cy", SizeType.Emu);
                    break;
                case "positionH":
                case "positionV":
                    if (!simplePos) {
                        var pos = n.localName == "positionH" ? posX : posY;
                        var relativeFrom = xml.stringAttr(n, "relativeFrom");
                        var alignNode = xml.byTagName(n, "align");
                        var offsetNode = xml.byTagName(n, "posOffset");
                        if (relativeFrom)
                            pos.relative = relativeFrom;
                        if (alignNode)
                            pos.align = alignNode.textContent;
                        if (offsetNode)
                            pos.offset = xml.sizeValue(offsetNode, SizeType.Emu);
                    }
                    break;
                case "wrapTopAndBottom":
                    wrapType = "wrapTopAndBottom";
                    break;
                case "wrapNone":
                    wrapType = "wrapNone";
                    break;
                case "graphic":
                    var g = this.parseGraphic(n);
                    if (g)
                        result.children.push(g);
                    break;
            }
        }
        if (wrapType == "wrapTopAndBottom") {
            result.style['display'] = 'block';
            if (posX.align) {
                result.style['text-align'] = posX.align;
                result.style['width'] = "100%";
            }
        }
        else if (wrapType == "wrapNone") {
            result.style['display'] = 'block';
            result.style['position'] = 'relative';
            result.style["width"] = "0px";
            result.style["height"] = "0px";
            if (posX.offset)
                result.style["left"] = posX.offset;
            if (posY.offset)
                result.style["top"] = posY.offset;
        }
        else if (isAnchor && !simplePos && !(posX.align == 'left' || posX.align == 'right')) {
            result.style['position'] = 'absolute';
            result.style['top'] = posY.offset;
            result.style['left'] = posX.offset;
        }
        else if (isAnchor && (posX.align == 'left' || posX.align == 'right')) {
            result.style["float"] = posX.align;
            result.style["margin-left"] = '7pt';
            result.style["margin-right"] = '7pt';
        }
        result.posX = posX;
        result.posY = posY;
        return result;
    };
    DocumentParser.prototype.parseGraphic = function (elem) {
        var graphicData = xml.byTagName(elem, "graphicData");
        for (var _i = 0, _a = xml.elements(graphicData); _i < _a.length; _i++) {
            var n = _a[_i];
            switch (n.localName) {
                case "pic":
                    return this.parsePicture(n);
            }
        }
        return null;
    };
    DocumentParser.prototype.parsePicture = function (elem) {
        var result = new image_1.Image();
        var blipFill = xml.byTagName(elem, "blipFill");
        var blip = xml.byTagName(blipFill, "blip");
        var srcRect = xml.byTagName(blipFill, "srcRect");
        var stretch = xml.byTagName(blipFill, "stretch");
        result.src = xml.stringAttr(blip, "embed");
        if (srcRect) {
            result.crop = {
                left: xml.intAttr(srcRect, "l"),
                right: xml.intAttr(srcRect, "r"),
                top: xml.intAttr(srcRect, "t"),
                bottom: xml.intAttr(srcRect, "b")
            };
        }
        if (stretch) {
            var fillRect = xml.byTagName(stretch, "fillRect");
            if (fillRect) {
                result.stretch = {
                    left: xml.intAttr(fillRect, "l"),
                    right: xml.intAttr(fillRect, "r"),
                    top: xml.intAttr(fillRect, "t"),
                    bottom: xml.intAttr(fillRect, "b")
                };
            }
        }
        var spPr = xml.byTagName(elem, "spPr");
        var xfrm = xml.byTagName(spPr, "xfrm");
        result.style["position"] = "relative";
        for (var _i = 0, _a = xml.elements(xfrm); _i < _a.length; _i++) {
            var n = _a[_i];
            switch (n.localName) {
                case "ext":
                    result.style["width"] = xml.sizeAttr(n, "cx", SizeType.Emu);
                    result.style["height"] = xml.sizeAttr(n, "cy", SizeType.Emu);
                    break;
                case "off":
                    result.style["left"] = xml.sizeAttr(n, "x", SizeType.Emu);
                    result.style["top"] = xml.sizeAttr(n, "y", SizeType.Emu);
                    break;
            }
        }
        return result;
    };
    DocumentParser.prototype.parseTable = function (node) {
        var _this = this;
        var result = new table_1.Table();
        xml.foreach(node, function (c) {
            switch (c.localName) {
                case "tr":
                    result.children.push(_this.parseTableRow(c));
                    break;
                case "tblGrid":
                    result.columns = _this.parseTableColumns(c);
                    break;
                case "tblPr":
                    _this.parseTableProperties(c, result);
                    break;
            }
        });
        result.children.forEach(function (row, rowIndex) {
            if (row instanceof row_1.Row) {
                row.children.forEach(function (cell, cellIndex) {
                    if (cell instanceof cell_1.Cell && cell.props.vMerge === 'restart') {
                        var rowSpan = 1;
                        for (var i = rowIndex + 1; i < result.children.length; i++) {
                            var nextRow = result.children[i];
                            if (nextRow instanceof row_1.Row) {
                                var nextCell = nextRow.children[cellIndex];
                                if (nextCell instanceof cell_1.Cell && (!nextCell.props.vMerge || nextCell.props.vMerge === 'restart')) {
                                    break;
                                }
                                rowSpan++;
                            }
                        }
                        cell.props.rowSpan = rowSpan;
                    }
                });
            }
        });
        return result;
    };
    DocumentParser.prototype.parseTableColumns = function (node) {
        var result = [];
        xml.foreach(node, function (n) {
            switch (n.localName) {
                case "gridCol":
                    result.push({ width: common_2.lengthAttr(n, common_1.ns.wordml, "w") });
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseTableProperties = function (elem, table) {
        var _this = this;
        table.style = {};
        table.cellStyle = {};
        this.parseDefaultProperties(elem, table.style, table.cellStyle, function (c) {
            switch (c.localName) {
                case "tblStyle":
                    table.className = xml.className(c, "val");
                    break;
                case "tblLook":
                    utils.addElementClass(table, values.classNameOftblLook(c));
                    break;
                case "tblpPr":
                    _this.parseTablePosition(c, table);
                    break;
                default:
                    return false;
            }
            return true;
        });
        switch (table.style["text-align"]) {
            case "center":
                delete table.style["text-align"];
                table.style["margin-left"] = "auto";
                table.style["margin-right"] = "auto";
                break;
            case "right":
                delete table.style["text-align"];
                table.style["margin-left"] = "auto";
                break;
        }
    };
    DocumentParser.prototype.parseTablePosition = function (node, table) {
        var topFromText = xml.sizeAttr(node, "topFromText");
        var bottomFromText = xml.sizeAttr(node, "bottomFromText");
        var rightFromText = xml.sizeAttr(node, "rightFromText");
        var leftFromText = xml.sizeAttr(node, "leftFromText");
        table.style["float"] = 'left';
        table.style["margin-bottom"] = values.addSize(table.style["margin-bottom"], bottomFromText);
        table.style["margin-left"] = values.addSize(table.style["margin-left"], leftFromText);
        table.style["margin-right"] = values.addSize(table.style["margin-right"], rightFromText);
        table.style["margin-top"] = values.addSize(table.style["margin-top"], topFromText);
    };
    DocumentParser.prototype.parseTableRow = function (node) {
        var _this = this;
        var result = new row_1.Row();
        xml.foreach(node, function (c) {
            switch (c.localName) {
                case "tc":
                    result.children.push(_this.parseTableCell(c));
                    break;
                case "trPr":
                    _this.parseTableRowProperties(c, result);
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseTableRowProperties = function (elem, row) {
        row.style = this.parseDefaultProperties(elem, {}, null, function (c) {
            switch (c.localName) {
                case "cnfStyle":
                    row.className = values.classNameOfCnfStyle(c);
                    break;
                default:
                    return false;
            }
            return true;
        });
    };
    DocumentParser.prototype.parseTableCell = function (node) {
        var _this = this;
        var result = new cell_1.Cell();
        xml.foreach(node, function (c) {
            switch (c.localName) {
                case "tbl":
                    result.children.push(_this.parseTable(c));
                    break;
                case "p":
                    result.children.push(_this.parseParagraph(c));
                    _this.checkAndMergeConsecutivePragraphBorder(result.children);
                    break;
                case "tcPr":
                    _this.parseTableCellProperties(c, result);
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseTableCellProperties = function (elem, cell) {
        cell.style = this.parseDefaultProperties(elem, {}, cell.childrenStyle, function (c) {
            switch (c.localName) {
                case "gridSpan":
                    cell.props.gridSpan = xml.intAttr(c, "val", null);
                    break;
                case "vMerge":
                    cell.props.vMerge = xml.stringAttr(c, "val");
                    if (!cell.props.vMerge)
                        cell.props.vMerge = 'continue';
                    break;
                case "cnfStyle":
                    cell.className = values.classNameOfCnfStyle(c);
                    break;
                default:
                    return false;
            }
            return true;
        });
    };
    DocumentParser.prototype.parseDefaultProperties = function (elem, style, childStyle, handler) {
        var _this = this;
        if (style === void 0) { style = null; }
        if (childStyle === void 0) { childStyle = null; }
        if (handler === void 0) { handler = null; }
        style = style || {};
        xml.foreach(elem, function (c) {
            switch (c.localName) {
                case "jc":
                    style["text-align"] = values.valueOfJc(c);
                    break;
                case "textAlignment":
                    style["vertical-align"] = values.valueOfTextAlignment(c);
                    break;
                case "textDirection":
                    var textDirection = xml.stringAttr(c, "val");
                    switch (textDirection) {
                        case "btLr":
                            if (childStyle) {
                                childStyle["writing-mode"] = "vertical-rl";
                                childStyle["transform"] = "rotate(180deg)";
                                childStyle["width"] = "max-content";
                                childStyle["height"] = "max-content";
                            }
                            else {
                                style["writing-mode"] = "vertical-rl";
                                style["transform"] = "rotate(180deg)";
                            }
                            break;
                        case "lrTb":
                        case "lrTbV":
                        case "tbLrV":
                        case "tbRl":
                        case "tbRlV":
                            break;
                    }
                    break;
                case "color":
                    style["color"] = xml.colorAttr(c, "val", null, exports.autos.color);
                    break;
                case "sz":
                    style["font-size"] = style["min-height"] = xml.sizeAttr(c, "val", SizeType.FontSize);
                    break;
                case "shd":
                    style["background-color"] = xml.colorAttr(c, "fill", null, exports.autos.shd);
                    break;
                case "highlight":
                    style["background-color"] = xml.colorAttr(c, "val", null, exports.autos.highlight);
                    break;
                case "tcW":
                    if (_this.ignoreWidth)
                        break;
                case "tblW":
                    style["width"] = values.valueOfSize(c, "w");
                    break;
                case "trHeight":
                    _this.parseTrHeight(c, style);
                    break;
                case "strike":
                    style["text-decoration"] = values.valueOfStrike(c);
                    break;
                case "b":
                    style["font-weight"] = values.valueOfBold(c);
                    break;
                case "caps":
                    style["text-transform"] = "uppercase";
                    break;
                case "i":
                    style["font-style"] = "italic";
                    break;
                case "u":
                    _this.parseUnderline(c, style);
                    break;
                case "ind":
                    _this.parseIndentation(c, style, 'padding');
                    break;
                case "tblInd":
                    _this.parseIndentation(c, style, 'margin');
                    break;
                case "rFonts":
                    _this.parseFont(c, style);
                    break;
                case "tblBorders":
                    _this.parseBorderProperties(c, childStyle || style);
                    break;
                case "tblCellSpacing":
                    style["border-spacing"] = values.valueOfMargin(c);
                    style["border-collapse"] = "separate";
                    break;
                case "pBdr":
                    _this.parseBorderProperties(c, style);
                    break;
                case "bdr":
                    style["border"] = values.valueOfBorder(c);
                    break;
                case "tcBorders":
                    _this.parseBorderProperties(c, style);
                    break;
                case "noWrap":
                    break;
                case "tblCellMar":
                case "tcMar":
                    _this.parseMarginProperties(c, childStyle || style);
                    break;
                case "tblLayout":
                    style["table-layout"] = values.valueOfTblLayout(c);
                    break;
                case "vAlign":
                    style["vertical-align"] = values.valueOfTextAlignment(c);
                    break;
                case "spacing":
                    if (elem.localName == "pPr")
                        _this.parseSpacing(c, style);
                    break;
                case "lang":
                case "noProof":
                case "webHidden":
                    break;
                default:
                    if (handler != null && !handler(c))
                        _this.debug && console.warn("DOCX: Unknown document element: " + c.localName);
                    break;
            }
        });
        return style;
    };
    DocumentParser.prototype.parseUnderline = function (node, style) {
        var val = xml.stringAttr(node, "val");
        if (val == null || val == "none")
            return;
        switch (val) {
            case "dash":
            case "dashDotDotHeavy":
            case "dashDotHeavy":
            case "dashedHeavy":
            case "dashLong":
            case "dashLongHeavy":
            case "dotDash":
            case "dotDotDash":
                style["text-decoration-style"] = "dashed";
                break;
            case "dotted":
            case "dottedHeavy":
                style["text-decoration-style"] = "dotted";
                break;
            case "double":
                style["text-decoration-style"] = "double";
                break;
            case "single":
            case "thick":
                style["text-decoration"] = "underline";
                break;
            case "wave":
            case "wavyDouble":
            case "wavyHeavy":
                style["text-decoration-style"] = "wavy";
                break;
            case "words":
                style["text-decoration"] = "underline";
                break;
        }
        var col = xml.colorAttr(node, "color");
        if (col)
            style["text-decoration-color"] = col;
    };
    DocumentParser.prototype.parseFont = function (node, style) {
        var ascii = xml.stringAttr(node, "ascii");
        if (ascii)
            style["font-family"] = ascii;
    };
    DocumentParser.prototype.parseIndentation = function (node, style, marginOrPadding) {
        var firstLine = xml.sizeAttr(node, "firstLine");
        var hanging = xml.sizeAttr(node, "hanging");
        var left = xml.sizeAttr(node, "left");
        var start = xml.sizeAttr(node, "start");
        var right = xml.sizeAttr(node, "right");
        var end = xml.sizeAttr(node, "end");
        var numericFirstLine = firstLine ? parseInt(firstLine.replace('pt', '')) : 0;
        var numericHanging = hanging ? parseInt(hanging.replace('pt', '')) : 0;
        var numericTextIndent = numericFirstLine - numericHanging;
        if (numericTextIndent !== 0) {
            var textIndent = numericTextIndent + "pt";
            style["text-indent"] = textIndent;
        }
        if (left || start)
            style[marginOrPadding + "-left"] = left || start;
        if (right || end)
            style[marginOrPadding + "-right"] = right || end;
    };
    DocumentParser.prototype.parseSpacing = function (node, style) {
        var before = xml.sizeAttr(node, "before");
        var after = xml.sizeAttr(node, "after");
        var line = xml.intAttr(node, "line", null);
        var lineRule = xml.stringAttr(node, "lineRule");
        style["margin-top"] = '2pt';
        if (before && before !== '0.00pt')
            style["margin-top"] = before;
        if (after)
            style["margin-bottom"] = after;
        if (line !== null) {
            switch (lineRule) {
                case "auto":
                    style["line-height"] = "" + (line / 240).toFixed(2);
                    break;
                case "atLeast":
                    style["line-height"] = "calc(100% + " + line / 20 + "pt)";
                    break;
                default:
                    style["line-height"] = style["min-height"] = line / 20 + "pt";
                    break;
            }
        }
    };
    DocumentParser.prototype.parseMarginProperties = function (node, output) {
        xml.foreach(node, function (c) {
            switch (c.localName) {
                case "left":
                    output["padding-left"] = values.valueOfMargin(c);
                    break;
                case "right":
                    output["padding-right"] = values.valueOfMargin(c);
                    break;
                case "top":
                    output["padding-top"] = values.valueOfMargin(c);
                    break;
                case "bottom":
                    output["padding-bottom"] = values.valueOfMargin(c);
                    break;
            }
        });
    };
    DocumentParser.prototype.parseTrHeight = function (node, output) {
        switch (xml.stringAttr(node, "hRule")) {
            case "exact":
                output["height"] = xml.sizeAttr(node, "val");
                break;
            case "atLeast":
            default:
                output["height"] = xml.sizeAttr(node, "val");
                break;
        }
    };
    DocumentParser.prototype.parseBorderProperties = function (node, output) {
        xml.foreach(node, function (c) {
            switch (c.localName) {
                case "start":
                case "left":
                    output["border-left"] = values.valueOfBorder(c);
                    output["bdr-left"] = values.valueOfBorder(c);
                    break;
                case "end":
                case "right":
                    output["border-right"] = values.valueOfBorder(c);
                    output["bdr-right"] = values.valueOfBorder(c);
                    break;
                case "top":
                    output["border-top"] = values.valueOfBorder(c);
                    output["bdr-top"] = values.valueOfBorder(c);
                    break;
                case "bottom":
                    output["border-bottom"] = values.valueOfBorder(c);
                    output["bdr-bottom"] = values.valueOfBorder(c);
                    break;
                case "between":
                    output["bdr-between"] = values.valueOfBorder(c);
                    break;
            }
        });
    };
    return DocumentParser;
}());
exports.DocumentParser = DocumentParser;
var SizeType;
(function (SizeType) {
    SizeType[SizeType["FontSize"] = 0] = "FontSize";
    SizeType[SizeType["Dxa"] = 1] = "Dxa";
    SizeType[SizeType["Emu"] = 2] = "Emu";
    SizeType[SizeType["Border"] = 3] = "Border";
    SizeType[SizeType["Percent"] = 4] = "Percent";
})(SizeType || (SizeType = {}));
var xml = (function () {
    function xml() {
    }
    xml.parse = function (xmlString, skipDeclaration) {
        if (skipDeclaration === void 0) { skipDeclaration = true; }
        if (skipDeclaration)
            xmlString = xmlString.replace(/<[?].*[?]>/, "");
        return new DOMParser().parseFromString(xmlString, "application/xml").firstChild;
    };
    xml.elements = function (node, tagName) {
        if (tagName === void 0) { tagName = null; }
        var result = [];
        for (var i = 0; i < node.childNodes.length; i++) {
            var n = node.childNodes[i];
            if (n.nodeType == 1 && (tagName == null || n.localName == tagName))
                result.push(n);
        }
        return result;
    };
    xml.foreach = function (node, cb) {
        for (var i = 0; i < node.childNodes.length; i++) {
            var n = node.childNodes[i];
            if (n.nodeType == 1)
                cb(n);
        }
    };
    xml.byTagName = function (elem, tagName) {
        for (var i = 0; i < elem.childNodes.length; i++) {
            var n = elem.childNodes[i];
            if (n.nodeType == 1 && n.localName == tagName)
                return elem.childNodes[i];
        }
        return null;
    };
    xml.elementStringAttr = function (elem, nodeName, attrName) {
        var n = xml.byTagName(elem, nodeName);
        return n ? xml.stringAttr(n, attrName) : null;
    };
    xml.stringAttr = function (node, attrName) {
        var elem = node;
        for (var i = 0; i < elem.attributes.length; i++) {
            var attr = elem.attributes.item(i);
            if (attr.localName == attrName)
                return attr.value;
        }
        return null;
    };
    xml.colorAttr = function (node, attrName, defValue, autoColor) {
        if (defValue === void 0) { defValue = null; }
        if (autoColor === void 0) { autoColor = 'black'; }
        var v = xml.stringAttr(node, attrName);
        switch (v) {
            case "yellow":
                return v;
            case "auto":
                return autoColor;
        }
        return v ? "#" + v : defValue;
    };
    xml.boolAttr = function (node, attrName, defValue) {
        if (defValue === void 0) { defValue = false; }
        var v = xml.stringAttr(node, attrName);
        switch (v) {
            case "1": return true;
            case "0": return false;
        }
        return defValue;
    };
    xml.intAttr = function (node, attrName, defValue) {
        if (defValue === void 0) { defValue = 0; }
        var val = xml.stringAttr(node, attrName);
        return val ? parseInt(xml.stringAttr(node, attrName)) : defValue;
    };
    xml.sizeAttr = function (node, attrName, type) {
        if (type === void 0) { type = SizeType.Dxa; }
        return xml.convertSize(xml.stringAttr(node, attrName), type);
    };
    xml.sizeValue = function (node, type) {
        if (type === void 0) { type = SizeType.Dxa; }
        return xml.convertSize(node.textContent, type);
    };
    xml.convertSize = function (val, type) {
        if (type === void 0) { type = SizeType.Dxa; }
        if (val == null || val.indexOf("pt") > -1)
            return val;
        var intVal = parseInt(val);
        switch (type) {
            case SizeType.Dxa: return (0.05 * intVal).toFixed(2) + "pt";
            case SizeType.Emu: return (intVal / 12700).toFixed(2) + "pt";
            case SizeType.FontSize: return (0.5 * intVal).toFixed(2) + "pt";
            case SizeType.Border: return (0.125 * intVal).toFixed(2) + "pt";
            case SizeType.Percent: return (0.02 * intVal).toFixed(2) + "%";
        }
        return val;
    };
    xml.className = function (node, attrName) {
        var val = xml.stringAttr(node, attrName);
        return val && val.replace(/[ .]+/g, '-').replace(/[&]+/g, 'and');
    };
    return xml;
}());
var values = (function () {
    function values() {
    }
    values.valueOfBold = function (c) {
        return xml.boolAttr(c, "val", true) ? "bold" : "normal";
    };
    values.valueOfSize = function (c, attr) {
        var type = SizeType.Dxa;
        switch (xml.stringAttr(c, "type")) {
            case "dxa": break;
            case "pct":
                type = SizeType.Percent;
                break;
        }
        return xml.sizeAttr(c, attr, type);
    };
    values.valueOfStrike = function (c) {
        return xml.boolAttr(c, "val", true) ? "line-through" : "none";
    };
    values.valueOfMargin = function (c) {
        return xml.sizeAttr(c, "w");
    };
    values.valueOfRelType = function (c) {
        switch (xml.sizeAttr(c, "Type")) {
            case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings":
                return dom_1.DomRelationshipType.Settings;
            case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme":
                return dom_1.DomRelationshipType.Theme;
            case "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects":
                return dom_1.DomRelationshipType.StylesWithEffects;
            case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles":
                return dom_1.DomRelationshipType.Styles;
            case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable":
                return dom_1.DomRelationshipType.FontTable;
            case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image":
                return dom_1.DomRelationshipType.Image;
            case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings":
                return dom_1.DomRelationshipType.WebSettings;
        }
        return dom_1.DomRelationshipType.Unknown;
    };
    values.valueOfBorder = function (c) {
        var type = xml.stringAttr(c, "val");
        if (type == "nil")
            return "none";
        var color = xml.colorAttr(c, "color");
        var size = xml.sizeAttr(c, "sz", SizeType.Border);
        return size + " solid " + (color == "auto" ? "black" : color);
    };
    values.valueOfTblLayout = function (c) {
        var type = xml.stringAttr(c, "val");
        return type == "fixed" ? "fixed" : "auto";
    };
    values.classNameOfCnfStyle = function (c) {
        var className = "";
        var val = xml.stringAttr(c, "val");
        if (val[0] == "1")
            className += " first-row";
        if (val[1] == "1")
            className += " last-row";
        if (val[2] == "1")
            className += " first-col";
        if (val[3] == "1")
            className += " last-col";
        if (val[4] == "1")
            className += " odd-col";
        if (val[5] == "1")
            className += " even-col";
        if (val[6] == "1")
            className += " odd-row";
        if (val[7] == "1")
            className += " even-row";
        if (val[8] == "1")
            className += " ne-cell";
        if (val[9] == "1")
            className += " nw-cell";
        if (val[10] == "1")
            className += " se-cell";
        if (val[11] == "1")
            className += " sw-cell";
        return className.trim();
    };
    values.valueOfJc = function (c) {
        var type = xml.stringAttr(c, "val");
        switch (type) {
            case "start":
            case "left": return "left";
            case "center": return "center";
            case "end":
            case "right": return "right";
            case "both": return "justify";
        }
        return type;
    };
    values.valueOfTextAlignment = function (c) {
        var type = xml.stringAttr(c, "val");
        switch (type) {
            case "auto":
            case "baseline": return "baseline";
            case "top": return "top";
            case "center": return "middle";
            case "bottom": return "bottom";
        }
        return type;
    };
    values.addSize = function (a, b) {
        if (a == null)
            return b;
        if (b == null)
            return a;
        return "calc(" + a + " + " + b + ")";
    };
    values.checkMask = function (num, mask) {
        return (num & mask) == mask;
    };
    values.classNameOftblLook = function (c) {
        var className = "";
        if (xml.boolAttr(c, "firstColumn"))
            className += " first-col";
        if (xml.boolAttr(c, "firstRow"))
            className += " first-row";
        if (xml.boolAttr(c, "lastColumn"))
            className += " lat-col";
        if (xml.boolAttr(c, "lastRow"))
            className += " last-row";
        if (xml.boolAttr(c, "noHBand"))
            className += " no-hband";
        if (xml.boolAttr(c, "noVBand"))
            className += " no-vband";
        return className.trim();
    };
    return values;
}());


/***/ }),

/***/ "./src/document.ts":
/*!*************************!*\
  !*** ./src/document.ts ***!
  \*************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var JSZip = __webpack_require__(/*! jszip */ "jszip");
var PartType;
(function (PartType) {
    PartType["Document"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    PartType["Numbering"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
    PartType["FontTable"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";
    PartType["Style"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
    PartType["DocumentRelations"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.relationships+xml";
    PartType["NumberingRelations"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering.relationships+xml";
    PartType["FontRelations"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable.relationships+xml";
})(PartType = exports.PartType || (exports.PartType = {}));
var normalizePath = function (path) {
    return path ? path.replace(/^\/+/, '') : path;
};
var Document = (function () {
    function Document() {
        this.zip = new JSZip();
        this.docRelations = null;
        this.fontRelations = null;
        this.numRelations = null;
        this.styles = null;
        this.fonts = null;
        this.numbering = null;
        this.document = null;
        this.renderContextRelations = null;
    }
    Document.load = function (blob, parser) {
        var d = new Document();
        d.parser = parser;
        return d.zip.loadAsync(blob).then(function (z) {
            return d.loadContentType().then(function () { return d; });
        });
    };
    Document.prototype.loadDocumentImage = function (id) {
        return this.loadResource(this.renderContextRelations, id, "blob")
            .then(function (x) { return x ? URL.createObjectURL(x) : null; });
    };
    Document.prototype.loadNumberingImage = function (id) {
        return this.loadResource(this.numRelations, id, "blob")
            .then(function (x) { return x ? URL.createObjectURL(x) : null; });
    };
    Document.prototype.loadFont = function (id, key) {
        return this.loadResource(this.fontRelations, id, "uint8array")
            .then(function (x) { return x ? URL.createObjectURL(new Blob([deobfuscate(x, key)])) : x; });
    };
    Document.prototype.loadAndSetRenderContextToHeaderOrFooter = function (id) {
        return __awaiter(this, void 0, void 0, function () {
            var headerOrFooterPath, headerOrFooterRelationsPath, relationsFile, relationsXml;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        headerOrFooterPath = this.getResoucePath(this.docRelations, id);
                        headerOrFooterRelationsPath = normalizePath(this.getRelPath(headerOrFooterPath));
                        relationsFile = this.zip.files[headerOrFooterRelationsPath];
                        if (!relationsFile) return [3, 2];
                        return [4, relationsFile.async("text")];
                    case 1:
                        relationsXml = _a.sent();
                        this.renderContextRelations = this.parser.parseDocumentRelationsFile(relationsXml);
                        _a.label = 2;
                    case 2: return [2, this.loadResource(this.docRelations, id, "text")
                            .then(function (resource) { return resource ? _this.parser.parseHeaderOrFooter(resource) : null; })];
                }
            });
        });
    };
    Document.prototype.setRenderContextToMainDocument = function () {
        this.renderContextRelations = this.docRelations;
    };
    Document.prototype.getHyperlinkTarget = function (id) {
        var rel = this.renderContextRelations.find(function (x) { return x.id == id; });
        return rel.target;
    };
    Document.prototype.getRelPath = function (path) {
        if (!path)
            return path;
        var beginning = path.substr(0, path.lastIndexOf("/") + 1);
        var remaining = path.replace(beginning, "");
        return beginning + "_rels/" + remaining + ".rels";
    };
    Document.prototype.loadContentType = function () {
        var _this = this;
        var contentTypePart = this.zip.files['[Content_Types].xml'];
        if (!contentTypePart) {
            throw new Error("Invalid office open xml document, missing [Content_Types].xml");
        }
        return contentTypePart.async("text").then(function (xml) {
            var parts = _this.parser.parseContentTypeFile(xml);
            var files = [
                _this.loadPart(PartType.DocumentRelations, normalizePath(_this.getRelPath(parts.get(PartType.Document)))),
                _this.loadPart(PartType.FontRelations, normalizePath(_this.getRelPath(parts.get(PartType.FontTable)))),
                _this.loadPart(PartType.NumberingRelations, normalizePath(_this.getRelPath(parts.get(PartType.Numbering)))),
                _this.loadPart(PartType.Style, normalizePath(parts.get(PartType.Style))),
                _this.loadPart(PartType.FontTable, normalizePath(parts.get(PartType.FontTable))),
                _this.loadPart(PartType.Numbering, normalizePath(parts.get(PartType.Numbering))),
                _this.loadPart(PartType.Document, normalizePath(parts.get(PartType.Document)))
            ];
            return Promise.all(files.filter(function (x) { return x != null; }));
        });
    };
    Document.prototype.resolvePath = function (path) {
        return normalizePath(path.replace(/([^/]+)\/\.\./g, ""));
    };
    Document.prototype.getResoucePath = function (relations, id) {
        var rel = relations.find(function (x) { return x.id == id; });
        return rel ? rel.target.startsWith("/") ? normalizePath(rel.target) : this.resolvePath("word/" + rel.target) : null;
    };
    Document.prototype.loadResource = function (relations, id, outputType) {
        if (outputType === void 0) { outputType = "base64"; }
        var path = this.getResoucePath(relations, id);
        return path ? this.zip.files[path].async(outputType) : Promise.resolve(null);
    };
    Document.prototype.loadPart = function (part, partPath) {
        var _this = this;
        var f = this.zip.files[partPath];
        return f ? f.async("text").then(function (xml) {
            switch (part) {
                case PartType.FontRelations:
                    _this.fontRelations = _this.parser.parseDocumentRelationsFile(xml);
                    break;
                case PartType.DocumentRelations:
                    _this.docRelations = _this.parser.parseDocumentRelationsFile(xml);
                    break;
                case PartType.NumberingRelations:
                    _this.numRelations = _this.parser.parseDocumentRelationsFile(xml);
                    break;
                case PartType.Style:
                    _this.styles = _this.parser.parseStylesFile(xml);
                    break;
                case PartType.Numbering:
                    _this.numbering = _this.parser.parseNumberingFile(xml);
                    break;
                case PartType.Document:
                    _this.document = _this.parser.parseDocumentFile(xml);
                    break;
                case PartType.FontTable:
                    _this.fontTable = _this.parser.parseFontTable(xml);
                    break;
            }
            return _this;
        }) : null;
    };
    return Document;
}());
exports.Document = Document;
function deobfuscate(data, guidKey) {
    var len = 16;
    var trimmed = guidKey.replace(/{|}|-/g, "");
    var numbers = new Array(len);
    for (var i = 0; i < len; i++)
        numbers[len - i - 1] = parseInt(trimmed.substr(i * 2, 2), 16);
    for (var i = 0; i < 32; i++)
        data[i] = data[i] ^ numbers[i % len];
    return data;
}
exports.deobfuscate = deobfuscate;


/***/ }),

/***/ "./src/docx-preview.ts":
/*!*****************************!*\
  !*** ./src/docx-preview.ts ***!
  \*****************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
Object.defineProperty(exports, "__esModule", { value: true });
var document_1 = __webpack_require__(/*! ./document */ "./src/document.ts");
var document_parser_1 = __webpack_require__(/*! ./document-parser */ "./src/document-parser.ts");
var html_renderer_1 = __webpack_require__(/*! ./html-renderer */ "./src/html-renderer.ts");
function renderAsync(data, bodyContainer, styleContainer, userOptions) {
    if (styleContainer === void 0) { styleContainer = null; }
    if (userOptions === void 0) { userOptions = null; }
    var parser = new document_parser_1.DocumentParser();
    var renderer = new html_renderer_1.HtmlRenderer(window.document);
    var options = __assign({ ignoreHeight: false, ignoreWidth: false, ignoreFonts: false, breakPages: true, debug: false, experimental: false, className: "docx", inWrapper: true }, userOptions);
    parser.ignoreWidth = options.ignoreWidth;
    parser.debug = options.debug || parser.debug;
    renderer.className = options.className || "docx";
    renderer.inWrapper = options.inWrapper;
    return document_1.Document.load(data, parser).then(function (doc) {
        renderer.render(doc, bodyContainer, styleContainer, options);
        return doc;
    });
}
exports.renderAsync = renderAsync;


/***/ }),

/***/ "./src/dom/common.ts":
/*!***************************!*\
  !*** ./src/dom/common.ts ***!
  \***************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
exports.ns = {
    wordml: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    relationships: "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
};
function renderLength(l) {
    return !l ? null : "" + l.value + l.type;
}
exports.renderLength = renderLength;


/***/ }),

/***/ "./src/dom/dom.ts":
/*!************************!*\
  !*** ./src/dom/dom.ts ***!
  \************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var DomType;
(function (DomType) {
    DomType["Document"] = "document";
    DomType["HeaderOrFooter"] = "headerOrFooter";
})(DomType = exports.DomType || (exports.DomType = {}));
var DomRelationshipType;
(function (DomRelationshipType) {
    DomRelationshipType[DomRelationshipType["Settings"] = 0] = "Settings";
    DomRelationshipType[DomRelationshipType["Theme"] = 1] = "Theme";
    DomRelationshipType[DomRelationshipType["StylesWithEffects"] = 2] = "StylesWithEffects";
    DomRelationshipType[DomRelationshipType["Styles"] = 3] = "Styles";
    DomRelationshipType[DomRelationshipType["FontTable"] = 4] = "FontTable";
    DomRelationshipType[DomRelationshipType["Image"] = 5] = "Image";
    DomRelationshipType[DomRelationshipType["WebSettings"] = 6] = "WebSettings";
    DomRelationshipType[DomRelationshipType["Unknown"] = 7] = "Unknown";
})(DomRelationshipType = exports.DomRelationshipType || (exports.DomRelationshipType = {}));


/***/ }),

/***/ "./src/dom/render-context.ts":
/*!***********************************!*\
  !*** ./src/dom/render-context.ts ***!
  \***********************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var RenderContext = (function () {
    function RenderContext() {
        this.currentPageNumber = 1;
    }
    RenderContext.prototype.numberingClass = function (id, lvl) {
        return this.className + "-num-" + id + "-" + lvl;
    };
    return RenderContext;
}());
exports.RenderContext = RenderContext;


/***/ }),

/***/ "./src/elements/bookmark.ts":
/*!**********************************!*\
  !*** ./src/elements/bookmark.ts ***!
  \**********************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
Object.defineProperty(exports, "__esModule", { value: true });
var element_base_1 = __webpack_require__(/*! ./element-base */ "./src/elements/element-base.ts");
var xml_serialize_1 = __webpack_require__(/*! ../parser/xml-serialize */ "./src/parser/xml-serialize.ts");
var Bookmark = (function (_super) {
    __extends(Bookmark, _super);
    function Bookmark() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Bookmark.prototype.render = function (ctx) {
        var elem = ctx.html.createElement("span");
        elem.id = this.name;
        return elem;
    };
    __decorate([
        xml_serialize_1.fromAttribute("name")
    ], Bookmark.prototype, "name", void 0);
    return Bookmark;
}(element_base_1.ElementBase));
exports.Bookmark = Bookmark;


/***/ }),

/***/ "./src/elements/break.ts":
/*!*******************************!*\
  !*** ./src/elements/break.ts ***!
  \*******************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
Object.defineProperty(exports, "__esModule", { value: true });
var element_base_1 = __webpack_require__(/*! ./element-base */ "./src/elements/element-base.ts");
var xml_serialize_1 = __webpack_require__(/*! ../parser/xml-serialize */ "./src/parser/xml-serialize.ts");
var Break = (function (_super) {
    __extends(Break, _super);
    function Break() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.break = "textWrapping";
        return _this;
    }
    Break.prototype.render = function (ctx) {
        if (this.break === 'page') {
            var pageBreak = ctx.html.createElement('span');
            pageBreak.className = 'page-break';
            return pageBreak;
        }
        return this.break == "textWrapping" ? ctx.html.createElement("br") : null;
    };
    __decorate([
        xml_serialize_1.fromAttribute("type")
    ], Break.prototype, "break", void 0);
    return Break;
}(element_base_1.ElementBase));
exports.Break = Break;


/***/ }),

/***/ "./src/elements/cell.ts":
/*!******************************!*\
  !*** ./src/elements/cell.ts ***!
  \******************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var element_base_1 = __webpack_require__(/*! ./element-base */ "./src/elements/element-base.ts");
var Cell = (function (_super) {
    __extends(Cell, _super);
    function Cell() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.props = {};
        return _this;
    }
    Cell.prototype.render = function (ctx) {
        var elem = this.renderContainer(ctx, "td");
        if (this.props.gridSpan)
            elem.colSpan = this.props.gridSpan;
        if (this.props.rowSpan) {
            elem.rowSpan = this.props.rowSpan;
        }
        if (this.props.vMerge && this.props.vMerge !== 'restart') {
            return null;
        }
        return elem;
    };
    return Cell;
}(element_base_1.ContainerBase));
exports.Cell = Cell;


/***/ }),

/***/ "./src/elements/drawing.ts":
/*!*********************************!*\
  !*** ./src/elements/drawing.ts ***!
  \*********************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var element_base_1 = __webpack_require__(/*! ./element-base */ "./src/elements/element-base.ts");
var Drawing = (function (_super) {
    __extends(Drawing, _super);
    function Drawing() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.posX = null;
        _this.posY = null;
        _this.style = {};
        return _this;
    }
    Drawing.prototype.render = function (ctx) {
        var elem = this.renderContainer(ctx, "div");
        elem.style.display = "inline-block";
        elem.style.textIndent = "0px";
        return elem;
    };
    return Drawing;
}(element_base_1.ContainerBase));
exports.Drawing = Drawing;


/***/ }),

/***/ "./src/elements/element-base.ts":
/*!**************************************!*\
  !*** ./src/elements/element-base.ts ***!
  \**************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
Object.defineProperty(exports, "__esModule", { value: true });
var utils_1 = __webpack_require__(/*! ../utils */ "./src/utils.ts");
var ElementBase = (function () {
    function ElementBase() {
    }
    ElementBase.prototype.render = function (ctx) {
        return null;
    };
    return ElementBase;
}());
exports.ElementBase = ElementBase;
var ContainerBase = (function (_super) {
    __extends(ContainerBase, _super);
    function ContainerBase() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.children = [];
        _this.childrenStyle = {};
        return _this;
    }
    ContainerBase.prototype.renderContainer = function (ctx, tagName) {
        var _this = this;
        var elem = ctx.html.createElement(tagName);
        renderStyleValues(this.style, elem);
        if (this.className)
            elem.className = utils_1.appendClass(elem.className, this.className);
        for (var _i = 0, _a = this.children.map(function (c) {
            if ((c['posX'] && c['posX'].relative !== 'page') || (c['posY'] && c['posY'].relative !== 'page')) {
                elem.style.position = "relative";
            }
            if (_this.childrenStyle) {
                c.style = __assign(__assign({}, _this.childrenStyle), c.style);
            }
            return c.render(ctx);
        }).filter(function (x) { return x != null; }); _i < _a.length; _i++) {
            var n = _a[_i];
            elem.appendChild(n);
        }
        return elem;
    };
    return ContainerBase;
}(ElementBase));
exports.ContainerBase = ContainerBase;
function renderStyleValues(style, ouput) {
    if (style == null)
        return;
    for (var key in style) {
        if (style.hasOwnProperty(key)) {
            ouput.style[key] = style[key];
        }
    }
}
exports.renderStyleValues = renderStyleValues;


/***/ }),

/***/ "./src/elements/hyperlink.ts":
/*!***********************************!*\
  !*** ./src/elements/hyperlink.ts ***!
  \***********************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
Object.defineProperty(exports, "__esModule", { value: true });
var element_base_1 = __webpack_require__(/*! ./element-base */ "./src/elements/element-base.ts");
var xml_serialize_1 = __webpack_require__(/*! ../parser/xml-serialize */ "./src/parser/xml-serialize.ts");
var Hyperlink = (function (_super) {
    __extends(Hyperlink, _super);
    function Hyperlink() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Hyperlink.prototype.render = function (ctx) {
        var a = this.renderContainer(ctx, "a");
        if (this.anchor)
            a.href = "#" + this.anchor;
        if (this.refId) {
            a.href = ctx.document.getHyperlinkTarget(this.refId);
            a.target = '_blank';
            a.rel = "noopener noreferrer";
        }
        return a;
    };
    __decorate([
        xml_serialize_1.fromAttribute("anchor")
    ], Hyperlink.prototype, "anchor", void 0);
    __decorate([
        xml_serialize_1.fromAttribute("id")
    ], Hyperlink.prototype, "refId", void 0);
    return Hyperlink;
}(element_base_1.ContainerBase));
exports.Hyperlink = Hyperlink;


/***/ }),

/***/ "./src/elements/image.ts":
/*!*******************************!*\
  !*** ./src/elements/image.ts ***!
  \*******************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var element_base_1 = __webpack_require__(/*! ./element-base */ "./src/elements/element-base.ts");
var Image = (function (_super) {
    __extends(Image, _super);
    function Image() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.style = {};
        return _this;
    }
    Image.prototype.render = function (ctx) {
        var result = ctx.html.createElement("img");
        console.log(this);
        if (ctx.document) {
            ctx.document.loadDocumentImage(this.src).then(function (x) {
                result.src = x;
            });
        }
        if (this.crop) {
            var cropWrapper = ctx.html.createElement("div");
            cropWrapper.appendChild(result);
            element_base_1.renderStyleValues(this.style, cropWrapper);
            cropWrapper.style.overflow = "hidden";
            var width = parseInt(cropWrapper.style.width.replace('pt', ''));
            var height = parseInt(cropWrapper.style.height.replace('pt', ''));
            var imageWidth = width / (1 - (this.crop.left / 100000) - (this.crop.right / 100000));
            var imageHeight = height / (1 - (this.crop.top / 100000) - (this.crop.bottom / 100000));
            result.style.marginLeft = "-" + imageWidth * this.crop.left / 100000 + "pt";
            result.style.marginRight = "-" + imageWidth * this.crop.right / 100000 + "pt";
            result.style.marginTop = "-" + imageHeight * this.crop.top / 100000 + "pt";
            result.style.marginBottom = "-" + imageHeight * this.crop.bottom / 100000 + "pt";
            result.style.width = imageWidth + "pt";
            result.style.height = imageHeight + "pt";
            return cropWrapper;
        }
        else {
            element_base_1.renderStyleValues(this.style, result);
        }
        return result;
    };
    return Image;
}(element_base_1.ElementBase));
exports.Image = Image;


/***/ }),

/***/ "./src/elements/paragraph.ts":
/*!***********************************!*\
  !*** ./src/elements/paragraph.ts ***!
  \***********************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
Object.defineProperty(exports, "__esModule", { value: true });
var element_base_1 = __webpack_require__(/*! ./element-base */ "./src/elements/element-base.ts");
var utils_1 = __webpack_require__(/*! ../utils */ "./src/utils.ts");
var xml_serialize_1 = __webpack_require__(/*! ../parser/xml-serialize */ "./src/parser/xml-serialize.ts");
var run_1 = __webpack_require__(/*! ./run */ "./src/elements/run.ts");
var Paragraph = (function (_super) {
    __extends(Paragraph, _super);
    function Paragraph() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.defaultRunStyle = {};
        _this.props = {};
        return _this;
    }
    Paragraph.prototype.render = function (ctx) {
        var elem = this.renderContainer(ctx, "p");
        if (this.props.numbering) {
            var numberingClass = ctx.numberingClass(this.props.numbering.id, this.props.numbering.level);
            elem.className = utils_1.appendClass(elem.className, numberingClass);
        }
        if (this.children.length > 0) {
            if (this.children.find(function (c) { return c instanceof run_1.Run && (!c.style || Object.keys(c.style).length === 0); })) {
                this.style = __assign(__assign({}, this.style), this.defaultRunStyle);
                element_base_1.renderStyleValues(this.style, elem);
            }
        }
        else {
            this.style = __assign(__assign({}, this.style), this.defaultRunStyle);
            element_base_1.renderStyleValues(this.style, elem);
        }
        return elem;
    };
    Paragraph = __decorate([
        xml_serialize_1.element("p")
    ], Paragraph);
    return Paragraph;
}(element_base_1.ContainerBase));
exports.Paragraph = Paragraph;


/***/ }),

/***/ "./src/elements/row.ts":
/*!*****************************!*\
  !*** ./src/elements/row.ts ***!
  \*****************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var element_base_1 = __webpack_require__(/*! ./element-base */ "./src/elements/element-base.ts");
var Row = (function (_super) {
    __extends(Row, _super);
    function Row() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Row.prototype.render = function (ctx) {
        return this.renderContainer(ctx, "tr");
    };
    return Row;
}(element_base_1.ContainerBase));
exports.Row = Row;


/***/ }),

/***/ "./src/elements/run.ts":
/*!*****************************!*\
  !*** ./src/elements/run.ts ***!
  \*****************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
Object.defineProperty(exports, "__esModule", { value: true });
var element_base_1 = __webpack_require__(/*! ./element-base */ "./src/elements/element-base.ts");
var xml_serialize_1 = __webpack_require__(/*! ../parser/xml-serialize */ "./src/parser/xml-serialize.ts");
var Run = (function (_super) {
    __extends(Run, _super);
    function Run() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.props = {};
        return _this;
    }
    Run.prototype.render = function (ctx) {
        if (this.fldCharType)
            return null;
        var elem = this.renderContainer(ctx, "span");
        var wrapper = null;
        if (this.href) {
            wrapper = ctx.html.createElement("a");
            wrapper.href = this.href;
        }
        else if (this.instrText) {
            if (this.instrText.startsWith('PAGE')) {
                elem.innerText = "" + ctx.currentPageNumber;
            }
            if (this.instrText.startsWith('NUMPAGES')) {
                elem.className = 'total-pages';
                elem.innerText = "NUMPAGES";
            }
        }
        else {
            switch (this.props.verticalAlignment) {
                case "subscript":
                    wrapper = ctx.html.createElement("sub");
                    break;
                case "superscript":
                    wrapper = ctx.html.createElement("sup");
                    break;
            }
        }
        if (wrapper == null)
            return elem;
        wrapper.appendChild(elem);
        return wrapper;
    };
    Run = __decorate([
        xml_serialize_1.element("r")
    ], Run);
    return Run;
}(element_base_1.ContainerBase));
exports.Run = Run;


/***/ }),

/***/ "./src/elements/section.ts":
/*!*********************************!*\
  !*** ./src/elements/section.ts ***!
  \*********************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var element_base_1 = __webpack_require__(/*! ./element-base */ "./src/elements/element-base.ts");
var Section = (function (_super) {
    __extends(Section, _super);
    function Section() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.props = {};
        return _this;
    }
    Section.prototype.render = function (ctx) {
        return null;
    };
    return Section;
}(element_base_1.ContainerBase));
exports.Section = Section;


/***/ }),

/***/ "./src/elements/symbol.ts":
/*!********************************!*\
  !*** ./src/elements/symbol.ts ***!
  \********************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
Object.defineProperty(exports, "__esModule", { value: true });
var element_base_1 = __webpack_require__(/*! ./element-base */ "./src/elements/element-base.ts");
var xml_serialize_1 = __webpack_require__(/*! ../parser/xml-serialize */ "./src/parser/xml-serialize.ts");
var Symbol = (function (_super) {
    __extends(Symbol, _super);
    function Symbol() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Symbol.prototype.render = function (ctx) {
        var span = ctx.html.createElement("span");
        span.style.fontFamily = this.font;
        span.innerHTML = "&#x" + this.char + ";";
        return span;
    };
    __decorate([
        xml_serialize_1.fromAttribute("font")
    ], Symbol.prototype, "font", void 0);
    __decorate([
        xml_serialize_1.fromAttribute("char")
    ], Symbol.prototype, "char", void 0);
    return Symbol;
}(element_base_1.ElementBase));
exports.Symbol = Symbol;


/***/ }),

/***/ "./src/elements/tab.ts":
/*!*****************************!*\
  !*** ./src/elements/tab.ts ***!
  \*****************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var element_base_1 = __webpack_require__(/*! ./element-base */ "./src/elements/element-base.ts");
var javascript_1 = __webpack_require__(/*! ../javascript */ "./src/javascript.ts");
var paragraph_1 = __webpack_require__(/*! ./paragraph */ "./src/elements/paragraph.ts");
var Tab = (function (_super) {
    __extends(Tab, _super);
    function Tab() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Tab.prototype.render = function (ctx) {
        var _this = this;
        var tabSpan = ctx.html.createElement("span");
        tabSpan.innerHTML = "&emsp;";
        if (ctx.options.experimental) {
            setTimeout(function () {
                var paragraph = findParent(_this);
                paragraph.props.tabs.sort(function (a, b) { return a.position.value - b.position.value; });
                tabSpan.style.display = "inline-block";
                javascript_1.updateTabStop(tabSpan, paragraph.props.tabs);
            }, 0);
        }
        return tabSpan;
    };
    return Tab;
}(element_base_1.ElementBase));
exports.Tab = Tab;
function findParent(elem) {
    var parent = elem.parent;
    while (parent != null && !(parent instanceof paragraph_1.Paragraph))
        parent = parent.parent;
    return parent;
}


/***/ }),

/***/ "./src/elements/table.ts":
/*!*******************************!*\
  !*** ./src/elements/table.ts ***!
  \*******************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var element_base_1 = __webpack_require__(/*! ./element-base */ "./src/elements/element-base.ts");
var common_1 = __webpack_require__(/*! ../dom/common */ "./src/dom/common.ts");
var Table = (function (_super) {
    __extends(Table, _super);
    function Table() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Table.prototype.render = function (ctx) {
        var elem = this.renderContainer(ctx, "table");
        if (this.columns)
            elem.appendChild(this.renderTableColumns(ctx, this.columns));
        return elem;
    };
    Table.prototype.renderTableColumns = function (ctx, columns) {
        var result = ctx.html.createElement("colGroup");
        for (var _i = 0, columns_1 = columns; _i < columns_1.length; _i++) {
            var col = columns_1[_i];
            var colElem = ctx.html.createElement("col");
            if (col.width)
                colElem.style.width = common_1.renderLength(col.width);
            result.appendChild(colElem);
        }
        return result;
    };
    return Table;
}(element_base_1.ContainerBase));
exports.Table = Table;


/***/ }),

/***/ "./src/elements/text.ts":
/*!******************************!*\
  !*** ./src/elements/text.ts ***!
  \******************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
Object.defineProperty(exports, "__esModule", { value: true });
var element_base_1 = __webpack_require__(/*! ./element-base */ "./src/elements/element-base.ts");
var xml_serialize_1 = __webpack_require__(/*! ../parser/xml-serialize */ "./src/parser/xml-serialize.ts");
var Text = (function (_super) {
    __extends(Text, _super);
    function Text() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Text.prototype.render = function (context) {
        return context.html.createTextNode(this.text);
    };
    __decorate([
        xml_serialize_1.fromText()
    ], Text.prototype, "text", void 0);
    Text = __decorate([
        xml_serialize_1.element("t")
    ], Text);
    return Text;
}(element_base_1.ElementBase));
exports.Text = Text;
var Symbol = (function (_super) {
    __extends(Symbol, _super);
    function Symbol() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Symbol.prototype.render = function (ctx) {
        var span = ctx.html.createElement("span");
        span.style.fontFamily = this.font;
        span.innerHTML = "&#x" + this.char + ";";
        return span;
    };
    return Symbol;
}(element_base_1.ElementBase));
exports.Symbol = Symbol;


/***/ }),

/***/ "./src/html-renderer.ts":
/*!******************************!*\
  !*** ./src/html-renderer.ts ***!
  \******************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var __spreadArrays = (this && this.__spreadArrays) || function () {
    for (var s = 0, i = 0, il = arguments.length; i < il; i++) s += arguments[i].length;
    for (var r = Array(s), k = 0, i = 0; i < il; i++)
        for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++)
            r[k] = a[j];
    return r;
};
Object.defineProperty(exports, "__esModule", { value: true });
var element_base_1 = __webpack_require__(/*! ./elements/element-base */ "./src/elements/element-base.ts");
var break_1 = __webpack_require__(/*! ./elements/break */ "./src/elements/break.ts");
var paragraph_1 = __webpack_require__(/*! ./elements/paragraph */ "./src/elements/paragraph.ts");
var table_1 = __webpack_require__(/*! ./elements/table */ "./src/elements/table.ts");
var render_context_1 = __webpack_require__(/*! ./dom/render-context */ "./src/dom/render-context.ts");
var section_1 = __webpack_require__(/*! ./elements/section */ "./src/elements/section.ts");
var HtmlRenderer = (function () {
    function HtmlRenderer(htmlDocument) {
        this.htmlDocument = htmlDocument;
        this.inWrapper = true;
        this.className = "docx";
        this._renderContext = new render_context_1.RenderContext();
        this._renderContext.html = htmlDocument;
    }
    HtmlRenderer.prototype.render = function (document, bodyContainer, styleContainer, options) {
        if (styleContainer === void 0) { styleContainer = null; }
        this.document = document;
        this.options = options;
        this._renderContext.options = options;
        this._renderContext.className = this.className;
        this._renderContext.document = document;
        styleContainer = styleContainer || bodyContainer;
        removeAllElements(styleContainer);
        removeAllElements(bodyContainer);
        appendComment(styleContainer, "docxjs library predefined styles");
        styleContainer.appendChild(this.renderDefaultStyle());
        appendComment(styleContainer, "docx document styles");
        styleContainer.appendChild(this.renderStyles(document.styles));
        if (document.numbering) {
            appendComment(styleContainer, "docx document numbering styles");
            styleContainer.appendChild(this.renderNumbering(document.numbering, styleContainer));
        }
        if (!options.ignoreFonts)
            this.renderFontTable(document.fontTable, styleContainer);
        var wrapper = this.renderWrapper();
        bodyContainer.appendChild(wrapper);
        var container = this.inWrapper ? wrapper : bodyContainer;
        this.renderSections(container, document.document);
    };
    HtmlRenderer.prototype.renderFontTable = function (fonts, styleContainer) {
        for (var _i = 0, _a = fonts.filter(function (x) { return !x.refId; }); _i < _a.length; _i++) {
            var f = _a[_i];
            appendComment(styleContainer, "Importing Google Font " + f.name);
            styleContainer.appendChild(createGoogleFontElement(f.name));
        }
        var _loop_1 = function (f) {
            this_1.document.loadFont(f.refId, f.fontKey).then(function (fontData) {
                var cssTest = "@font-face {\n                    font-family: \"" + f.name + "\";\n                    src: url(" + fontData + ");\n                }";
                appendComment(styleContainer, "Font " + f.name);
                styleContainer.appendChild(createStyleElement(cssTest));
            });
        };
        var this_1 = this;
        for (var _b = 0, _c = fonts.filter(function (x) { return x.refId; }); _b < _c.length; _b++) {
            var f = _c[_b];
            _loop_1(f);
        }
        var fontFamily = "." + this.className + "-wrapper {\n            font-family: " + fonts.map(function (f) { return "\"" + f.name + "\""; }).join(',') + ";\n        }";
        appendComment(styleContainer, "Apply fonts");
        styleContainer.appendChild(createStyleElement(fontFamily));
    };
    HtmlRenderer.prototype.processClassName = function (className) {
        if (!className)
            return this.className;
        return this.className + "_" + className;
    };
    HtmlRenderer.prototype.processStyles = function (styles) {
        var stylesMap = {};
        for (var _i = 0, _a = styles.filter(function (x) { return x.id != null; }); _i < _a.length; _i++) {
            var style = _a[_i];
            stylesMap[style.id] = style;
        }
        for (var _b = 0, _c = styles.filter(function (x) { return x.basedOn; }); _b < _c.length; _b++) {
            var style = _c[_b];
            var baseStyle = stylesMap[style.basedOn];
            if (baseStyle) {
                var _loop_2 = function (styleValues) {
                    baseValues = baseStyle.styles.filter(function (x) { return x.target == styleValues.target; });
                    if (baseValues && baseValues.length > 0)
                        this_2.copyStyleProperties(baseValues[0].values, styleValues.values);
                };
                var this_2 = this, baseValues;
                for (var _d = 0, _e = style.styles; _d < _e.length; _d++) {
                    var styleValues = _e[_d];
                    _loop_2(styleValues);
                }
            }
            else if (this.options.debug)
                console.warn("Can't find base style " + style.basedOn);
        }
        for (var _f = 0, styles_1 = styles; _f < styles_1.length; _f++) {
            var style = styles_1[_f];
            style.id = this.processClassName(style.id);
        }
        return stylesMap;
    };
    HtmlRenderer.prototype.processElement = function (element) {
        if (element.children) {
            for (var _i = 0, _a = element.children; _i < _a.length; _i++) {
                var e = _a[_i];
                e.className = this.processClassName(e.className);
                e.parent = element;
                if (e instanceof table_1.Table) {
                    this.processTable(e);
                }
                else {
                    this.processElement(e);
                }
            }
        }
    };
    HtmlRenderer.prototype.processTable = function (table) {
        for (var _i = 0, _a = table.children; _i < _a.length; _i++) {
            var r = _a[_i];
            for (var _b = 0, _c = r.children; _b < _c.length; _b++) {
                var c = _c[_b];
                c.style = this.copyStyleProperties(table.cellStyle, c.style, [
                    "border-left", "border-right", "border-top", "border-bottom",
                    "padding-left", "padding-right", "padding-top", "padding-bottom"
                ]);
                this.processElement(c);
            }
        }
    };
    HtmlRenderer.prototype.copyStyleProperties = function (input, output, attrs) {
        if (attrs === void 0) { attrs = null; }
        if (!input)
            return output;
        if (output == null)
            output = {};
        if (attrs == null)
            attrs = Object.getOwnPropertyNames(input);
        for (var _i = 0, attrs_1 = attrs; _i < attrs_1.length; _i++) {
            var key = attrs_1[_i];
            if (input.hasOwnProperty(key) && !output.hasOwnProperty(key))
                output[key] = input[key];
        }
        return output;
    };
    HtmlRenderer.prototype.createSection = function (className, props, header, footer) {
        var elem = this.htmlDocument.createElement("section");
        elem.style.position = "relative";
        elem.className = className;
        if (props) {
            if (props.pageMargins) {
                elem.style.paddingLeft = this.renderLength(props.pageMargins.left);
                elem.style.paddingRight = this.renderLength(props.pageMargins.right);
                elem.style.paddingTop = header ? this.renderLength(props.pageMargins.header) : this.renderLength(props.pageMargins.top);
                elem.style.paddingBottom = footer ? this.renderLength(props.pageMargins.footer) : this.renderLength(props.pageMargins.bottom);
            }
            if (props.pageSize) {
                if (!this.options.ignoreWidth)
                    elem.style.width = this.renderLength(props.pageSize.width);
                if (!this.options.ignoreHeight)
                    elem.style.height = this.renderLength(props.pageSize.height);
            }
            if (props.columns && props.columns.numberOfColumns) {
                elem.style.columnCount = "" + props.columns.numberOfColumns;
                elem.style.columnGap = this.renderLength(props.columns.space);
                if (props.columns.separator) {
                    elem.style.columnRule = "1px solid black";
                }
            }
        }
        return elem;
    };
    HtmlRenderer.prototype.renderSections = function (into, document) {
        return __awaiter(this, void 0, void 0, function () {
            var result, sections, sectionNumber, section, sectionProps, pickedHeaderRef, pickedFooterRef, sectionElement, pickedHeader, header, main, pickedFooter, footer, remainingElementsAfterConstraintReached, newSection;
            var _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        result = [];
                        this.processElement(document);
                        sections = document.children;
                        sectionNumber = 1;
                        _b.label = 1;
                    case 1:
                        if (!(sections.length > 0)) return [3, 6];
                        this._renderContext.currentPageNumber = sectionNumber;
                        section = sections.shift();
                        if (!(section instanceof section_1.Section)) {
                            return [2, []];
                        }
                        sectionProps = section.props;
                        pickedHeaderRef = this.pickHeaderOrFooterRef(sectionProps.headers || {}, sectionNumber);
                        pickedFooterRef = this.pickHeaderOrFooterRef(sectionProps.footers || {}, sectionNumber);
                        sectionElement = this.createSection(this.className, sectionProps, pickedHeaderRef, pickedFooterRef);
                        into.appendChild(sectionElement);
                        if (!pickedHeaderRef) return [3, 3];
                        return [4, this.document.loadAndSetRenderContextToHeaderOrFooter(pickedHeaderRef)];
                    case 2:
                        pickedHeader = _b.sent();
                        if (pickedHeader) {
                            header = this.htmlDocument.createElement("header");
                            this.renderElements(pickedHeader.children, header);
                            sectionElement.appendChild(header);
                        }
                        _b.label = 3;
                    case 3:
                        main = this.htmlDocument.createElement("main");
                        sectionElement.appendChild(main);
                        if (!pickedFooterRef) return [3, 5];
                        return [4, this.document.loadAndSetRenderContextToHeaderOrFooter(pickedFooterRef)];
                    case 4:
                        pickedFooter = _b.sent();
                        if (pickedFooter) {
                            footer = this.htmlDocument.createElement("footer");
                            this.renderElements(pickedFooter.children, footer);
                            sectionElement.appendChild(footer);
                        }
                        _b.label = 5;
                    case 5:
                        this.document.setRenderContextToMainDocument();
                        remainingElementsAfterConstraintReached = this.renderElements(section.children, main, true).remainingElementsAfterConstraintReached;
                        if (remainingElementsAfterConstraintReached && remainingElementsAfterConstraintReached.length > 0) {
                            if (sections.length > 0) {
                                (_a = sections[0].children).unshift.apply(_a, remainingElementsAfterConstraintReached);
                            }
                            else {
                                newSection = new section_1.Section();
                                newSection.props = sectionProps;
                                newSection.children = remainingElementsAfterConstraintReached;
                                sections.push(newSection);
                            }
                        }
                        result.push(sectionElement);
                        sectionNumber++;
                        return [3, 1];
                    case 6:
                        this.htmlDocument.querySelectorAll('.total-pages').forEach(function (elem) { return elem.innerText = "" + (sectionNumber - 1); });
                        return [2, result];
                }
            });
        });
    };
    HtmlRenderer.prototype.pickHeaderOrFooterRef = function (headersOrFootersByType, sectionNumber) {
        if (headersOrFootersByType['first'] && sectionNumber === 1) {
            return headersOrFootersByType['first'].refId;
        }
        else if (headersOrFootersByType['even'] && sectionNumber % 2 === 0) {
            return headersOrFootersByType['even'].refId;
        }
        else if (headersOrFootersByType['default']) {
            return headersOrFootersByType['default'].refId;
        }
        return undefined;
    };
    HtmlRenderer.prototype.splitBySection = function (elements) {
        var current = { sectProps: null, elements: [] };
        var result = [current];
        function splitElement(elem, revert) {
            var children = elem.children;
            var newElem = Object.create(Object.getPrototypeOf(elem));
            Object.assign(newElem, elem);
            var _a = revert ? [elem, newElem] : [newElem, elem], f = _a[0], s = _a[1];
            f.children = children.slice(pBreakIndex);
            s.children = children.slice(0, rBreakIndex);
            return newElem;
        }
        for (var _i = 0, elements_1 = elements; _i < elements_1.length; _i++) {
            var elem = elements_1[_i];
            current.elements.push(elem);
            if (elem instanceof paragraph_1.Paragraph) {
                var sectProps = elem.props.sectionProps;
                var pBreakIndex = -1;
                var rBreakIndex = -1;
                if (this.options.breakPages && elem.children) {
                    pBreakIndex = elem.children.findIndex(function (r) {
                        var _a, _b;
                        rBreakIndex = (_b = (_a = r.children) === null || _a === void 0 ? void 0 : _a.findIndex(function (t) { return (t instanceof break_1.Break) && t.break == "page"; })) !== null && _b !== void 0 ? _b : -1;
                        return rBreakIndex != -1;
                    });
                }
                if (!sectProps) {
                }
                if (sectProps || pBreakIndex != -1) {
                    current.sectProps = sectProps;
                    current = { sectProps: null, elements: [] };
                    result.push(current);
                }
                if (pBreakIndex != -1) {
                    var breakRun = elem.children[pBreakIndex];
                    var splitRun = rBreakIndex < breakRun.children.length - 1;
                    if (pBreakIndex < elem.children.length - 1 || splitRun) {
                        current.elements.push(splitElement(elem, false));
                        if (splitRun) {
                            elem.children.push(splitElement(breakRun, true));
                        }
                    }
                }
            }
        }
        return result;
    };
    HtmlRenderer.prototype.renderLength = function (l) {
        return !l ? null : "" + l.value + l.type;
    };
    HtmlRenderer.prototype.renderWrapper = function () {
        var wrapper = document.createElement("div");
        wrapper.className = this.className + "-wrapper";
        return wrapper;
    };
    HtmlRenderer.prototype.renderDefaultStyle = function () {
        var styleText = "." + this.className + "-wrapper { background: gray; padding: 30px; padding-bottom: 0px; display: flex; flex-flow: column; align-items: center; } \n                ." + this.className + "-wrapper section." + this.className + " { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }\n                ." + this.className + " { color: black; }\n                section." + this.className + " { box-sizing: border-box; }\n                ." + this.className + " table { border-collapse: collapse; }\n                ." + this.className + " table td, ." + this.className + " table th { vertical-align: top; }\n                ." + this.className + " p { margin: 0pt; }\n                ." + this.className + " p:empty:before { content: ' '; white-space: pre; }\n                \n                section." + this.className + " {\n                    display: flex;\n                    flex-flow: column;\n                    height: 100%;\n                  }\n                  \n                  section." + this.className + " header {\n                    flex: 0 1 auto;\n                  }\n                  \n                  section." + this.className + " main {\n                    flex: 1 1 auto;\n                  }\n                  \n                  section." + this.className + " footer {\n                    flex: 0 1 auto;\n                  }\n                \n                ";
        return createStyleElement(styleText);
    };
    HtmlRenderer.prototype.renderNumbering = function (styles, styleContainer) {
        var _this = this;
        var styleText = "";
        var rootCounters = [];
        var _loop_3 = function () {
            selector = "p." + this_3.numberingClass(num.id, num.level);
            listStyleType = "none";
            if (num.levelText && num.format == "decimal") {
                var counter = this_3.numberingCounter(num.id, num.level);
                if (num.level > 0) {
                    styleText += this_3.styleToString("p." + this_3.numberingClass(num.id, num.level - 1), {
                        "counter-reset": counter
                    });
                }
                else {
                    rootCounters.push(counter);
                }
                styleText += this_3.styleToString(selector + ":before", {
                    "content": this_3.levelTextToContent(num.levelText, num.id),
                    "counter-increment": counter,
                    "margin-right": "8pt",
                });
                styleText += this_3.styleToString(selector, __assign({ "display": "list-item", "list-style-position": "inside", "list-style-type": "none" }, num.style));
            }
            else if (num.bullet) {
                var valiable_1 = ("--" + this_3.className + "-" + num.bullet.src).toLowerCase();
                styleText += this_3.styleToString(selector + ":before", {
                    "content": "' '",
                    "display": "inline-block",
                    "background": "var(" + valiable_1 + ")"
                }, num.bullet.style);
                this_3.document.loadNumberingImage(num.bullet.src).then(function (data) {
                    var text = "." + _this.className + "-wrapper { " + valiable_1 + ": url(" + data + ") }";
                    styleContainer.appendChild(createStyleElement(text));
                });
            }
            else {
                listStyleType = this_3.numFormatToCssValue(num.format);
            }
            styleText += this_3.styleToString(selector, __assign({ "display": "list-item", "list-style-position": "inside", "list-style-type": listStyleType }, num.style));
        };
        var this_3 = this, selector, listStyleType;
        for (var _i = 0, styles_2 = styles; _i < styles_2.length; _i++) {
            var num = styles_2[_i];
            _loop_3();
        }
        if (rootCounters.length > 0) {
            styleText += this.styleToString("." + this.className + "-wrapper", {
                "counter-reset": rootCounters.join(" ")
            });
        }
        return createStyleElement(styleText);
    };
    HtmlRenderer.prototype.renderStyles = function (styles) {
        var styleText = "";
        var stylesMap = this.processStyles(styles);
        for (var _i = 0, styles_3 = styles; _i < styles_3.length; _i++) {
            var style = styles_3[_i];
            var subStyles = style.styles;
            if (style.linked) {
                var linkedStyle = style.linked && stylesMap[style.linked];
                if (linkedStyle)
                    subStyles = subStyles.concat(linkedStyle.styles);
                else if (this.options.debug)
                    console.warn("Can't find linked style " + style.linked);
            }
            for (var _a = 0, subStyles_1 = subStyles; _a < subStyles_1.length; _a++) {
                var subStyle = subStyles_1[_a];
                var selector = "";
                if (style.target == subStyle.target)
                    selector += style.target + "." + style.id;
                else if (style.target)
                    selector += style.target + "." + style.id + " " + subStyle.target;
                else
                    selector += "." + style.id + " " + subStyle.target;
                if (style.isDefault && style.target)
                    selector = "." + this.className + " " + style.target + ", " + selector;
                styleText += this.styleToString(selector, subStyle.values);
            }
        }
        return createStyleElement(styleText);
    };
    HtmlRenderer.prototype.renderElements = function (elems, into, heightConstrained) {
        if (elems == null)
            return null;
        var appendedElements = [];
        var remainingElements = __spreadArrays(elems);
        if (into) {
            for (var _i = 0, elems_1 = elems; _i < elems_1.length; _i++) {
                var c = elems_1[_i];
                if (!(c instanceof element_base_1.ElementBase)) {
                    continue;
                }
                var renderedElement = c.render(this._renderContext);
                if (!renderedElement) {
                    continue;
                }
                var containerBeforeHeight = into.getBoundingClientRect().height;
                into.appendChild(renderedElement);
                var containerAfterHeight = into.getBoundingClientRect().height;
                var containsPageBreak = renderedElement.querySelector('.page-break');
                if (containsPageBreak && appendedElements.length !== 0) {
                    into.removeChild(renderedElement);
                    return { renderedElements: appendedElements, remainingElementsAfterConstraintReached: remainingElements };
                }
                else if (heightConstrained && containerBeforeHeight !== containerAfterHeight) {
                    into.removeChild(renderedElement);
                    return { renderedElements: appendedElements, remainingElementsAfterConstraintReached: remainingElements };
                }
                else {
                    appendedElements.push(renderedElement);
                    remainingElements.shift();
                }
            }
        }
        return { renderedElements: appendedElements, remainingElementsAfterConstraintReached: [] };
    };
    HtmlRenderer.prototype.numberingClass = function (id, lvl) {
        return this.className + "-num-" + id + "-" + lvl;
    };
    HtmlRenderer.prototype.styleToString = function (selectors, values, cssText) {
        if (cssText === void 0) { cssText = null; }
        var result = selectors + " {\r\n";
        for (var key in values) {
            result += "  " + key + ": " + values[key] + ";\r\n";
        }
        if (cssText)
            result += ";" + cssText;
        return result + "}\r\n";
    };
    HtmlRenderer.prototype.numberingCounter = function (id, lvl) {
        return this.className + "-num-" + id + "-" + lvl;
    };
    HtmlRenderer.prototype.levelTextToContent = function (text, id) {
        var _this = this;
        var result = text.replace(/%\d*/g, function (s) {
            var lvl = parseInt(s.substring(1), 10) - 1;
            return "\"counter(" + _this.numberingCounter(id, lvl) + ")\"";
        });
        return '"' + result + '"';
    };
    HtmlRenderer.prototype.numFormatToCssValue = function (format) {
        var mapping = {
            "none": "none",
            "bullet": "disc",
            "decimal": "decimal",
            "lowerLetter": "lower-alpha",
            "upperLetter": "upper-alpha",
            "lowerRoman": "lower-roman",
            "upperRoman": "upper-roman",
        };
        return mapping[format] || format;
    };
    return HtmlRenderer;
}());
exports.HtmlRenderer = HtmlRenderer;
function appentElements(container, children) {
    for (var _i = 0, children_1 = children; _i < children_1.length; _i++) {
        var c = children_1[_i];
        container.appendChild(c);
    }
}
function removeAllElements(elem) {
    while (elem.firstChild) {
        elem.removeChild(elem.firstChild);
    }
}
function createStyleElement(cssText) {
    var styleElement = document.createElement("style");
    styleElement.type = "text/css";
    styleElement.innerHTML = cssText;
    return styleElement;
}
function createGoogleFontElement(fontName) {
    var fontLinkElement = document.createElement("link");
    fontLinkElement.rel = "stylesheet";
    fontLinkElement.href = "https://fonts.googleapis.com/css?family=" + fontName;
    return fontLinkElement;
}
function appendComment(elem, comment) {
    elem.appendChild(document.createComment(comment));
}


/***/ }),

/***/ "./src/javascript.ts":
/*!***************************!*\
  !*** ./src/javascript.ts ***!
  \***************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
function updateTabStop(elem, tabs, pixelToPoint) {
    if (pixelToPoint === void 0) { pixelToPoint = 72 / 96; }
    var p = elem.closest("p");
    var tbb = elem.getBoundingClientRect();
    var pbb = p.getBoundingClientRect();
    var left = (tbb.left - pbb.left) * pixelToPoint;
    var tab = tabs.find(function (t) { return t.style != "clear" && t.position.value > left; });
    if (tab == null)
        return;
    elem.style.display = "inline-block";
    elem.style.width = (tab.position.value - left) + "pt";
    switch (tab.leader) {
        case "dot":
        case "middleDot":
            elem.style.borderBottom = "1px black dotted";
            break;
        case "hyphen":
        case "heavy":
        case "underscore":
            elem.style.borderBottom = "1px black solid";
            break;
    }
}
exports.updateTabStop = updateTabStop;


/***/ }),

/***/ "./src/parser/common.ts":
/*!******************************!*\
  !*** ./src/parser/common.ts ***!
  \******************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var common_1 = __webpack_require__(/*! ../dom/common */ "./src/dom/common.ts");
function elements(elem, namespaceURI, localName) {
    if (namespaceURI === void 0) { namespaceURI = null; }
    if (localName === void 0) { localName = null; }
    var result = [];
    for (var i = 0; i < elem.childNodes.length; i++) {
        var n = elem.childNodes[i];
        if (n.nodeType == 1
            && (namespaceURI == null || n.namespaceURI == namespaceURI)
            && (localName == null || n.localName == localName))
            result.push(n);
    }
    return result;
}
exports.elements = elements;
function stringAttr(elem, namespaceURI, name) {
    return elem.getAttributeNS(namespaceURI, name);
}
exports.stringAttr = stringAttr;
function intAttr(elem, namespaceURI, name) {
    var val = elem.getAttributeNS(namespaceURI, name);
    return val ? parseInt(val) : null;
}
exports.intAttr = intAttr;
function colorAttr(elem, namespaceURI, name) {
    var val = elem.getAttributeNS(namespaceURI, name);
    return val ? "#" + val : null;
}
exports.colorAttr = colorAttr;
function boolAttr(elem, namespaceURI, name, defaultValue) {
    if (defaultValue === void 0) { defaultValue = false; }
    var val = elem.getAttributeNS(namespaceURI, name);
    if (val == null)
        return defaultValue;
    return val === "true" || val === "1";
}
exports.boolAttr = boolAttr;
exports.LengthUsage = {
    Dxa: { mul: 0.05, unit: "pt" },
    Emu: { mul: 1 / 12700, unit: "pt" },
    FontSize: { mul: 0.5, unit: "pt" },
    Border: { mul: 0.125, unit: "pt" },
    Percent: { mul: 0.02, unit: "%" },
    LineHeight: { mul: 1 / 240, unit: null }
};
function lengthAttr(elem, namespaceURI, name, usage) {
    if (usage === void 0) { usage = exports.LengthUsage.Dxa; }
    var val = elem.getAttributeNS(namespaceURI, name);
    return val ? { value: parseInt(val) * usage.mul, type: usage.unit } : null;
}
exports.lengthAttr = lengthAttr;
function parseBorder(elem) {
    return {
        type: stringAttr(elem, common_1.ns.wordml, "val"),
        color: colorAttr(elem, common_1.ns.wordml, "color"),
        size: lengthAttr(elem, common_1.ns.wordml, "sz", exports.LengthUsage.Border)
    };
}
exports.parseBorder = parseBorder;
function parseBorders(elem) {
    var result = {};
    for (var _i = 0, _a = elements(elem, common_1.ns.wordml); _i < _a.length; _i++) {
        var e = _a[_i];
        switch (e.localName) {
            case "left":
                result.left = parseBorder(e);
                break;
            case "top":
                result.top = parseBorder(e);
                break;
            case "right":
                result.right = parseBorder(e);
                break;
            case "botton":
                result.botton = parseBorder(e);
                break;
        }
    }
    return result;
}
exports.parseBorders = parseBorders;


/***/ }),

/***/ "./src/parser/paragraph.ts":
/*!*********************************!*\
  !*** ./src/parser/paragraph.ts ***!
  \*********************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var xml = __webpack_require__(/*! ./common */ "./src/parser/common.ts");
var common_1 = __webpack_require__(/*! ../dom/common */ "./src/dom/common.ts");
var section_1 = __webpack_require__(/*! ./section */ "./src/parser/section.ts");
function parseParagraphProperties(elem, props) {
    if (elem.namespaceURI != common_1.ns.wordml)
        return false;
    switch (elem.localName) {
        case "tabs":
            props.tabs = parseTabs(elem);
            break;
        case "sectPr":
            props.sectionProps = section_1.parseSectionProperties(elem);
            break;
        case "numPr":
            props.numbering = parseNumbering(elem);
            break;
        case "spacing":
            props.lineSpacing = parseLineSpacing(elem);
            return false;
            break;
        default:
            return false;
    }
    return true;
}
exports.parseParagraphProperties = parseParagraphProperties;
function parseTabs(elem) {
    return xml.elements(elem, common_1.ns.wordml, "tab")
        .map(function (e) { return ({
        position: xml.lengthAttr(e, common_1.ns.wordml, "pos"),
        leader: xml.stringAttr(e, common_1.ns.wordml, "leader"),
        style: xml.stringAttr(e, common_1.ns.wordml, "val")
    }); });
}
function parseNumbering(elem) {
    var result = {};
    for (var _i = 0, _a = xml.elements(elem, common_1.ns.wordml); _i < _a.length; _i++) {
        var e = _a[_i];
        switch (e.localName) {
            case "numId":
                result.id = xml.stringAttr(e, common_1.ns.wordml, "val");
                break;
            case "ilvl":
                result.level = xml.intAttr(e, common_1.ns.wordml, "val");
                break;
        }
    }
    return result;
}
function parseLineSpacing(elem) {
    return {
        before: xml.lengthAttr(elem, common_1.ns.wordml, "before"),
        after: xml.lengthAttr(elem, common_1.ns.wordml, "after"),
        line: xml.intAttr(elem, common_1.ns.wordml, "line"),
        lineRule: xml.stringAttr(elem, common_1.ns.wordml, "lineRule")
    };
}


/***/ }),

/***/ "./src/parser/section.ts":
/*!*******************************!*\
  !*** ./src/parser/section.ts ***!
  \*******************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
Object.defineProperty(exports, "__esModule", { value: true });
var common_1 = __webpack_require__(/*! ../dom/common */ "./src/dom/common.ts");
var xml = __webpack_require__(/*! ./common */ "./src/parser/common.ts");
function parseSectionProperties(elem) {
    var _a, _b;
    var section = {};
    for (var _i = 0, _c = xml.elements(elem, common_1.ns.wordml); _i < _c.length; _i++) {
        var e = _c[_i];
        switch (e.localName) {
            case "pgSz":
                section.pageSize = {
                    width: xml.lengthAttr(e, common_1.ns.wordml, "w"),
                    height: xml.lengthAttr(e, common_1.ns.wordml, "h"),
                    orientation: xml.stringAttr(e, common_1.ns.wordml, "orient")
                };
                break;
            case "pgMar":
                section.pageMargins = {
                    left: xml.lengthAttr(e, common_1.ns.wordml, "left"),
                    right: xml.lengthAttr(e, common_1.ns.wordml, "right"),
                    top: xml.lengthAttr(e, common_1.ns.wordml, "top"),
                    bottom: xml.lengthAttr(e, common_1.ns.wordml, "bottom"),
                    header: xml.lengthAttr(e, common_1.ns.wordml, "header"),
                    footer: xml.lengthAttr(e, common_1.ns.wordml, "footer"),
                    gutter: xml.lengthAttr(e, common_1.ns.wordml, "gutter"),
                };
                break;
            case "cols":
                section.columns = parseColumns(e);
                break;
            case "footerReference":
                section.footers = __assign(__assign({}, section.footers), (_a = {}, _a[xml.stringAttr(e, common_1.ns.wordml, "type")] = { refId: xml.stringAttr(e, common_1.ns.relationships, "id") }, _a));
                break;
            case "headerReference":
                section.headers = __assign(__assign({}, section.headers), (_b = {}, _b[xml.stringAttr(e, common_1.ns.wordml, "type")] = { refId: xml.stringAttr(e, common_1.ns.relationships, "id") }, _b));
                break;
        }
    }
    return section;
}
exports.parseSectionProperties = parseSectionProperties;
function parseColumns(elem) {
    return {
        numberOfColumns: xml.intAttr(elem, common_1.ns.wordml, "num"),
        space: xml.lengthAttr(elem, common_1.ns.wordml, "space"),
        separator: xml.boolAttr(elem, common_1.ns.wordml, "sep"),
        equalWidth: xml.boolAttr(elem, common_1.ns.wordml, "equalWidth", true),
        columns: xml.elements(elem, common_1.ns.wordml, "col")
            .map(function (e) { return ({
            width: xml.lengthAttr(e, common_1.ns.wordml, "w"),
            space: xml.lengthAttr(e, common_1.ns.wordml, "space")
        }); })
    };
}


/***/ }),

/***/ "./src/parser/xml-serialize.ts":
/*!*************************************!*\
  !*** ./src/parser/xml-serialize.ts ***!
  \*************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var schemaSymbol = Symbol("open-xml-schema");
function element(name) {
    return function (target) {
        var schema = getPrototypeXmlSchema(target);
        schema.elemName = name;
    };
}
exports.element = element;
function fromText(convert) {
    if (convert === void 0) { convert = null; }
    return function (target, prop) {
        var schema = getPrototypeXmlSchema(target);
        schema.text = { prop: prop, convert: convert };
    };
}
exports.fromText = fromText;
function fromAttribute(attrName, convert) {
    if (convert === void 0) { convert = null; }
    return function (target, prop) {
        var schema = getPrototypeXmlSchema(target);
        schema.attrs[attrName] = { prop: prop, convert: convert };
    };
}
exports.fromAttribute = fromAttribute;
function deserialize(n, output) {
    var proto = Object.getPrototypeOf(output);
    var schema = proto[schemaSymbol];
    if (schema == null)
        return output;
    if (schema.text) {
        var prop = schema.text;
        output[prop.prop] = prop.convert ? prop.convert(n.textContent) : n.textContent;
    }
    for (var i = 0, l = n.attributes.length; i < l; i++) {
        var attr = n.attributes.item(i);
        var prop = schema.attrs[attr.localName];
        if (prop == null)
            continue;
        output[prop.prop] = prop.convert ? prop.convert(attr.value) : attr.value;
    }
    return output;
}
exports.deserialize = deserialize;
function getPrototypeXmlSchema(proto) {
    return proto[schemaSymbol] || (proto[schemaSymbol] = {
        text: null,
        attrs: {}
    });
}


/***/ }),

/***/ "./src/utils.ts":
/*!**********************!*\
  !*** ./src/utils.ts ***!
  \**********************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
function addElementClass(element, className) {
    return element.className = appendClass(element.className, className);
}
exports.addElementClass = addElementClass;
function appendClass(classList, className) {
    return (!classList) ? className : classList + " " + className;
}
exports.appendClass = appendClass;


/***/ }),

/***/ "jszip":
/*!************************!*\
  !*** external "JSZip" ***!
  \************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_jszip__;

/***/ })

/******/ });
});
//# sourceMappingURL=docx-preview.js.map