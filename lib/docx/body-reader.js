exports.createBodyReader = createBodyReader;
exports._readNumberingProperties = readNumberingProperties;
exports._readSpacingProperties = readSpacingProperties;

var dingbatToUnicode = require("dingbat-to-unicode");
var _ = require("underscore");

var documents = require("../documents");
var Result = require("../results").Result;
var warning = require("../results").warning;
var xml = require("../xml");
var transforms = require("../transforms");
var uris = require("./uris");

function createBodyReader(options) {
    return {
        readXmlElement: function(element) {
            return new BodyReader(options).readXmlElement(element);
        },
        readXmlElements: function(elements) {
            return new BodyReader(options).readXmlElements(elements);
        }
    };
}

function BodyReader(options) {
    var complexFieldStack = [];
    var currentInstrText = [];
    var insertions = [];
    var deletions = [];

    // When a paragraph is marked as deleted, its contents should be combined
    // with the following paragraph. See 17.13.5.15 del (Deleted Paragraph) of
    // ECMA-376 4th edition Part 1.
    var deletedParagraphContents = [];

    var relationships = options.relationships;
    var contentTypes = options.contentTypes;
    var docxFile = options.docxFile;
    var files = options.files;
    var numbering = options.numbering;
    var styles = options.styles;

    function readXmlElements(elements) {
        var results = elements.map(readXmlElement);
        return combineResults(results);
    }

    function readXmlElement(element) {
        if (element.type === "element") {
            var handler = xmlElementReaders[element.name];
            if (handler) {
                return handler(element);
            } else if (!Object.prototype.hasOwnProperty.call(ignoreElements, element.name)) {
                var message = warning("An unrecognised element was ignored: " + element.name);
                return emptyResultWithMessages([message]);
            }
        }
        return emptyResult();
    }

    function readParagraphProperties(element) {
        return readParagraphStyle(element).map(function(style) {
            var numberingAttrs = readNumberingProperties(style.styleId, element.firstOrEmpty("w:numPr"), numbering);
            return {
                type: "paragraphProperties",
                styleId: style.styleId,
                styleName: style.name,
                alignment: element.firstOrEmpty("w:jc").attributes["w:val"],
                numbering: numberingAttrs,
                indent: readParagraphIndent(element.firstOrEmpty("w:ind")),
                spacing: readSpacingProperties(element)
            };
        });
    }

    function readParagraphIndent(element) {
        return {
            left: element.attributes["w:left"],
            right: element.attributes["w:right"],
            start: element.attributes["w:start"],
            end: element.attributes["w:end"],
            firstLine: element.attributes["w:firstLine"],
            hanging: element.attributes["w:hanging"]
        };
    }

    function readRunProperties(element) {
        return readRunStyle(element).map(function(style) {
            var fontSizeString = element.firstOrEmpty("w:sz").attributes["w:val"];
            // w:sz gives the font size in half points, so halve the value to get the size in points
            var fontSize = /^[0-9]+$/.test(fontSizeString) ? parseInt(fontSizeString, 10) / 2 : null;

            return {
                type: "runProperties",
                styleId: style.styleId,
                styleName: style.name,
                verticalAlignment: element.firstOrEmpty("w:vertAlign").attributes["w:val"],
                font: element.firstOrEmpty("w:rFonts").attributes["w:ascii"],
                fontSize: fontSize,
                color: element.firstOrEmpty("w:color").attributes["w:val"],
                isBold: readBooleanElement(element.first("w:b")),
                isUnderline: readUnderline(element.first("w:u")),
                isItalic: readBooleanElement(element.first("w:i")),
                isStrikethrough: readBooleanElement(element.first("w:strike")),
                isAllCaps: readBooleanElement(element.first("w:caps")),
                isSmallCaps: readBooleanElement(element.first("w:smallCaps")),
                highlight: readHighlightValue(element.firstOrEmpty("w:highlight").attributes["w:val"]),
                shading: element.firstOrEmpty("w:shd").attributes["w:fill"]
            };
        });
    }

    function readUnderline(element) {
        if (element) {
            var value = element.attributes["w:val"];
            return value !== undefined && value !== "false" && value !== "0" && value !== "none";
        } else {
            return false;
        }
    }

    function readBooleanElement(element) {
        if (element) {
            var value = element.attributes["w:val"];
            return value !== "false" && value !== "0";
        } else {
            return false;
        }
    }

    function readBooleanAttributeValue(value) {
        return value !== "false" && value !== "0";
    }

    function readHighlightValue(value) {
        if (!value || value === "none") {
            return null;
        } else {
            return value;
        }
    }

    function readParagraphStyle(element) {
        return readStyle(element, "w:pStyle", "Paragraph", styles.findParagraphStyleById);
    }

    function readRunStyle(element) {
        return readStyle(element, "w:rStyle", "Run", styles.findCharacterStyleById);
    }

    function readTableStyle(element) {
        return readStyle(element, "w:tblStyle", "Table", styles.findTableStyleById);
    }

    function readStyle(element, styleTagName, styleType, findStyleById) {
        var messages = [];
        var styleElement = element.first(styleTagName);
        var styleId = null;
        var name = null;
        if (styleElement) {
            styleId = styleElement.attributes["w:val"];
            if (styleId) {
                var style = findStyleById(styleId);
                if (style) {
                    name = style.name;
                } else {
                    messages.push(undefinedStyleWarning(styleType, styleId));
                }
            }
        }
        return elementResultWithMessages({styleId: styleId, name: name}, messages);
    }
    
    function extractShapeDimensions(style) {
        var dimensions = {
            width: null,
            height: null
        };
    
        if (!style) {
            return dimensions;
        }
    
        var dimensionPattern = /(?:width|height):([\d.\-]+)(pt|px|cm|in|dxa|emu)?/g;
        var match;
  
        while ((match = dimensionPattern.exec(style)) !== null) {
            var rawValue = parseFloat(match[1]);
            var unit = match[2] || inferUnitFromRawValue(rawValue);
            var property = match[0].startsWith('width') ? 'width' : 'height';
    
            dimensions[property] = convertToPixels(rawValue, unit);
        }
  
        return dimensions;
    }

    function inferUnitFromRawValue(rawValue) {
        if (rawValue > 10000) {
            return 'EMU';
        } else if (rawValue > 1000) {
            return 'DXA';
        } else {
            return 'pt';
        }
    }

    function readShapeElement(element) {
        var style = element.attributes["style"] || "";
        var imageDataElement = element.first("v:imagedata");
        if (imageDataElement) {
            var dimensions = extractShapeDimensions(style);
            return readImageData(imageDataElement, dimensions);
        }
        return readChildElements(element);
    }

    function readFldChar(element) {
        var type = element.attributes["w:fldCharType"];
        if (type === "begin") {
            complexFieldStack.push({type: "begin", fldChar: element});
            currentInstrText = [];
        } else if (type === "end") {
            var complexFieldEnd = complexFieldStack.pop();
            if (complexFieldEnd.type === "begin") {
                complexFieldEnd = parseCurrentInstrText(complexFieldEnd);
            }
            if (complexFieldEnd.type === "checkbox") {
                return elementResult(documents.checkbox({
                    checked: complexFieldEnd.checked
                }));
            }
        } else if (type === "separate") {
            var complexFieldSeparate = complexFieldStack.pop();
            var complexField = parseCurrentInstrText(complexFieldSeparate);

            complexFieldStack.push(complexField);
        }
        return emptyResult();
    }

    function currentHyperlinkOptions() {
        var topHyperlink = _.last(complexFieldStack.filter(function(complexField) {
            return complexField.type === "hyperlink";
        }));
        return topHyperlink ? topHyperlink.options : null;
    }

    function parseCurrentInstrText(complexField) {
        return parseInstrText(
            currentInstrText.join(''),
            complexField.type === "begin"
                ? complexField.fldChar
                : xml.emptyElement
        );
    }

    function parseInstrText(instrText, fldChar) {
        var externalLinkResult = /\s*HYPERLINK "(.*)"/.exec(instrText);
        if (externalLinkResult) {
            return {type: "hyperlink", options: {href: externalLinkResult[1]}};
        }

        var internalLinkResult = /\s*HYPERLINK\s+\\l\s+"(.*)"/.exec(instrText);
        if (internalLinkResult) {
            return {type: "hyperlink", options: {anchor: internalLinkResult[1]}};
        }

        var checkboxResult = /\s*FORMCHECKBOX\s*/.exec(instrText);
        if (checkboxResult) {
            var checkboxElement = fldChar
                .firstOrEmpty("w:ffData")
                .firstOrEmpty("w:checkBox");
            var checkedElement = checkboxElement.first("w:checked");
            var checked = checkedElement == null
                ? readBooleanElement(checkboxElement.first("w:default"))
                : readBooleanElement(checkedElement);
            return {type: "checkbox", checked: checked};
        }

        return {type: "unknown"};
    }

    function readInstrText(element) {
        currentInstrText.push(element.text());
        return emptyResult();
    }

    function readSymbol(element) {
        // See 17.3.3.30 sym (Symbol Character) of ECMA-376 4th edition Part 1
        var font = element.attributes["w:font"];
        var char = element.attributes["w:char"];
        var unicodeCharacter = dingbatToUnicode.hex(font, char);
        if (unicodeCharacter == null && /^F0..$/.test(char)) {
            unicodeCharacter = dingbatToUnicode.hex(font, char.substring(2));
        }

        if (unicodeCharacter == null) {
            return emptyResultWithMessages([warning(
                "A w:sym element with an unsupported character was ignored: char " +  char + " in font " + font
            )]);
        } else {
            return elementResult(new documents.Text(unicodeCharacter.string));
        }
    }

    function noteReferenceReader(noteType) {
        return function(element) {
            var noteId = element.attributes["w:id"];
            return elementResult(new documents.NoteReference({
                noteType: noteType,
                noteId: noteId
            }));
        };
    }

    function readCommentReference(element) {
        return elementResult(documents.commentReference({
            commentId: element.attributes["w:id"]
        }));
    }

    function readCommentRangeStart(element) {
        return elementResult(documents.commentRangeStart({
            commentId: element.attributes["w:id"]
        }));
    }

    function readCommentRangeEnd(element) {
        return elementResult(documents.commentRangeEnd({
            commentId: element.attributes["w:id"]
        }));
    }

    function readChildElements(element) {
        return readXmlElements(element.children);
    }

    var xmlElementReaders = {
        "w:p": function(element) {
            var paragraphPropertiesElement = element.firstOrEmpty("w:pPr");

            var isDeleted = !!paragraphPropertiesElement
                .firstOrEmpty("w:rPr")
                .first("w:del");

            if (isDeleted) {
                element.children.forEach(function(child) {
                    deletedParagraphContents.push(child);
                });
                return emptyResult();
            } else {
                var childrenXml = element.children;
                if (deletedParagraphContents.length > 0) {
                    childrenXml = deletedParagraphContents.concat(childrenXml);
                    deletedParagraphContents = [];
                }
                return ReadResult.map(
                    readParagraphProperties(paragraphPropertiesElement),
                    readXmlElements(childrenXml),
                    function(properties, children) {
                        return new documents.Paragraph(children, properties);
                    }
                ).insertExtra();
            }
        },
        "w:r": function(element) {
            return ReadResult.map(
                readRunProperties(element.firstOrEmpty("w:rPr")),
                readXmlElements(element.children),
                function(properties, children) {
                    var hyperlinkOptions = currentHyperlinkOptions();
                    if (hyperlinkOptions !== null) {
                        children = [new documents.Hyperlink(children, hyperlinkOptions)];
                    }

                    return new documents.Run(children, properties);
                }
            );
        },
        "w:fldChar": readFldChar,
        "w:instrText": readInstrText,
        "w:t": function(element) {
            return elementResult(new documents.Text(element.text()));
        },
        "w:delText": function(element) {
            return elementResult(new documents.Text(element.text()));
        },
        "w:tab": function(element) {
            return elementResult(new documents.Tab());
        },
        "w:noBreakHyphen": function() {
            return elementResult(new documents.Text("\u2011"));
        },
        "w:softHyphen": function(element) {
            return elementResult(new documents.Text("\u00AD"));
        },
        "w:sym": readSymbol,
        "w:hyperlink": function(element) {
            var relationshipId = element.attributes["r:id"];
            var anchor = element.attributes["w:anchor"];
            return readXmlElements(element.children).map(function(children) {
                function create(options) {
                    var targetFrame = element.attributes["w:tgtFrame"] || null;

                    return new documents.Hyperlink(
                        children,
                        _.extend({targetFrame: targetFrame}, options)
                    );
                }

                if (relationshipId) {
                    var href = relationships.findTargetByRelationshipId(relationshipId);
                    if (anchor) {
                        href = uris.replaceFragment(href, anchor);
                    }
                    return create({href: href});
                } else if (anchor) {
                    return create({anchor: anchor});
                } else {
                    return children;
                }
            });
        },
        "w:tbl": readTable,
        "w:tr": readTableRow,
        "w:tc": readTableCell,
        "w:footnoteReference": noteReferenceReader("footnote"),
        "w:endnoteReference": noteReferenceReader("endnote"),
        "w:commentReference": readCommentReference,
        "w:commentRangeStart": readCommentRangeStart,
        "w:commentRangeEnd": readCommentRangeEnd,
        "w:br": function(element) {
            var breakType = element.attributes["w:type"];
            if (breakType == null || breakType === "textWrapping") {
                return elementResult(documents.lineBreak);
            } else if (breakType === "page") {
                return elementResult(documents.pageBreak);
            } else if (breakType === "column") {
                return elementResult(documents.columnBreak);
            } else {
                return emptyResultWithMessages([warning("Unsupported break type: " + breakType)]);
            }
        },
        "w:bookmarkStart": function(element){
            var name = element.attributes["w:name"];
            if (name === "_GoBack") {
                return emptyResult();
            } else {
                return elementResult(new documents.BookmarkStart({name: name}));
            }
        },

        "mc:AlternateContent": function(element) {
            return readChildElements(element.firstOrEmpty("mc:Fallback"));
        },

        "w:sdt": function(element) {
            var contentResult = readXmlElements(element.firstOrEmpty("w:sdtContent").children);
            return contentResult.map(function(content) {
                // From the WordML standard: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/3350cb64-931f-41f7-8824-f18b2568ce66
                //
                // > A CT_SdtCheckbox element that specifies that the parent
                // > structured document tag is a checkbox when displayed in the
                // > document. The parent structured document tag contents MUST
                // > contain a single character and optionally an additional
                // > character in a deleted run.

                var checkbox = element
                    .firstOrEmpty("w:sdtPr")
                    .first("wordml:checkbox");

                if (checkbox) {
                    var checkedElement = checkbox.first("wordml:checked");
                    var isChecked = !!checkedElement && readBooleanAttributeValue(
                        checkedElement.attributes["wordml:val"]
                    );
                    var documentCheckbox = documents.checkbox({
                        checked: isChecked
                    });

                    var hasCheckbox = false;
                    var replacedContent = content.map(transforms._elementsOfType(
                        documents.types.text,
                        function(text) {
                            if (text.value.length > 0 && !hasCheckbox) {
                                hasCheckbox = true;
                                return documentCheckbox;
                            } else {
                                return text;
                            }
                        }
                    ));

                    if (hasCheckbox) {
                        return replacedContent;
                    } else {
                        return documentCheckbox;
                    }

                } else {
                    return content;
                }
            });
        },

        "w:ins": function(element) {
            insertions.push(element);
            var attributes = {};
            if (element.attributes["w:author"]) {
                attributes = Object.assign({}, {authorName: element.attributes["w:author"]}, attributes);
            }
            if (element.attributes["w:date"]) {
                attributes = Object.assign({}, {date: element.attributes["w:date"]}, attributes);
            }
            if (element.attributes["w:id"]) {
                attributes = Object.assign({}, {changeId: element.attributes["w:id"]}, attributes);
            }
            if (element.children && element.children.length > 0) {
                var propertiesHolder = element.children[0].first("w:rPr");
                if (propertiesHolder) {
                    var attributesHolder = propertiesHolder.first("w:rPrChange");
                    if (attributesHolder) {
                        var childAttributes = {
                            authorName: attributesHolder.attributes["w:author"],
                            date: attributesHolder.attributes["w:date"],
                            changeId: attributesHolder.attributes["w:id"]
                        };
                        attributes = Object.assign({}, childAttributes, attributes);
                    }
                }
            }
            return ReadResult.map(
                readRunProperties(element.firstOrEmpty("w:rPr")),
                readXmlElements(element.children),
                function(properties, children) {
                    var insProperties = Object.assign({}, properties, attributes);
                    children = [new documents.Ins(children, insProperties)];
                    return new documents.Run(children, properties);
                }
            );
        },
        "w:del": function(element) {
            deletions.push(element);
            var attributes = {};
            if (element.attributes["w:author"]) {
                attributes = Object.assign({}, {authorName: element.attributes["w:author"]}, attributes);
            }
            if (element.attributes["w:date"]) {
                attributes = Object.assign({}, {date: element.attributes["w:date"]}, attributes);
            }
            if (element.attributes["w:id"]) {
                attributes = Object.assign({}, {changeId: element.attributes["w:id"]}, attributes);
            }
            if (element.children && element.children.length > 0) {
                var propertiesHolder = element.children[0].first("w:rPr");
                if (propertiesHolder) {
                    var attributesHolder = propertiesHolder.first("w:rPrChange");
                    if (attributesHolder) {
                        var childAttributes = {
                            authorName: attributesHolder.attributes["w:author"],
                            date: attributesHolder.attributes["w:date"],
                            changeId: attributesHolder.attributes["w:id"]
                        };
                        attributes = Object.assign({}, childAttributes, attributes);
                    }
                }
            }
            return ReadResult.map(
                readRunProperties(element.firstOrEmpty("w:rPr")),
                readXmlElements(element.children),
                function(properties, children) {
                    var delProperties = Object.assign({}, properties, attributes);
                    children = [new documents.Del(children, delProperties)];
                    return new documents.Run(children, properties);
                }
            );
        },
        "w:object": readChildElements,
        "w:smartTag": readChildElements,
        "w:drawing": readChildElements,
        "w:pict": function(element) {
            return readChildElements(element).toExtra();
        },
        "v:roundrect": readChildElements,
        "v:shape": readShapeElement,
        "v:textbox": readChildElements,
        "w:txbxContent": readChildElements,
        "wp:inline": readDrawingElement,
        "wp:anchor": readDrawingElement,
        "v:imagedata": readImageData,
        "v:group": readChildElements,
        "v:rect": readChildElements
    };

    return {
        readXmlElement: readXmlElement,
        readXmlElements: readXmlElements
    };


    function readTable(element) {
        var propertiesResult = readTableProperties(element.firstOrEmpty("w:tblPr"));
        return readXmlElements(element.children)
            .flatMap(calculateRowSpans)
            .flatMap(function(children) {
                return propertiesResult.map(function(properties) {

                    // Check if any cell has borders to enhance isBordered detection
                    var hasAnyCellBorders = checkForCellBorders(children);
                    
                    // If table doesn't have borders but cells do, mark table as bordered
                    if (!properties.isBordered && hasAnyCellBorders) {
                        properties.isBordered = true;
                    }
                    
                    return documents.Table(children, properties);
                });
            });
    }

    // Helper function to check if any cell in the table has borders
    function checkForCellBorders(children) {
        for (var i = 0; i < children.length; i++) {
            var child = children[i];
            if (child.type === 'tableRow' && child.children) {
                for (var j = 0; j < child.children.length; j++) {
                    var cell = child.children[j];
                    if (cell.type === 'tableCell' && cell.borders) {
                        return true;
                    }
                }
            }
        }
        return false;
    }

    function readTableProperties(element) {
        var returnValue = Object.assign({});
        var borderOptions = element.firstOrEmpty("w:tblBorders");
        
        // Read detailed table borders
        var tableBorders = readDetailedBorders(borderOptions);
        returnValue.borders = tableBorders;
        
        // Keep existing isBordered logic for backward compatibility, enhanced with border detection
        if (borderOptions && borderOptions.children && borderOptions.children.length > 0) {
            readTableBorders(borderOptions.children[0]).map(function(borders) {
                returnValue.isBordered = borders.isBordered;
            });
        } else {
            returnValue.isBordered = false;
        }

        // If we found detailed borders but isBordered is false, set it to true
        // But we have to skip the insideV and insideH borders
        if ((tableBorders && (tableBorders['top'] || tableBorders['right'] || tableBorders['bottom'] || tableBorders['left'])) && !returnValue.isBordered) {
            returnValue.isBordered = true;
        }
        
        return readTableStyle(element).map(function(style) {
            return {
                styleId: style.styleId,
                styleName: style.name,
                isBordered: returnValue.isBordered,
                borders: returnValue.borders
            };
        });
    }

    function readTableBorders(element) {
        var isBordered = false;
        if (element.attributes.length && element.attributes["w:val"] !== "nil" && element.attributes["w:val"] !== "none") {
            isBordered = true;
        }
        return elementResult({isBordered: isBordered});
    }

    // Read detailed border information from a borders container element
    function readDetailedBorders(bordersElement) {
        if (!bordersElement || !bordersElement.children || bordersElement.children.length === 0) {
            return null;
        }

        var borders = {
            top: null,
            right: null,
            bottom: null,
            left: null,
            insideV: null,
            insideH: null
        };

        bordersElement.children.forEach(function(borderElement) {
            if (!borderElement || borderElement.type !== 'element') {
                return;
            }
            
            var side = borderElement.name ? borderElement.name.replace('w:', '') : '';
            var border = readSingleBorder(borderElement);
            
            if (side && (side === 'top' || side === 'bottom' || side === 'left' || side === 'right' || side === 'insideV' || side === 'insideH')) {
                borders[side] = border;
            }
        });

        // Return null if no borders are found, otherwise return the borders object
        var hasBorders = borders.top || borders.right || borders.bottom || borders.left || borders.insideV || borders.insideH;

        return hasBorders ? borders : null;
    }

    // Read a single border element and extract its properties
    function readSingleBorder(borderElement) {
        if (!borderElement || !borderElement.attributes) {
            return null;
        }

        var val = borderElement.attributes["w:val"];
        
        // Skip nil and none borders
        if (!val || val === "nil" || val === "none") {
            return null;
        }

        return {
            style: val,
            width: parseFloat(borderElement.attributes["w:sz"]) || 0,
            color: borderElement.attributes["w:color"] || "000000",
            space: parseFloat(borderElement.attributes["w:space"]) || 0
        };
    }

    function readTableRow(element) {
        var properties = element.firstOrEmpty("w:trPr");

        // See 17.13.5.12 del (Deleted Table Row) of ECMA-376 4th edition Part 1
        var isDeleted = !!properties.first("w:del");
        if (isDeleted) {
            return emptyResult();
        }

        var isHeader = !!properties.first("w:tblHeader");
        return readXmlElements(element.children).map(function(children) {
            return documents.TableRow(children, {isHeader: isHeader});
        });
    }

    function readTableCell(element) {
        return readXmlElements(element.children).map(function(children) {
            var properties = element.firstOrEmpty("w:tcPr");

            var gridSpan = properties.firstOrEmpty("w:gridSpan").attributes["w:val"];
            var colSpan = gridSpan ? parseInt(gridSpan, 10) : 1;

            // Diving by 10 to get closer to the pixel size we want
            var widthProp = properties.firstOrEmpty("w:tcW").attributes["w:w"];
            var width = null;
            if (widthProp) {
                // Convert dxa to pt, then use 20/12 px/pt to convert to px
                width = widthProp / 12;
            }

            // Background color
            var cellColor = properties.firstOrEmpty("w:shd").attributes["w:fill"];

            // Read detailed cell borders
            var cellBorders = readDetailedBorders(properties.firstOrEmpty("w:tcBorders"));

            console.log(cellBorders);

            var cell = documents.TableCell(children, {
                colSpan: colSpan,
                width: width,
                bgColor: cellColor,
                borders: cellBorders
            });
            cell._vMerge = readVMerge(properties);
            return cell;
        });
    }

    function readVMerge(properties) {
        var element = properties.first("w:vMerge");
        if (element) {
            var val = element.attributes["w:val"];
            return val === "continue" || !val;
        } else {
            return null;
        }
    }

    function calculateRowSpans(rows) {
        var unexpectedNonRows = _.any(rows, function(row) {
            return row.type !== documents.types.tableRow;
        });
        if (unexpectedNonRows) {
            return elementResultWithMessages(rows, [warning(
                "unexpected non-row element in table, cell merging may be incorrect"
            )]);
        }
        var unexpectedNonCells = _.any(rows, function(row) {
            return _.any(row.children, function(cell) {
                return cell.type !== documents.types.tableCell;
            });
        });
        if (unexpectedNonCells) {
            return elementResultWithMessages(rows, [warning(
                "unexpected non-cell element in table row, cell merging may be incorrect"
            )]);
        }

        var columns = {};

        rows.forEach(function(row) {
            var cellIndex = 0;
            row.children.forEach(function(cell) {
                if (cell._vMerge && columns[cellIndex]) {
                    columns[cellIndex].rowSpan++;
                } else {
                    columns[cellIndex] = cell;
                    cell._vMerge = false;
                }
                cellIndex += cell.colSpan;
            });
        });

        rows.forEach(function(row) {
            row.children = row.children.filter(function(cell) {
                return !cell._vMerge;
            });
            row.children.forEach(function(cell) {
                delete cell._vMerge;
            });
        });

        return elementResult(rows);
    }

    function readDrawingElement(element) {
        var blipFills = element
            .getElementsByTagName("a:graphic")
            .getElementsByTagName("a:graphicData")
            .getElementsByTagName("pic:pic")
            .getElementsByTagName("pic:blipFill");

        var results = [];
        blipFills.forEach(function(blipFill) {
            var blips = blipFill.getElementsByTagName("a:blip");
            blips.forEach(function(blip) {
                results.push(readBlip(element, blip, blipFill));
            });
        });

        return combineResults(results);
    }

    function getUnitType(element) {
        var namespace = element.name.split(":")[0];
        if (namespace === "wp") {
            // `wp:extent` is defined to use EMUs in Word XML
            return "EMU";
        } else if (namespace === "a") {
            // Assume EMUs for `a:ext` as well
            return "EMU";
        } else {
            // Fallback to assume dxa for unknown contexts
            return "DXA";
        }
    }

    function convertToPixels(rawValue, unitType) {
        var pixels;
  
        switch (unitType) {
        case "EMU":
            pixels = rawValue * (96 / 914400);
            break;
        case "DXA":
                // was 12, changed to 20
            pixels = (rawValue / 20) * (96 / 72);
            break;
        case "pt":
            pixels = rawValue * (96 / 72);
            break;
        case "cm":
            pixels = rawValue * (96 / 2.54);
            break;
        case "in":
            pixels = rawValue * 96;
            break;
        case "px":
        default:
            pixels = rawValue;
            break;
        }
      
        return pixels;
      /* return {
          originalValue: rawValue,
          originalUnit: unitType,
          pixels: pixels
      }; */
    }

    function readBlip(element, blip, blipFill) {
        var properties = element.first("wp:docPr").attributes;
        var altText = isBlank(properties.descr) ? properties.title : properties.descr;
        var blipImageFile = findBlipImageFile(blip);
        var dimensionsHolder = element.firstOrEmpty("wp:extent");

        if (!dimensionsHolder) {
            return emptyResultWithMessages([warning("Missing dimensions for the image")]);
        }

        var dimensionsAttributes = dimensionsHolder.attributes;
        var widthRaw = dimensionsAttributes.cx;
        var heightRaw = dimensionsAttributes.cy;

      // Detect the unit type (EMU or DXA)
        var unitType = getUnitType(element);

      // Convert dimensions to pixels based on the detected unit
        var width = widthRaw ? convertToPixels(widthRaw, unitType) : null;
        var height = heightRaw ? convertToPixels(heightRaw, unitType) : null;

      // Get further formatting options
        var imageProperties = parseImageProperties(element);

        // Extract cropping information from a:srcRect in blipFill
        var srcRect = blipFill ? blipFill.firstOrEmpty("a:srcRect") : null;
        
        var cropping = null;
        var isCropped = false;

        if (srcRect && srcRect.attributes) {
            var attrs = srcRect.attributes;
            cropping = {
                left: attrs.l ? parseInt(attrs.l, 10) : 0,
                top: attrs.t ? parseInt(attrs.t, 10) : 0,
                right: attrs.r ? parseInt(attrs.r, 10) : 0,
                bottom: attrs.b ? parseInt(attrs.b, 10) : 0
            };
            
            // Check if any cropping is applied (values are in percentage * 1000, so > 0 means cropped)
            isCropped = cropping.left > 0 || cropping.top > 0 || cropping.right > 0 || cropping.bottom > 0;
        } else {
            // Set default cropping values
            cropping = {
                left: 0,
                top: 0,
                right: 0,
                bottom: 0
            };
        }

        if (blipImageFile === null) {
            return emptyResultWithMessages([warning("Could not find image file for a:blip element")]);
        } else {
            return readImage(blipImageFile, altText, width, height, imageProperties, cropping, isCropped);
        }
    }

    function isBlank(value) {
        return value == null || /^\s*$/.test(value);
    }

    function findBlipImageFile(blip) {
        var embedRelationshipId = blip.attributes["r:embed"];
        var linkRelationshipId = blip.attributes["r:link"];
        if (embedRelationshipId) {
            return findEmbeddedImageFile(embedRelationshipId);
        } else if (linkRelationshipId) {
            var imagePath = relationships.findTargetByRelationshipId(linkRelationshipId);
            return {
                path: imagePath,
                read: files.read.bind(files, imagePath)
            };
        } else {
            return null;
        }
    }

    function readImageData(element, dimensions) {
        var dimensionsObject = dimensions || {};
        var relationshipId = element.attributes['r:id'];
        if (relationshipId) {
            return readImage(
                findEmbeddedImageFile(relationshipId),
                element.attributes["o:title"],
                dimensionsObject.width,
                dimensionsObject.height,
                null, // imageProperties
                null, // srcRect
                false // isCropped
            );
        } else {
            return emptyResultWithMessages([warning("A v:imagedata element without a relationship ID was ignored")]);
        }
    }

    function findEmbeddedImageFile(relationshipId) {
        var path = uris.uriToZipEntryName("word", relationships.findTargetByRelationshipId(relationshipId));
        return {
            path: path,
            read: docxFile.read.bind(docxFile, path)
        };
    }

    function parseImageProperties(element) {
        var floating = null;
        var wrappingStyle = null;
        var alignmentH = null;
        var alignmentV = null;
        var positionOffsetH = null;
        var positionOffsetV = null;
  
      // Detect floating and wrapping style
        if (element.name === "wp:inline") {
            floating = "inline";
        } else if (element.name === "wp:anchor") {
            floating = "floating";
  
            // Detect wrapping style
            var wrapElement = null;
            for (var i = 0; i < element.children.length; i++) {
                if (element.children[i].name.indexOf("wp:wrap") === 0) {
                    wrapElement = element.children[i];
                    break;
                }
            }
            wrappingStyle = wrapElement ? wrapElement.name.replace("wp:", "") : null;
        }
  
      // Detect alignment
        var positionH = null;
        var positionV = null;
        for (var j = 0; j < element.children.length; j++) {
            if (element.children[j].name === "wp:positionH") {
                positionH = element.children[j];
            } else if (element.children[j].name === "wp:positionV") {
                positionV = element.children[j];
            }
        }
  
        if (positionH) {
            alignmentH = positionH.attributes.relativeFrom;
            var posOffsetH = null;
            for (var k = 0; k < positionH.children.length; k++) {
                if (positionH.children[k].name === "wp:posOffset") {
                    posOffsetH = positionH.children[k];
                    break;
                }
            }
            positionOffsetH = posOffsetH ? parseInt(posOffsetH.children[0].value, 10) : null;
        }
  
        if (positionV) {
            alignmentV = positionV.attributes.relativeFrom;
            var posOffsetV = null;
            for (var l = 0; l < positionV.children.length; l++) {
                if (positionV.children[l].name === "wp:posOffset") {
                    posOffsetV = positionV.children[l];
                    break;
                }
            }
            positionOffsetV = posOffsetV ? parseInt(posOffsetV.children[0].value, 10) : null;
        }
    
        return {
            floating: floating,
            wrappingStyle: wrappingStyle,
            alignmentH: alignmentH,
            alignmentV: alignmentV,
            positionOffsetH: positionOffsetH,
            positionOffsetV: positionOffsetV
        };
    }

    function readImage(imageFile, altText, width, height, imageProperties, srcRect, isCropped) {
        var contentType = contentTypes.findContentType(imageFile.path);

        var image = documents.Image({
            readImage: imageFile.read,
            altText: altText,
            contentType: contentType,
            width: width,
            height: height,
            imageProperties: imageProperties,
            srcRect: srcRect,
            isCropped: isCropped
        });
        var warnings = supportedImageTypes[contentType] ?
            [] : warning("Image of type " + contentType + " is unlikely to display in web browsers");
        return elementResultWithMessages(image, warnings);
    }

    function undefinedStyleWarning(type, styleId) {
        return warning(
            type + " style with ID " + styleId + " was referenced but not defined in the document");
    }
}


function readNumberingProperties(styleId, element, numbering) {
    var level = element.firstOrEmpty("w:ilvl").attributes["w:val"];
    var numId = element.firstOrEmpty("w:numId").attributes["w:val"];
    if (level !== undefined && numId !== undefined) {
        var levelInfo = numbering.findLevel(numId, level);
        if (levelInfo) {
            var returnObject = Object.assign({}, levelInfo, {
                numId: numId,
                // Add visibility information for bullet lists
                isVisibleBulletList: numbering.isVisibleBulletList && numbering.isVisibleBulletList(numId, level),
                bulletCharacter: numbering.getBulletCharacter && numbering.getBulletCharacter(numId, level)
            });
            return returnObject;
        }
        return null;
    }

    if (styleId != null) {
        var levelByStyleId = numbering.findLevelByParagraphStyleId(styleId);
        if (levelByStyleId != null) {
            // For style-based numbering, we can't easily determine visibility without numId
            // but we can still provide the basic information
            return Object.assign({}, levelByStyleId, {
                isVisibleBulletList: levelByStyleId.isBullet && !levelByStyleId.isTentative,
                bulletCharacter: levelByStyleId.isBullet ? (levelByStyleId.delimiter || "•") : null
            });
        }
    }

    return null;
}

function readSpacingProperties(element) {
    var properties = element.firstOrEmpty("w:spacing").attributes;
    return {
        after: properties["w:after"],
        before: properties["w:before"],
        line: properties["w:line"],
        lineRule: properties["w:lineRule"]
    };
}

var supportedImageTypes = {
    "image/png": true,
    "image/gif": true,
    "image/jpeg": true,
    "image/svg+xml": true,
    "image/tiff": true
};

var ignoreElements = {
    "office-word:wrap": true,
    "v:shadow": true,
    "v:shapetype": true,
    "w:annotationRef": true,
    "w:bookmarkEnd": true,
    "w:sectPr": true,
    "w:proofErr": true,
    "w:lastRenderedPageBreak": true,
    // "w:commentRangeStart": true,
    // "w:commentRangeEnd": true,
    "w:del": true,
    "w:footnoteRef": true,
    "w:endnoteRef": true,
    "w:pPr": true,
    "w:rPr": true,
    "w:tblPr": true,
    "w:tblGrid": true,
    "w:trPr": true,
    "w:tcPr": true
};

function emptyResultWithMessages(messages) {
    return new ReadResult(null, null, messages);
}

function emptyResult() {
    return new ReadResult(null);
}

function elementResult(element) {
    return new ReadResult(element);
}

function elementResultWithMessages(element, messages) {
    return new ReadResult(element, null, messages);
}

function ReadResult(element, extra, messages) {
    this.value = element || [];
    this.extra = extra || [];
    this._result = new Result({
        element: this.value,
        extra: extra
    }, messages);
    this.messages = this._result.messages;
}

ReadResult.prototype.toExtra = function() {
    return new ReadResult(null, joinElements(this.extra, this.value), this.messages);
};

ReadResult.prototype.insertExtra = function() {
    var extra = this.extra;
    if (extra && extra.length) {
        return new ReadResult(joinElements(this.value, extra), null, this.messages);
    } else {
        return this;
    }
};

ReadResult.prototype.map = function(func) {
    var result = this._result.map(function(value) {
        return func(value.element);
    });
    return new ReadResult(result.value, this.extra, result.messages);
};

ReadResult.prototype.flatMap = function(func) {
    var result = this._result.flatMap(function(value) {
        return func(value.element)._result;
    });
    return new ReadResult(result.value.element, joinElements(this.extra, result.value.extra), result.messages);
};

ReadResult.map = function(first, second, func) {
    return new ReadResult(
        func(first.value, second.value),
        joinElements(first.extra, second.extra),
        first.messages.concat(second.messages)
    );
};

function combineResults(results) {
    var result = Result.combine(_.pluck(results, "_result"));
    return new ReadResult(
        _.flatten(_.pluck(result.value, "element")),
        _.filter(_.flatten(_.pluck(result.value, "extra")), identity),
        result.messages
    );
}

function joinElements(first, second) {
    return _.flatten([first, second]);
}

function identity(value) {
    return value;
}
