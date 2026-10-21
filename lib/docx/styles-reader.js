exports.readStylesXml = readStylesXml;
exports.Styles = Styles;
exports.defaultStyles = new Styles({}, {});

function Styles(paragraphStyles, characterStyles, tableStyles, numberingStyles, customStyles) {
    return {
        getCustomStyles: function() {
            return customStyles || {};
        },
        findParagraphStyleById: function(styleId) {
            return paragraphStyles[styleId];
        },
        findCharacterStyleById: function(styleId) {
            return characterStyles[styleId];
        },
        findTableStyleById: function(styleId) {
            return tableStyles[styleId];
        },
        findNumberingStyleById: function(styleId) {
            return numberingStyles[styleId];
        }
    };
}

Styles.EMPTY = new Styles({}, {}, {}, {});

function readStylesXml(root) {
    var paragraphStyles = Object.create(null);
    var characterStyles = Object.create(null);
    var tableStyles = Object.create(null);
    var numberingStyles = Object.create(null);
    var styleElements = Object.create(null);

    root.getElementsByTagName("w:style").forEach(function(styleElement) {
        var style = readStyleElement(styleElement);
        var styleSet;

        switch (style.type) {
        case "paragraph":
            styleSet = paragraphStyles;
            break;

        case "character":
            styleSet = characterStyles;
            break;

        case "table":
            styleSet = tableStyles;
            break;

        case "numbering":
            styleSet = numberingStyles;
            break;
        }

        // Per 17.7.4.17 style (Style Definition) of ECMA-376 4th edition Part 1:
        //
        // > If multiple style definitions each declare the same value for their
        // > styleId, then the first such instance shall keep its current
        // > identifier with all other instances being reassigned in any manner
        // > desired.
        //
        // For the purpose of conversion, there's no point holding onto styles
        // with reassigned style IDs, so we ignore such style definitions.

        if (styleSet && styleSet[style.styleId] === undefined) {
            styleSet[style.styleId] = style;
        }

        if (styleElements[style.styleId] === undefined) {
            styleElements[style.styleId] = styleElement;
        }
    });

    var customStyles = Object.create(null);
    Object.keys(styleElements).forEach(function(styleId) {
        var styleElement = styleElements[styleId];
        customStyles[styleId] = {
            name: styleName(styleElement),
            type: styleElement.attributes["w:type"],
            properties: readStyleProperties(styleId, styleElements, Object.create(null))
        };
    });

    return new Styles(paragraphStyles, characterStyles, tableStyles, numberingStyles, customStyles);
}

// Resolves the formatting of a style, including formatting inherited through
// w:basedOn. A property set on the style itself replaces the inherited value.
function readStyleProperties(styleId, styleElements, seenStyleIds) {
    var styleElement = styleElements[styleId];
    if (!styleElement || seenStyleIds[styleId]) {
        return {};
    }
    seenStyleIds[styleId] = true;

    var properties = {};

    var basedOnElement = styleElement.first("w:basedOn");
    if (basedOnElement) {
        properties = readStyleProperties(basedOnElement.attributes["w:val"], styleElements, seenStyleIds);
    }

    var numPr = styleElement.firstOrEmpty("w:pPr").firstOrEmpty("w:numPr");
    var numId = numPr.firstOrEmpty("w:numId").attributes["w:val"];
    var level = numPr.firstOrEmpty("w:ilvl").attributes["w:val"];
    if (numId !== undefined || level !== undefined) {
        properties.numbering = {numId: numId, level: level};
    }

    var rPr = styleElement.first("w:rPr");
    if (rPr) {
        var fontSize = rPr.firstOrEmpty("w:sz").attributes["w:val"];
        if (fontSize !== undefined) {
            properties.fontSize = /^[0-9]+$/.test(fontSize) ? parseInt(fontSize, 10) / 2 : undefined;
        }

        var color = rPr.first("w:color");
        if (color) {
            properties.color = color.attributes["w:val"];
        }

        properties.underline = readUnderline(rPr.first("w:u"));
        properties.bold = readToggle(rPr.first("w:b"));
        properties.italic = readToggle(rPr.first("w:i"));
        properties.strike = readToggle(rPr.first("w:strike"));

        var fonts = rPr.first("w:rFonts");
        if (fonts) {
            properties.font = fonts.attributes["w:ascii"] || fonts.attributes["w:eastAsia"] ||
                fonts.attributes["w:hAnsi"] || fonts.attributes["w:cs"];
        }

        var highlight = rPr.first("w:highlight");
        if (highlight) {
            properties.highlight = highlight.attributes["w:val"];
        }
    }

    var pPr = styleElement.first("w:pPr");
    if (pPr) {
        var spacing = pPr.first("w:spacing");
        if (spacing) {
            properties.spacing = {
                before: spacing.attributes["w:before"],
                after: spacing.attributes["w:after"],
                line: spacing.attributes["w:line"],
                lineRule: spacing.attributes["w:lineRule"]
            };
        }

        var alignment = pPr.first("w:jc");
        if (alignment) {
            properties.alignment = alignment.attributes["w:val"];
        }

        var indent = pPr.first("w:ind");
        if (indent) {
            properties.indent = {
                left: indent.attributes["w:left"],
                right: indent.attributes["w:right"],
                start: indent.attributes["w:start"],
                end: indent.attributes["w:end"],
                hanging: indent.attributes["w:hanging"],
                firstLine: indent.attributes["w:firstLine"]
            };
        }
    }

    Object.keys(properties).forEach(function(key) {
        if (properties[key] === undefined) {
            delete properties[key];
        }
    });
    return properties;
}

function readToggle(element) {
    return element && element.attributes["w:val"] !== "false" && element.attributes["w:val"] !== "0" ? true : undefined;
}

function readUnderline(element) {
    return element && element.attributes["w:val"] !== "none" && element.attributes["w:val"] !== "0" ? true : undefined;
}

function readStyleElement(styleElement) {
    var type = styleElement.attributes["w:type"];

    if (type === "numbering") {
        return readNumberingStyleElement(type, styleElement);
    } else {
        var styleId = readStyleId(styleElement);
        var name = styleName(styleElement);
        return {type: type, styleId: styleId, name: name};
    }
}

function styleName(styleElement) {
    var nameElement = styleElement.first("w:name");
    return nameElement ? nameElement.attributes["w:val"] : null;
}

function readNumberingStyleElement(type, styleElement) {
    var styleId = readStyleId(styleElement);

    var numId = styleElement
        .firstOrEmpty("w:pPr")
        .firstOrEmpty("w:numPr")
        .firstOrEmpty("w:numId")
        .attributes["w:val"];

    return {type: type, numId: numId, styleId: styleId};
}

function readStyleId(styleElement) {
    return styleElement.attributes["w:styleId"];
}
