exports.readStylesXml = readStylesXml;
exports.Styles = Styles;
exports.defaultStyles = new Styles({}, {});

function Styles(paragraphStyles, characterStyles, tableStyles, numberingStyles, customStyles) {
    return {
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
        },
        getCustomStyles: function() {
            return customStyles;
        }
    };
}

Styles.EMPTY = new Styles({}, {}, {}, {}, {});

function readStylesXml(root) {
    var paragraphStyles = {};
    var characterStyles = {};
    var tableStyles = {};
    var numberingStyles = {};
    var customStyles = {};

    var styles = {
        "paragraph": paragraphStyles,
        "character": characterStyles,
        "table": tableStyles,
    };

    root.getElementsByTagName("w:style").forEach(function(styleElement) {
        var style = readStyleElement(styleElement);
        style.element = styleElement; // Store the raw XML element for inheritance purposes

        if (style.type === "numbering") {
            numberingStyles[style.styleId] = readNumberingStyleElement(styleElement);
        } else {
            var styleSet = styles[style.type];
            if (styleSet) {
                styleSet[style.styleId] = style;
            }
        }

        if (isCustomStyle(styleElement)) {
            customStyles[style.styleId] = {
                name: style.name,
                type: style.type,
                properties: extractStyleProperties(styleElement, {
                    paragraph: paragraphStyles,
                    character: characterStyles,
                    table: tableStyles,
                    numbering: numberingStyles,
                })
            };
        }
    });

    var returnObject = new Styles(paragraphStyles, characterStyles, tableStyles, numberingStyles, customStyles);

    return returnObject;
}

function readStyleElement(styleElement) {
    var type = styleElement.attributes["w:type"];
    var styleId = styleElement.attributes["w:styleId"];
    var name = styleName(styleElement);
    return { type: type, styleId: styleId, name: name };
}

function styleName(styleElement) {
    var nameElement = styleElement.first("w:name");
    return nameElement ? nameElement.attributes["w:val"] : null;
}

function readNumberingStyleElement(styleElement) {
  var numberingElement = styleElement.firstOrEmpty("w:pPr").firstOrEmpty("w:numPr");
    var numId = numberingElement
        .firstOrEmpty("w:numId")
        .attributes["w:val"];
        var level = numberingElement
        .firstOrEmpty("w:ilvl")
        .attributes["w:val"];

    return { 
      numId: numId,
      level: level,
    };
}

function isCustomStyle(styleElement) {
  return true;
  // Decided to return ALL styles as custom styles for better results
  // return !!styleElement.attributes["w:customStyle"];
}

// Helper to find style by styleId in the various style sets
function findStyleById(styleId, styles) {
  return (
      styles.paragraph[styleId] ||
      styles.character[styleId] ||
      styles.table[styleId] ||
      null
  );
}

// Helper function to extract properties of a style (e.g., font size, color, etc.)
function extractStyleProperties(styleElement, styles) {
  // Start with an empty properties object for each call
  var properties = {};

  // Clear properties that should not be inherited unless explicitly set
  properties.fontSize = undefined;
  properties.numbering = undefined;

  // Recursively fetch and merge base style properties if this style is based on another style
  var basedOnElement = styleElement.first("w:basedOn");
  if (basedOnElement) {
    var baseStyleId = basedOnElement.attributes["w:val"];
    var baseStyle = findStyleById(baseStyleId, styles);
    if (baseStyle && baseStyle.element) {
      var baseProperties = extractStyleProperties(baseStyle.element, styles);
      for (var key in baseProperties) {
        if (baseProperties.hasOwnProperty(key) && properties[key] === undefined) {
          properties[key] = baseProperties[key];
        }
      }
    }
  }

  // Read numbering style if present in the current style
  var numberingProperties = readNumberingStyleElement(styleElement);
  if (numberingProperties.numId || numberingProperties.level) {
    properties.numbering = numberingProperties;
  }

  // Extract the current style's properties
  var rPr = styleElement.first("w:rPr");
  var pPr = styleElement.first("w:pPr");

  // Run properties
  if (rPr) {
    var fontSizeString = rPr.firstOrEmpty("w:sz").attributes["w:val"];
    if (fontSizeString !== undefined) {
      properties.fontSize = /^[0-9]+$/.test(fontSizeString) ? parseInt(fontSizeString, 10) / 2 : undefined;
    }

    var colorElement = rPr.first("w:color");
    if (colorElement) {
      properties.color = colorElement.attributes["w:val"];
    }

    var underlineElement = rPr.first("w:u");
    properties.underline = underlineElement && underlineElement.attributes["w:val"] !== "none" && underlineElement.attributes["w:val"] !== "0" ? true : undefined;

    var boldElement = rPr.first("w:b");
    properties.bold = boldElement && boldElement.attributes["w:val"] !== "false" && boldElement.attributes["w:val"] !== "0" ? true : undefined;

    var italicElement = rPr.first("w:i");
    properties.italic = italicElement && italicElement.attributes["w:val"] !== "false" && italicElement.attributes["w:val"] !== "0" ? true : undefined;

    var strikeElement = rPr.first("w:strike");
    properties.strike = strikeElement && strikeElement.attributes["w:val"] !== "false" && strikeElement.attributes["w:val"] !== "0" ? true : undefined;

    var capsElement = rPr.first("w:caps");
    properties.isAllCaps = capsElement && capsElement.attributes["w:val"] !== "false" && capsElement.attributes["w:val"] !== "0" ? true : undefined;

    var smallCapsElement = rPr.first("w:smallCaps");
    properties.isSmallCaps = smallCapsElement && smallCapsElement.attributes["w:val"] !== "false" && smallCapsElement.attributes["w:val"] !== "0" ? true : undefined;

    var verticalAlignmentElement = rPr.first("w:vertAlign");
    if (verticalAlignmentElement) {
      properties.verticalAlignment = verticalAlignmentElement.attributes["w:val"];
    }

    var fontElement = rPr.first("w:rFonts");
    if (fontElement) {
      // Priority arguments should be taken:
      // w:ascii, w:eastAsia, w:hAnsi, w:cs
      properties.font = fontElement.attributes["w:ascii"] || fontElement.attributes["w:eastAsia"] || fontElement.attributes["w:hAnsi"] || fontElement.attributes["w:cs"];
    }

    var highlightElement = rPr.first("w:highlight");
    if (highlightElement) {
      properties.highlight = highlightElement.attributes["w:val"];
    }

    var shadingElement = rPr.first("w:shd");
    if (shadingElement) {
      properties.shading = shadingElement.attributes["w:fill"];
    }
  }

  // Paragraph properties
  if (pPr) {
    var spacingElement = pPr.first("w:spacing");
    if (spacingElement) {
      properties.spacing = {
        before: spacingElement.attributes["w:before"],
        after: spacingElement.attributes["w:after"],
        line: spacingElement.attributes["w:line"],
        lineRule: spacingElement.attributes["w:lineRule"]
      };
    }

    var alignmentElement = pPr.first("w:jc");
    if (alignmentElement) {
      properties.alignment = alignmentElement.attributes["w:val"];
    }

    var indElement = pPr.first("w:ind");
    if (indElement) {
      properties.indent = {
        left: indElement.attributes["w:left"],
        right: indElement.attributes["w:right"],
        start: indElement.attributes["w:start"],
        end: indElement.attributes["w:end"],
        hanging: indElement.attributes["w:hanging"],
        firstLine: indElement.attributes["w:firstLine"]
      };
    }
  }

  // Filter out undefined properties to ensure they don’t overwrite inherited values
  var cleanedProperties = {};
  for (var prop in properties) {
    if (properties[prop] !== undefined) {
      cleanedProperties[prop] = properties[prop];
    }
  }
  return cleanedProperties;
}

