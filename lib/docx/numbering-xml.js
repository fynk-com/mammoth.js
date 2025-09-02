var _ = require("underscore");

exports.readNumberingXml = readNumberingXml;
exports.Numbering = Numbering;
exports.defaultNumbering = new Numbering({}, {});

function Numbering(nums, abstractNums, styles) {
    var allLevels = _.flatten(_.values(abstractNums).map(function(abstractNum) {
        return _.values(abstractNum.levels);
    }));

    var levelsByParagraphStyleId = _.indexBy(
        allLevels.filter(function(level) {
            return level.paragraphStyleId != null;
        }),
        "paragraphStyleId"
    );

    function findLevel(numId, level) {
        var num = nums[numId];
        if (num) {
            var abstractNum = abstractNums[num.abstractNumId];
            if (!abstractNum) {
                return null;
            } else if (abstractNum.numStyleLink == null) {
                var levelInfo = abstractNums[num.abstractNumId].levels[level];
                return levelInfo || null;
            } else {
                var style = styles.findNumberingStyleById(abstractNum.numStyleLink);
                return findLevel(style.numId, level);
            }
        } else {
            return null;
        }
    }

    function findLevelByParagraphStyleId(styleId) {
        return levelsByParagraphStyleId[styleId] || null;
    }

    function isVisibleBulletList(numId, level) {
        var levelInfo = findLevel(numId, level);
        if (!levelInfo) {
            return false;
        }
        
        // Check if it's a bullet format
        if (!levelInfo.isBullet) {
            return false;
        }
        
        // Check if it's tentative (usually means it's not visually displayed)
        if (levelInfo.isTentative) {
            return false;
        }
        
        // Check if it has proper bullet text or uses bullet fonts
        var hasValidBulletText = levelInfo.delimiter !== undefined && levelInfo.delimiter !== "";
        var hasBulletFont = levelInfo.bulletFont && (
            levelInfo.bulletFont.ascii === "Symbol" ||
            levelInfo.bulletFont.ascii === "Wingdings" ||
            levelInfo.bulletFont.hAnsi === "Symbol" ||
            levelInfo.bulletFont.hAnsi === "Wingdings"
        );
        
        // A visible bullet list should have either valid bullet text or proper bullet font
        return hasValidBulletText || hasBulletFont;
    }
    
    function getBulletCharacter(numId, level) {
        var levelInfo = findLevel(numId, level);
        if (!levelInfo || !levelInfo.isBullet) {
            return null;
        }
        
        // Return the bullet character if available
        if (levelInfo.delimiter) {
            return levelInfo.delimiter;
        }
        
        // Default bullet characters based on font
        if (levelInfo.bulletFont) {
            if (levelInfo.bulletFont.ascii === "Symbol" || levelInfo.bulletFont.hAnsi === "Symbol") {
                return "•"; // Standard bullet
            } else if (levelInfo.bulletFont.ascii === "Wingdings" || levelInfo.bulletFont.hAnsi === "Wingdings") {
                return ""; // Wingdings bullet
            }
        }
        
        return "•"; // Default bullet
    }

    return {
        findLevel: findLevel,
        findLevelByParagraphStyleId: findLevelByParagraphStyleId,
        isVisibleBulletList: isVisibleBulletList,
        getBulletCharacter: getBulletCharacter,
        getNumberingStyles: function() {
            return abstractNums;
        }
    };
}

function readNumberingXml(root, options) {
    if (!options || !options.styles) {
        throw new Error("styles is missing");
    }

    var abstractNums = readAbstractNums(root);
    var nums = readNums(root, abstractNums);
    return new Numbering(nums, abstractNums, options.styles);
}

function readAbstractNums(root) {
    var abstractNums = {};
    root.getElementsByTagName("w:abstractNum").forEach(function(element) {
        var id = element.attributes["w:abstractNumId"];
        abstractNums[id] = readAbstractNum(element);
    });
    return abstractNums;
}

function readAbstractNum(element) {
    var levels = {};
    element.getElementsByTagName("w:lvl").forEach(function(levelElement) {
        var levelIndex = levelElement.attributes["w:ilvl"];
        var numFmt = levelElement.firstOrEmpty("w:numFmt").attributes["w:val"];
        var lvlText = levelElement.firstOrEmpty("w:lvlText").attributes["w:val"];
        var paragraphStyleId = levelElement.firstOrEmpty("w:pStyle").attributes["w:val"];
        var lvlJc = levelElement.firstOrEmpty("w:lvlJc").attributes["w:val"];
        var suff = levelElement.firstOrEmpty("w:suff").attributes["w:val"];
        var tentative = levelElement.attributes["w:tentative"];
        var start = levelElement.firstOrEmpty("w:start").attributes["w:val"];
        
        // Read font information for bullet characters
        var rPr = levelElement.firstOrEmpty("w:rPr");
        var bulletFont = null;
        if (rPr) {
            var rFonts = rPr.firstOrEmpty("w:rFonts");
            if (rFonts) {
                bulletFont = {
                    ascii: rFonts.attributes["w:ascii"],
                    hAnsi: rFonts.attributes["w:hAnsi"],
                    cs: rFonts.attributes["w:cs"],
                    hint: rFonts.attributes["w:hint"]
                };
            }
        }
        
        // Read indentation information
        var pPr = levelElement.firstOrEmpty("w:pPr");
        var indentation = null;
        if (pPr) {
            var ind = pPr.firstOrEmpty("w:ind");
            if (ind) {
                indentation = {
                    left: ind.attributes["w:left"],
                    hanging: ind.attributes["w:hanging"],
                    firstLine: ind.attributes["w:firstLine"]
                };
            }
        }
        
        // Determine if this is a visually displayed bullet list
        var isBullet = numFmt === "bullet";
        var isVisibleBullet = isBullet && !tentative && lvlText !== undefined;
        
        levels[levelIndex] = {
            isOrdered: numFmt !== "bullet",
            format: numFmt,
            delimiter: lvlText,
            level: levelIndex,
            paragraphStyleId: paragraphStyleId,
            justification: lvlJc,
            suffix: suff,
            isTentative: tentative === "1" || tentative === "true",
            startValue: start ? parseInt(start, 10) : 1,
            bulletFont: bulletFont,
            indentation: indentation,
            isBullet: isBullet,
            isVisibleBullet: isVisibleBullet
        };
    });

    var numStyleLink = element.firstOrEmpty("w:numStyleLink").attributes["w:val"];

    return {levels: levels, numStyleLink: numStyleLink};
}

function readNums(root) {
    var nums = {};
    root.getElementsByTagName("w:num").forEach(function(element) {
        var numId = element.attributes["w:numId"];
        var abstractNumId = element.first("w:abstractNumId").attributes["w:val"];
        nums[numId] = {abstractNumId: abstractNumId};
    });
    return nums;
}
