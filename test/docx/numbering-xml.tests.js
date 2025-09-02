var assert = require("assert");
var duck = require("duck");

var readNumberingXml = require("../../lib/docx/numbering-xml").readNumberingXml;
var stylesReader = require("../../lib/docx/styles-reader");
var XmlElement = require("../../lib/xml").Element;
var test = require("../test")(module);


test('w:num element inherits levels from w:abstractNum', function() {
    var numbering = readNumberingXml(
        new XmlElement("w:numbering", {}, [
            new XmlElement("w:abstractNum", {"w:abstractNumId": "42"}, [
                new XmlElement("w:lvl", {"w:ilvl": "0"}, [
                    new XmlElement("w:numFmt", {"w:val": "bullet"})
                ]),
                new XmlElement("w:lvl", {"w:ilvl": "1"}, [
                    new XmlElement("w:numFmt", {"w:val": "decimal"})
                ])
            ]),
            new XmlElement("w:num", {"w:numId": "47"}, [
                new XmlElement("w:abstractNumId", {"w:val": "42"})
            ])
        ]),
        {styles: stylesReader.defaultStyles}
    );
    duck.assertThat(numbering.findLevel("47", "0"), duck.hasProperties({
        isOrdered: false
    }));
    duck.assertThat(numbering.findLevel("47", "1"), duck.hasProperties({
        isOrdered: true
    }));
});


test('w:num element referencing non-existent w:abstractNumId is ignored', function() {
    var numbering = readNumberingXml(
        new XmlElement("w:numbering", {}, [
            new XmlElement("w:num", {"w:numId": "47"}, [
                new XmlElement("w:abstractNumId", {"w:val": "42"})
            ])
        ]),
        {styles: stylesReader.defaultStyles}
    );
    duck.assertThat(numbering.findLevel("47", "0"), duck.equalTo(null));
});

test('when level is missing w:numFmt then level is ordered', function() {
    var numbering = readNumberingXml(
        new XmlElement("w:numbering", {}, [
            new XmlElement("w:abstractNum", {"w:abstractNumId": "42"}, [
                new XmlElement("w:lvl", {"w:ilvl": "0"})
            ]),
            new XmlElement("w:num", {"w:numId": "47"}, [
                new XmlElement("w:abstractNumId", {"w:val": "42"})
            ])
        ]),
        {styles: stylesReader.defaultStyles}
    );
    duck.assertThat(numbering.findLevel("47", "0"), duck.hasProperties({
        isOrdered: true
    }));
});


test('when w:abstractNum has w:numStyleLink then style is used to find w:num', function() {
    var numbering = readNumberingXml(
        new XmlElement("w:numbering", {}, [
            new XmlElement("w:abstractNum", {"w:abstractNumId": "100"}, [
                new XmlElement("w:lvl", {"w:ilvl": "0"}, [
                    new XmlElement("w:numFmt", {"w:val": "decimal"})
                ])
            ]),
            new XmlElement("w:abstractNum", {"w:abstractNumId": "101"}, [
                new XmlElement("w:numStyleLink", {"w:val": "List1"})
            ]),
            new XmlElement("w:num", {"w:numId": "200"}, [
                new XmlElement("w:abstractNumId", {"w:val": "100"})
            ]),
            new XmlElement("w:num", {"w:numId": "201"}, [
                new XmlElement("w:abstractNumId", {"w:val": "101"})
            ])
        ]),
        {styles: new stylesReader.Styles({}, {}, {}, {"List1": {numId: "200"}})}
    );
    duck.assertThat(numbering.findLevel("201", "0"), duck.hasProperties({
        isOrdered: true
    }));
});


// See: 17.9.23 pStyle (Paragraph Style's Associated Numbering Level) in ECMA-376, 4th Edition
test('numbering level can be found by paragraph style ID', function() {
    var numbering = readNumberingXml(
        new XmlElement("w:numbering", {}, [
            new XmlElement("w:abstractNum", {"w:abstractNumId": "42"}, [
                new XmlElement("w:lvl", {"w:ilvl": "0"}, [
                    new XmlElement("w:numFmt", {"w:val": "bullet"})
                ])
            ]),
            new XmlElement("w:abstractNum", {"w:abstractNumId": "43"}, [
                new XmlElement("w:lvl", {"w:ilvl": "0"}, [
                    new XmlElement("w:pStyle", {"w:val": "List"}),
                    new XmlElement("w:numFmt", {"w:val": "decimal"})
                ])
            ])
        ]),
        {styles: stylesReader.defaultStyles}
    );
    duck.assertThat(numbering.findLevelByParagraphStyleId("List"), duck.hasProperties({
        isOrdered: true
    }));
    duck.assertThat(numbering.findLevelByParagraphStyleId("Paragraph"), duck.equalTo(null));
});

test('when styles is missing then error is thrown', function() {
    assert.throws(function() {
        readNumberingXml(new XmlElement("w:numbering", {}, []));
    }, /styles is missing/);
});

test('isVisibleBulletList returns true for bullet list with proper formatting', function() {
    var numbering = readNumberingXml(
        new XmlElement("w:numbering", {}, [
            new XmlElement("w:abstractNum", {"w:abstractNumId": "42"}, [
                new XmlElement("w:lvl", {"w:ilvl": "0"}, [
                    new XmlElement("w:numFmt", {"w:val": "bullet"}),
                    new XmlElement("w:lvlText", {"w:val": "•"}),
                    new XmlElement("w:rPr", {}, [
                        new XmlElement("w:rFonts", {"w:ascii": "Symbol", "w:hAnsi": "Symbol"})
                    ])
                ])
            ]),
            new XmlElement("w:num", {"w:numId": "47"}, [
                new XmlElement("w:abstractNumId", {"w:val": "42"})
            ])
        ]),
        {styles: stylesReader.defaultStyles}
    );
    duck.assertThat(numbering.isVisibleBulletList("47", "0"), duck.equalTo(true));
});

test('isVisibleBulletList returns false for tentative bullet list', function() {
    var numbering = readNumberingXml(
        new XmlElement("w:numbering", {}, [
            new XmlElement("w:abstractNum", {"w:abstractNumId": "42"}, [
                new XmlElement("w:lvl", {"w:ilvl": "0", "w:tentative": "1"}, [
                    new XmlElement("w:numFmt", {"w:val": "bullet"}),
                    new XmlElement("w:lvlText", {"w:val": "•"})
                ])
            ]),
            new XmlElement("w:num", {"w:numId": "47"}, [
                new XmlElement("w:abstractNumId", {"w:val": "42"})
            ])
        ]),
        {styles: stylesReader.defaultStyles}
    );
    duck.assertThat(numbering.isVisibleBulletList("47", "0"), duck.equalTo(false));
});

test('isVisibleBulletList returns false for non-bullet list', function() {
    var numbering = readNumberingXml(
        new XmlElement("w:numbering", {}, [
            new XmlElement("w:abstractNum", {"w:abstractNumId": "42"}, [
                new XmlElement("w:lvl", {"w:ilvl": "0"}, [
                    new XmlElement("w:numFmt", {"w:val": "decimal"})
                ])
            ]),
            new XmlElement("w:num", {"w:numId": "47"}, [
                new XmlElement("w:abstractNumId", {"w:val": "42"})
            ])
        ]),
        {styles: stylesReader.defaultStyles}
    );
    duck.assertThat(numbering.isVisibleBulletList("47", "0"), duck.equalTo(false));
});

test('getBulletCharacter returns correct character for bullet list', function() {
    var numbering = readNumberingXml(
        new XmlElement("w:numbering", {}, [
            new XmlElement("w:abstractNum", {"w:abstractNumId": "42"}, [
                new XmlElement("w:lvl", {"w:ilvl": "0"}, [
                    new XmlElement("w:numFmt", {"w:val": "bullet"}),
                    new XmlElement("w:lvlText", {"w:val": "→"})
                ])
            ]),
            new XmlElement("w:num", {"w:numId": "47"}, [
                new XmlElement("w:abstractNumId", {"w:val": "42"})
            ])
        ]),
        {styles: stylesReader.defaultStyles}
    );
    duck.assertThat(numbering.getBulletCharacter("47", "0"), duck.equalTo("→"));
});

test('level properties include extended bullet information', function() {
    var numbering = readNumberingXml(
        new XmlElement("w:numbering", {}, [
            new XmlElement("w:abstractNum", {"w:abstractNumId": "42"}, [
                new XmlElement("w:lvl", {"w:ilvl": "0", "w:tentative": "1"}, [
                    new XmlElement("w:numFmt", {"w:val": "bullet"}),
                    new XmlElement("w:lvlText", {"w:val": "•"}),
                    new XmlElement("w:lvlJc", {"w:val": "left"}),
                    new XmlElement("w:suff", {"w:val": "tab"}),
                    new XmlElement("w:start", {"w:val": "1"}),
                    new XmlElement("w:pPr", {}, [
                        new XmlElement("w:ind", {"w:left": "720", "w:hanging": "360"})
                    ]),
                    new XmlElement("w:rPr", {}, [
                        new XmlElement("w:rFonts", {"w:ascii": "Symbol", "w:hAnsi": "Symbol", "w:hint": "default"})
                    ])
                ])
            ]),
            new XmlElement("w:num", {"w:numId": "47"}, [
                new XmlElement("w:abstractNumId", {"w:val": "42"})
            ])
        ]),
        {styles: stylesReader.defaultStyles}
    );
    var level = numbering.findLevel("47", "0");
    duck.assertThat(level, duck.hasProperties({
        isBullet: true,
        isVisibleBullet: false, // tentative = true
        isTentative: true,
        justification: "left",
        suffix: "tab",
        startValue: 1,
        bulletFont: {
            ascii: "Symbol",
            hAnsi: "Symbol",
            cs: undefined,
            hint: "default"
        },
        indentation: {
            left: "720",
            hanging: "360",
            firstLine: undefined
        }
    }));
});

test('isVisibleBulletList returns false for non-existent level', function() {
    var numbering = readNumberingXml(
        new XmlElement("w:numbering", {}, [
            new XmlElement("w:abstractNum", {"w:abstractNumId": "42"}, [
                new XmlElement("w:lvl", {"w:ilvl": "1"}, [
                    new XmlElement("w:numFmt", {"w:val": "bullet"}),
                    new XmlElement("w:lvlText", {"w:val": "•"})
                ])
            ]),
            new XmlElement("w:num", {"w:numId": "47"}, [
                new XmlElement("w:abstractNumId", {"w:val": "42"})
            ])
        ]),
        {styles: stylesReader.defaultStyles}
    );
    // Level 0 doesn't exist, only level 1 exists
    duck.assertThat(numbering.isVisibleBulletList("47", "0"), duck.equalTo(false));
    duck.assertThat(numbering.findLevel("47", "0"), duck.equalTo(null));
});
