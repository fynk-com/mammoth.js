var assert = require("assert");

var readStylesXml = require("../../lib/docx/styles-reader").readStylesXml;
var XmlElement = require("../../lib/xml").Element;
var test = require("../test")(module);


test('paragraph style is null if no style with that ID exists', function() {
    var styles = readStylesXml(
        new XmlElement("w:styles", {}, [])
    );
    assert.equal(styles.findParagraphStyleById("Heading1"), null);
});

test('paragraph style can be found by ID', function() {
    var styles = readStylesXml(
        new XmlElement("w:styles", {}, [
            paragraphStyleElement("Heading1", "Heading 1")
        ])
    );
    assert.equal(styles.findParagraphStyleById("Heading1").styleId, "Heading1");
});

test('table style can be found by ID', function() {
    var styles = readStylesXml(
        new XmlElement("w:styles", {}, [
            tableStyleElement("TableNormal", "Normal Table")
        ])
    );
    assert.equal(styles.findTableStyleById("TableNormal").styleId, "TableNormal");
});

test('character style can be found by ID', function() {
    var styles = readStylesXml(
        new XmlElement("w:styles", {}, [
            characterStyleElement("Heading1Char", "Heading 1 Char")
        ])
    );
    assert.equal(styles.findCharacterStyleById("Heading1Char").styleId, "Heading1Char");
});

test('paragraph and character styles are distinct', function() {
    var styles = readStylesXml(
        new XmlElement("w:styles", {}, [
            paragraphStyleElement("Heading1", "Heading 1"),
            characterStyleElement("Heading1Char", "Heading 1 Char")
        ])
    );
    assert.equal(styles.findCharacterStyleById("Heading1"), null);
    assert.equal(styles.findParagraphStyleById("Heading1Char"), null);
});

test('character and table styles are distinct', function() {
    var styles = readStylesXml(
        new XmlElement("w:styles", {}, [
            tableStyleElement("Heading1", "Heading 1")
        ])
    );
    assert.equal(styles.findCharacterStyleById("Heading1"), null);
});

test('styles include names', function() {
    var styles = readStylesXml(
        new XmlElement("w:styles", {}, [
            paragraphStyleElement("Heading1", "Heading 1")
        ])
    );
    assert.equal(styles.findParagraphStyleById("Heading1").name, "Heading 1");
});

test('style name is null if w:name element does not exist', function() {
    var styles = readStylesXml(
        new XmlElement("w:styles", {}, [
            styleWithoutWNameElement("paragraph", "Heading1"),
            styleWithoutWNameElement("character", "Heading1Char")
        ])
    );
    assert.equal(styles.findParagraphStyleById("Heading1").name, null);
    assert.equal(styles.findCharacterStyleById("Heading1Char").name, null);
});

test('numbering style is null if no style with that ID exists', function() {
    var styles = readStylesXml(
        new XmlElement("w:styles", {}, [])
    );
    assert.equal(styles.findNumberingStyleById("List1"), null);
});

test('numbering style has null numId if style has no paragraph properties', function() {
    var styles = readStylesXml(
        new XmlElement("w:styles", {}, [
            new XmlElement("w:style", {"w:type": "numbering", "w:styleId": "List1"})
        ])
    );
    assert.equal(styles.findNumberingStyleById("List1").numId, null);
});

test('numbering style has numId read from paragraph properties', function() {
    var styles = readStylesXml(
        new XmlElement("w:styles", {}, [
            new XmlElement("w:style", {"w:type": "numbering", "w:styleId": "List1"}, [
                new XmlElement("w:pPr", {}, [
                    new XmlElement("w:numPr", {}, [
                        new XmlElement("w:numId", {"w:val": "42"})
                    ])
                ])
            ])
        ])
    );
    assert.equal(styles.findNumberingStyleById("List1").numId, "42");
});

test('when multiple style elements have same style ID then only first element is used', function() {
    var styles = readStylesXml(
        new XmlElement("w:styles", {}, [
            tableStyleElement("TableNormal", "Normal Table"),
            tableStyleElement("TableNormal", "Table Normal")
        ])
    );
    assert.equal(styles.findTableStyleById("TableNormal").name, "Normal Table");
});

test('w:type of __proto__ does not cause prototype pollution', function() {
    readStylesXml(
        new XmlElement("w:styles", {}, [
            new XmlElement("w:style", {"w:type": "__proto__", "w:styleId": "List1"}, [
                new XmlElement("w:name", {"w:val": "List 1"}, [])
            ])
        ])
    );

    assert.equal({}["List1"], undefined);
});

test('custom styles hold the name, type and formatting of every style', function() {
    var styles = readStylesXml(
        new XmlElement("w:styles", {}, [
            new XmlElement("w:style", {"w:type": "paragraph", "w:styleId": "Body"}, [
                new XmlElement("w:name", {"w:val": "Body"}),
                new XmlElement("w:pPr", {}, [
                    new XmlElement("w:jc", {"w:val": "center"}),
                    new XmlElement("w:spacing", {"w:before": "120", "w:after": "240", "w:line": "360", "w:lineRule": "auto"}),
                    new XmlElement("w:ind", {"w:left": "720", "w:hanging": "360"}),
                    new XmlElement("w:numPr", {}, [
                        new XmlElement("w:ilvl", {"w:val": "1"}),
                        new XmlElement("w:numId", {"w:val": "7"})
                    ])
                ]),
                new XmlElement("w:rPr", {}, [
                    new XmlElement("w:rFonts", {"w:ascii": "Georgia", "w:hAnsi": "Arial"}),
                    new XmlElement("w:sz", {"w:val": "24"}),
                    new XmlElement("w:color", {"w:val": "FF0000"}),
                    new XmlElement("w:highlight", {"w:val": "yellow"}),
                    new XmlElement("w:b"),
                    new XmlElement("w:i", {"w:val": "0"}),
                    new XmlElement("w:u", {"w:val": "single"}),
                    new XmlElement("w:strike", {"w:val": "true"})
                ])
            ])
        ])
    );

    assert.deepEqual(styles.getCustomStyles().Body, {
        name: "Body",
        type: "paragraph",
        properties: {
            alignment: "center",
            spacing: {before: "120", after: "240", line: "360", lineRule: "auto"},
            indent: {left: "720", right: undefined, start: undefined, end: undefined, hanging: "360", firstLine: undefined},
            numbering: {numId: "7", level: "1"},
            font: "Georgia",
            fontSize: 12,
            color: "FF0000",
            highlight: "yellow",
            bold: true,
            underline: true,
            strike: true
        }
    });
});

test('custom styles inherit formatting through w:basedOn', function() {
    var styles = readStylesXml(
        new XmlElement("w:styles", {}, [
            new XmlElement("w:style", {"w:type": "paragraph", "w:styleId": "Derived"}, [
                new XmlElement("w:basedOn", {"w:val": "Base"}),
                new XmlElement("w:pPr", {}, [new XmlElement("w:jc", {"w:val": "right"})])
            ]),
            new XmlElement("w:style", {"w:type": "paragraph", "w:styleId": "Base"}, [
                new XmlElement("w:pPr", {}, [new XmlElement("w:jc", {"w:val": "left"})]),
                new XmlElement("w:rPr", {}, [new XmlElement("w:sz", {"w:val": "28"})])
            ])
        ])
    );

    assert.deepEqual(styles.getCustomStyles().Derived.properties, {alignment: "right", fontSize: 14});
});

test('custom styles tolerate w:basedOn cycles', function() {
    var styles = readStylesXml(
        new XmlElement("w:styles", {}, [
            new XmlElement("w:style", {"w:type": "paragraph", "w:styleId": "A"}, [
                new XmlElement("w:basedOn", {"w:val": "B"}),
                new XmlElement("w:rPr", {}, [new XmlElement("w:sz", {"w:val": "20"})])
            ]),
            new XmlElement("w:style", {"w:type": "paragraph", "w:styleId": "B"}, [
                new XmlElement("w:basedOn", {"w:val": "A"}),
                new XmlElement("w:pPr", {}, [new XmlElement("w:jc", {"w:val": "both"})])
            ])
        ])
    );

    assert.deepEqual(styles.getCustomStyles().A.properties, {alignment: "both", fontSize: 10});
});

function paragraphStyleElement(id, name) {
    return styleElement("paragraph", id, name);
}

function characterStyleElement(id, name) {
    return styleElement("character", id, name);
}

function tableStyleElement(id, name) {
    return styleElement("table", id, name);
}

function styleElement(type, id, name) {
    return new XmlElement("w:style", {"w:type": type, "w:styleId": id}, [
        new XmlElement("w:name", {"w:val": name}, [])
    ]);
}

function styleWithoutWNameElement(type, id) {
    return new XmlElement("w:style", {"w:type": type, "w:styleId": id}, []);
}
