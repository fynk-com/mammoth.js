exports.read = read;
exports.readStyleDefinitions = readStyleDefinitions;
exports._findPartPaths = findPartPaths;

var promises = require("../promises");
var documents = require("../documents");
var Result = require("../results").Result;
var zipfile = require("../zipfile");

var readXmlFromZipFile = require("./office-xml-reader").readXmlFromZipFile;
var createBodyReader = require("./body-reader").createBodyReader;
var DocumentXmlReader = require("./document-xml-reader").DocumentXmlReader;
var relationshipsReader = require("./relationships-reader");
var contentTypesReader = require("./content-types-reader");
var numberingXml = require("./numbering-xml");
var stylesReader = require("./styles-reader");
var notesReader = require("./notes-reader");
var commentsReader = require("./comments-reader");
var Files = require("./files").Files;


function read(docxFile, input, options) {
    input = input || {};
    options = options || {};

    var files = new Files({
        externalFileAccess: options.externalFileAccess,
        relativeToFile: input.path
    });

    return promises.props({
        contentTypes: readContentTypesFromZipFile(docxFile),
        partPaths: findPartPaths(docxFile),
        docxFile: docxFile,
        files: files
    }).also(function(result) {
        return {
            styles: readStylesFromZipFile(docxFile, result.partPaths.styles)
        };
    }).also(function(result) {
        return {
            numbering: readNumberingFromZipFile(docxFile, result.partPaths.numbering, result.styles)
        };
    }).also(function(result) {
        return {
            footnotes: readXmlFileWithBody(result.partPaths.footnotes, result, function(bodyReader, xml) {
                if (xml) {
                    return notesReader.createFootnotesReader(bodyReader)(xml);
                } else {
                    return new Result([]);
                }
            }),
            endnotes: readXmlFileWithBody(result.partPaths.endnotes, result, function(bodyReader, xml) {
                if (xml) {
                    return notesReader.createEndnotesReader(bodyReader)(xml);
                } else {
                    return new Result([]);
                }
            }),
            comments: readXmlFileWithBody(result.partPaths.comments, result, function(bodyReader, xml) {
                if (xml) {
                    return commentsReader.createCommentsReader(bodyReader)(xml);
                } else {
                    return new Result([]);
                }
            }),
            headers: readHeadersOrFooters(result.partPaths.headers, result, documents.Header),
            footers: readHeadersOrFooters(result.partPaths.footers, result, documents.Footer)
        };
    }).also(function(result) {
        return {
            notes: result.footnotes.flatMap(function(footnotes) {
                return result.endnotes.map(function(endnotes) {
                    return new documents.Notes(footnotes.concat(endnotes));
                });
            })
        };
    }).then(function(result) {
        return readXmlFileWithBody(result.partPaths.mainDocument, result, function(bodyReader, xml) {
            return result.notes.flatMap(function(notes) {
                return result.comments.flatMap(function(comments) {
                    return result.headers.flatMap(function(headers) {
                        return result.footers.flatMap(function(footers) {
                            var reader = new DocumentXmlReader({
                                bodyReader: bodyReader,
                                notes: notes,
                                comments: comments,
                                headers: headers,
                                footers: footers
                            });
                            return reader.convertXmlToDocument(xml);
                        });
                    });
                });
            });
        });
    });
}

// Resolved style definitions (see styles-reader) and the abstract numbering
// definitions, keyed by abstractNumId.
function readStyleDefinitions(docxFile) {
    return findPartPaths(docxFile).then(function(partPaths) {
        return readStylesFromZipFile(docxFile, partPaths.styles).then(function(styles) {
            return readNumberingFromZipFile(docxFile, partPaths.numbering, styles).then(function(numbering) {
                return {
                    customStyles: styles.getCustomStyles(),
                    numberingStyles: numbering.getNumberingStyles()
                };
            });
        });
    });
}

function findPartPaths(docxFile) {
    return readPackageRelationships(docxFile).then(function(packageRelationships) {
        var mainDocumentPath = findPartPath({
            docxFile: docxFile,
            relationships: packageRelationships,
            relationshipType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            basePath: "",
            fallbackPath: "word/document.xml"
        });

        if (!docxFile.exists(mainDocumentPath)) {
            throw new Error("Could not find main document part. Are you sure this is a valid .docx file?");
        }

        return xmlFileReader({
            filename: relationshipsFilename(mainDocumentPath),
            readElement: relationshipsReader.readRelationships,
            defaultValue: relationshipsReader.defaultValue
        })(docxFile).then(function(documentRelationships) {
            function findPartRelatedToMainDocument(name) {
                return findPartPath({
                    docxFile: docxFile,
                    relationships: documentRelationships,
                    relationshipType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + name,
                    basePath: zipfile.splitPath(mainDocumentPath).dirname,
                    fallbackPath: "word/" + name + ".xml"
                });
            }

            function findPartsRelatedToMainDocument(name) {
                return findValidTargets({
                    docxFile: docxFile,
                    relationships: documentRelationships,
                    relationshipType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + name,
                    basePath: zipfile.splitPath(mainDocumentPath).dirname
                });
            }

            return {
                mainDocument: mainDocumentPath,
                comments: findPartRelatedToMainDocument("comments"),
                endnotes: findPartRelatedToMainDocument("endnotes"),
                footnotes: findPartRelatedToMainDocument("footnotes"),
                numbering: findPartRelatedToMainDocument("numbering"),
                styles: findPartRelatedToMainDocument("styles"),
                headers: findPartsRelatedToMainDocument("header"),
                footers: findPartsRelatedToMainDocument("footer")
            };
        });
    });
}

function findPartPath(options) {
    var docxFile = options.docxFile;
    var relationships = options.relationships;
    var relationshipType = options.relationshipType;
    var basePath = options.basePath;
    var fallbackPath = options.fallbackPath;

    var validTargets = findValidTargets({
        docxFile: docxFile,
        relationships: relationships,
        relationshipType: relationshipType,
        basePath: basePath
    });
    if (validTargets.length === 0) {
        return fallbackPath;
    } else {
        return validTargets[0];
    }
}

function findValidTargets(options) {
    var targets = options.relationships.findTargetsByType(options.relationshipType);
    var normalisedTargets = targets.map(function(target) {
        return stripPrefix(zipfile.joinPath(options.basePath, target), "/");
    });
    return normalisedTargets.filter(function(target) {
        return options.docxFile.exists(target);
    });
}

function stripPrefix(value, prefix) {
    if (value.substring(0, prefix.length) === prefix) {
        return value.substring(prefix.length);
    } else {
        return value;
    }
}

function xmlFileReader(options) {
    return function(zipFile) {
        return readXmlFromZipFile(zipFile, options.filename)
            .then(function(element) {
                return element ? options.readElement(element) : options.defaultValue;
            });
    };
}

function readXmlFileWithBody(filename, options, func) {
    var readRelationshipsFromZipFile = xmlFileReader({
        filename: relationshipsFilename(filename),
        readElement: relationshipsReader.readRelationships,
        defaultValue: relationshipsReader.defaultValue
    });

    return readRelationshipsFromZipFile(options.docxFile).then(function(relationships) {
        var bodyReader = new createBodyReader({
            relationships: relationships,
            contentTypes: options.contentTypes,
            docxFile: options.docxFile,
            numbering: options.numbering,
            styles: options.styles,
            files: options.files
        });
        return readXmlFromZipFile(options.docxFile, filename)
            .then(function(xml) {
                return func(bodyReader, xml);
            });
    });
}

function relationshipsFilename(filename) {
    var split = zipfile.splitPath(filename);
    return zipfile.joinPath(split.dirname, "_rels", split.basename + ".rels");
}

var readContentTypesFromZipFile = xmlFileReader({
    filename: "[Content_Types].xml",
    readElement: contentTypesReader.readContentTypesFromXml,
    defaultValue: contentTypesReader.defaultContentTypes
});

// Headers and footers in the order the main document's relationships list them.
function readHeadersOrFooters(paths, options, createElement) {
    return promises.Promise.all(paths.map(function(path) {
        return readXmlFileWithBody(path, options, function(bodyReader, xml) {
            return bodyReader.readXmlElements(xml.children).map(createElement);
        });
    })).then(Result.combine);
}

function readNumberingFromZipFile(zipFile, path, styles) {
    return xmlFileReader({
        filename: path,
        readElement: function(element) {
            return numberingXml.readNumberingXml(element, {styles: styles});
        },
        defaultValue: numberingXml.defaultNumbering
    })(zipFile);
}

function readStylesFromZipFile(zipFile, path) {
    return xmlFileReader({
        filename: path,
        readElement: stylesReader.readStylesXml,
        defaultValue: stylesReader.defaultStyles
    })(zipFile);
}

var readPackageRelationships = xmlFileReader({
    filename: "_rels/.rels",
    readElement: relationshipsReader.readRelationships,
    defaultValue: relationshipsReader.defaultValue
});
