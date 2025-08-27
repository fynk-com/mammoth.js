exports.read = read;
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
var headersFootersReader = require("./headers-footers-reader");
var Files = require("./files").Files;


function read(docxFile, input) {
    input = input || {};

    return promises.props({
        contentTypes: readContentTypesFromZipFile(docxFile),
        partPaths: findPartPaths(docxFile),
        docxFile: docxFile,
        files: input.path ? Files.relativeToFile(input.path) : new Files(null)
    }).also(function(result) {
        return {
            theme: readThemeFromZipFile(docxFile, result.partPaths.theme)
        };
    }).also(function(result) {
        return {
            styles: readStylesFromZipFileWithTheme(docxFile, result.partPaths.styles, result.theme)
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
            })
        };
    }).also(function(result) {
        return {
            notes: result.footnotes.flatMap(function(footnotes) {
                return result.endnotes.map(function(endnotes) {
                    return new documents.Notes(footnotes.concat(endnotes));
                });
            })
        };
    }).also(function(result) {
        return {
            headers: parseHeadersFooters(result.partPaths.headers, result, "header"),
            footers: parseHeadersFooters(result.partPaths.footers, result, "footer")
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
        }).then(function(finalDocument) {
            // Append the custom styles and document defaults to the final result
            var numberingStyles = result.numbering.getNumberingStyles(); // Get the numbering styles
            var customStyles = result.styles.getCustomStyles(); // Get the custom styles
            var documentDefaults = result.styles.getDocumentDefaults(); // Get the document defaults
            return Object.assign(finalDocument, {
                customStyles: customStyles,
                numberingStyles: numberingStyles,
                documentDefaults: documentDefaults
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

            return {
                mainDocument: mainDocumentPath,
                comments: findPartRelatedToMainDocument("comments"),
                endnotes: findPartRelatedToMainDocument("endnotes"),
                footnotes: findPartRelatedToMainDocument("footnotes"),
                numbering: findPartRelatedToMainDocument("numbering"),
                styles: findPartRelatedToMainDocument("styles"),
                theme: findPartRelatedToMainDocument("theme"),
                headers: findPartsByTypeWithBasePath(documentRelationships, "header", zipfile.splitPath(mainDocumentPath).dirname),
                footers: findPartsByTypeWithBasePath(documentRelationships, "footer", zipfile.splitPath(mainDocumentPath).dirname)
            };
        });
    });
}

function findPartsByTypeWithBasePath(relationships, type, basePath) {
    var relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + type;
    var targets = relationships.findTargetsByType(relationshipType);
    
    // Add the base path (usually "word/") to each target
    return targets.map(function(target) {
        return stripPrefix(zipfile.joinPath(basePath, target), "/");
    });
}

function findPartPath(options) {
    var docxFile = options.docxFile;
    var relationships = options.relationships;
    var relationshipType = options.relationshipType;
    var basePath = options.basePath;
    var fallbackPath = options.fallbackPath;

    var targets = relationships.findTargetsByType(relationshipType);
    var normalisedTargets = targets.map(function(target) {
        return stripPrefix(zipfile.joinPath(basePath, target), "/");
    });
    var validTargets = normalisedTargets.filter(function(target) {
        return docxFile.exists(target);
    });
    if (validTargets.length === 0) {
        return fallbackPath;
    } else {
        return validTargets[0];
    }
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

function readNumberingFromZipFile(zipFile, path, styles) {
    return xmlFileReader({
        filename: path,
        readElement: function(element) {
            return numberingXml.readNumberingXml(element, {styles: styles});
        },
        defaultValue: numberingXml.defaultNumbering
    })(zipFile);
}

// function readStylesFromZipFile(zipFile, path) {
//     return xmlFileReader({
//         filename: path,
//         readElement: stylesReader.readStylesXml,
//         defaultValue: stylesReader.defaultStyles
//     })(zipFile);
// }

function readStylesFromZipFileWithTheme(zipFile, path, theme) {
    return xmlFileReader({
        filename: path,
        readElement: function(element) {
            return stylesReader.readStylesXmlWithTheme(element, theme);
        },
        defaultValue: stylesReader.defaultStyles
    })(zipFile);
}

function readThemeFromZipFile(zipFile, path) {
    return xmlFileReader({
        filename: path,
        readElement: readThemeXml,
        defaultValue: null
    })(zipFile);
}

function readThemeXml(root) {
    var themeElements = root.first("a:themeElements");
    if (!themeElements) {
        return null;
    }

    var fontScheme = themeElements.first("a:fontScheme");
    if (!fontScheme) {
        return null;
    }

    var theme = {
        name: root.attributes["name"] || "Unknown Theme"
    };

    // Read major font (used for headings)
    var majorFont = fontScheme.first("a:majorFont");
    if (majorFont) {
        var majorLatin = majorFont.first("a:latin");
        theme.majorFont = majorLatin ? majorLatin.attributes["typeface"] : null;
    }

    // Read minor font (used for body text)
    var minorFont = fontScheme.first("a:minorFont");
    if (minorFont) {
        var minorLatin = minorFont.first("a:latin");
        theme.minorFont = minorLatin ? minorLatin.attributes["typeface"] : null;
    }

    return theme;
}

var readPackageRelationships = xmlFileReader({
    filename: "_rels/.rels",
    readElement: relationshipsReader.readRelationships,
    defaultValue: relationshipsReader.defaultValue
});

function parseHeadersFooters(partPaths, result, type) {
    if (!partPaths || partPaths.length === 0) {
        return promises.resolve(new Result([]));
    }

    var reader = type === "header" ?
        headersFootersReader.createHeadersReader :
        headersFootersReader.createFootersReader;
    
    return promises.all(partPaths.map(function(path, index) {
        return readXmlFileWithBody(path, result, function(bodyReader, xml) {
            if (xml) {
                var headerFooterType = determineHeaderFooterType(path);
                return reader(bodyReader)(xml, headerFooterType, index);
            } else {
                return new Result([]);
            }
        });
    })).then(function(results) {
        // Combine all results into a single result with all items
        var allItems = [];
        var allMessages = [];
        results.forEach(function(res) {
            if (res.value) {
                allItems.push(res.value);
            }
            allMessages = allMessages.concat(res.messages || []);
        });
        
        return new Result(allItems, allMessages);
    });
}

function determineHeaderFooterType(path) {
    if (path.includes("header1.xml") || path.includes("footer1.xml")) {
        return "first";
    }
    if (path.includes("header2.xml") || path.includes("footer2.xml")) {
        return "even";
    }
    if (path.includes("header3.xml") || path.includes("footer3.xml")) {
        return "odd";
    }
    return "default";
}
