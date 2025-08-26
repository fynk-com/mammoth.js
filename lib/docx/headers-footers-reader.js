var documents = require("../documents");

exports.createHeadersReader = createReader.bind(this, "header");
exports.createFootersReader = createReader.bind(this, "footer");

function createReader(partType, bodyReader) {
    function readHeaderFooterXml(element, headerFooterType, sectionIndex) {
        return bodyReader.readXmlElements(element.children)
            .map(function(body) {
                if (partType === "header") {
                    return new documents.Header(body, {
                        headerType: headerFooterType,
                        sectionIndex: sectionIndex
                    });
                } else {
                    return new documents.Footer(body, {
                        footerType: headerFooterType,
                        sectionIndex: sectionIndex
                    });
                }
            });
    }
    
    return readHeaderFooterXml;
}
