var promises = require("../promises");
var _ = require("underscore");

var xmldom = require("./xmldom");
var nodes = require("./nodes");
var Element = nodes.Element;

exports.readString = readString;

var Node = xmldom.Node;

function readString(xmlString, namespaceMap) {
    namespaceMap = namespaceMap || {};

    try {
        // Parse the XML document using xmldom
        var document = xmldom.parseFromString(xmlString, "text/xml");
    } catch (error) {
        console.error("Parsing error:", error);
        return promises.reject(new Error("Failed to parse XML string. Invalid syntax or content."));
    }

    if (document.documentElement.tagName === "parsererror" || document.documentElement.tagName.toLowerCase() === "html") {
        console.error("Parser error:", document.documentElement.textContent);
        return promises.reject(new Error(document.documentElement.textContent));
    }

    function convertNode(node) {
        switch (node.nodeType) {
            case Node.ELEMENT_NODE:
                return convertElement(node);
            case Node.TEXT_NODE:
                return nodes.text(node.nodeValue);
        }
    }

    function convertElement(element) {
        var convertedName = convertName(element);

        var convertedChildren = [];
        _.forEach(element.childNodes, function(childNode) {
            var convertedNode = convertNode(childNode);
            if (convertedNode) {
                convertedChildren.push(convertedNode);
            }
        });

        var convertedAttributes = {};
        _.forEach(element.attributes, function(attribute) {
            convertedAttributes[convertName(attribute)] = attribute.value;
        });

        return new Element(convertedName, convertedAttributes, convertedChildren);
    }

    function convertName(node) {
        if (node.namespaceURI) {
            var mappedPrefix = namespaceMap[node.namespaceURI];
            var prefix;
            if (mappedPrefix) {
                prefix = mappedPrefix + ":";
            } else {
                prefix = "{" + node.namespaceURI + "}";
            }
            return prefix + node.localName;
        } else {
            return node.localName;
        }
    }

    // Return the resolved promise containing the converted XML root node
    return promises.resolve(convertNode(document.documentElement));
}