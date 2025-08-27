var _ = require("underscore");

var promises = require("./promises");
var Html = require("./html");

exports.imgElement = imgElement;

function imgElement(func) {
    return function(element, messages) {
        // First, try to populate natural dimensions if they're not available
        var processElement = element._populateNaturalDimensions ?
            element._populateNaturalDimensions() :
            promises.resolve(element);
            
        return processElement.then(function(populatedElement) {
            return promises.when(func(populatedElement)).then(function(result) {
                var attributes = {};
                if (populatedElement.altText) {
                    attributes.alt = populatedElement.altText;
                }
                
                // Add original dimensions if available
                if (populatedElement.naturalWidth && populatedElement.naturalHeight) {
                    attributes["data-natural-width"] = String(populatedElement.naturalWidth);
                    attributes["data-natural-height"] = String(populatedElement.naturalHeight);
                    
                    // Add cropped status if available or calculate it
                    var isCropped = populatedElement.isCropped !== undefined ? populatedElement.isCropped :
                                   (populatedElement.width && populatedElement.width < populatedElement.naturalWidth) ||
                                   (populatedElement.height && populatedElement.height < populatedElement.naturalHeight);
                    
                    if (isCropped) {
                        attributes["data-is-cropped"] = 'true';
                    }
                }
                
                _.extend(attributes, result);

                return [Html.freshElement("img", attributes)];
            });
        });
    };
}

// Undocumented, but retained for backwards-compatibility with 0.3.x
exports.inline = exports.imgElement;

exports.dataUri = imgElement(function(element) {
    return element.readAsBase64String().then(function(imageBuffer) {
        return {
            src: "data:" + element.contentType + ";base64," + imageBuffer
        };
    });
});

// Enhanced version that includes natural dimensions
// Note: Natural dimensions are now automatically extracted by imgElement
exports.dataUriWithDimensions = imgElement(function(element) {
    return element.readAsBase64String().then(function(imageBuffer) {
        return {
            src: "data:" + element.contentType + ";base64," + imageBuffer
        };
    });
});
