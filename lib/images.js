var _ = require("underscore");

var promises = require("./promises");
var Html = require("./html");
var imageDimensions = require("./image-dimensions");

exports.imgElement = imgElement;

function imgElement(func) {
    return function(element, messages) {
        return promises.when(func(element)).then(function(result) {
            var attributes = {};
            if (element.altText) {
                attributes.alt = element.altText;
            }
            
            // Add original dimensions if available
            if (element.naturalWidth && element.naturalHeight) {
                attributes.dataNaturalWidth = String(element.naturalWidth);
                attributes.dataNaturalHeight = String(element.naturalHeight);
                
                // Add cropped status if available or calculate it
                var isCropped = element.isCropped !== undefined ? element.isCropped :
                               (element.width && element.width < element.naturalWidth) ||
                               (element.height && element.height < element.naturalHeight);
                
                if (isCropped) {
                    attributes.dataIsCropped = 'true';
                }
            }
            
            _.extend(attributes, result);

            return [Html.freshElement("img", attributes)];
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
exports.dataUriWithDimensions = imgElement(function(element) {
    return element.readAsArrayBuffer().then(function(arrayBuffer) {
        var buffer = Buffer.from(arrayBuffer);
        var naturalDims = imageDimensions.getImageDimensions(buffer);
        
        // Add natural dimensions to the element for use in imgElement
        if (naturalDims) {
            element.naturalWidth = naturalDims.width;
            element.naturalHeight = naturalDims.height;
            
            // Calculate if image is cropped (display dimensions smaller than natural)
            element.isCropped = (element.width && element.width < naturalDims.width) ||
                               (element.height && element.height < naturalDims.height);
        }
        
        return element.readAsBase64String().then(function(imageBuffer) {
            return {
                src: "data:" + element.contentType + ";base64," + imageBuffer
            };
        });
    });
});
