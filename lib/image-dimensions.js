// Simple image dimension reader for common formats
function getImageDimensions(buffer) {
    try {
        // PNG format
        if (buffer.length >= 24 &&
            buffer[0] === 0x89 && buffer[1] === 0x50 && buffer[2] === 0x4E && buffer[3] === 0x47) {
            var pngWidth = buffer.readUInt32BE(16);
            var pngHeight = buffer.readUInt32BE(20);
            return {width: pngWidth, height: pngHeight};
        }

        // JPEG format
        if (buffer.length >= 4 && buffer[0] === 0xFF && buffer[1] === 0xD8) {
            for (var i = 2; i < buffer.length - 4; i++) {
                if (buffer[i] === 0xFF && (buffer[i + 1] === 0xC0 || buffer[i + 1] === 0xC2)) {
                    var jpegHeight = buffer.readUInt16BE(i + 5);
                    var jpegWidth = buffer.readUInt16BE(i + 7);
                    return {width: jpegWidth, height: jpegHeight};
                }
            }
        }

        // GIF format
        if (buffer.length >= 10 &&
            buffer[0] === 0x47 && buffer[1] === 0x49 && buffer[2] === 0x46) {
            var gifWidth = buffer.readUInt16LE(6);
            var gifHeight = buffer.readUInt16LE(8);
            return {width: gifWidth, height: gifHeight};
        }

        // BMP format
        if (buffer.length >= 26 &&
            buffer[0] === 0x42 && buffer[1] === 0x4D) {
            var bmpWidth = buffer.readUInt32LE(18);
            var bmpHeight = buffer.readUInt32LE(22);
            return {width: bmpWidth, height: bmpHeight};
        }

        // WEBP format
        if (buffer.length >= 30 &&
            buffer[0] === 0x52 && buffer[1] === 0x49 && buffer[2] === 0x46 && buffer[3] === 0x46 &&
            buffer[8] === 0x57 && buffer[9] === 0x45 && buffer[10] === 0x42 && buffer[11] === 0x50) {
            // VP8 format
            if (buffer[12] === 0x56 && buffer[13] === 0x50 && buffer[14] === 0x38 && buffer[15] === 0x20) {
                var webpWidth = buffer.readUInt16LE(26) & 0x3FFF;
                var webpHeight = buffer.readUInt16LE(28) & 0x3FFF;
                return {width: webpWidth, height: webpHeight};
            }
        }

    } catch (e) {
        // If parsing fails, return null
    }

    return null;
}

exports.getImageDimensions = getImageDimensions;
