var _ = require("underscore");

var types = exports.types = {
    document: "document",
    paragraph: "paragraph",
    ins: "ins",
    del: "del",
    run: "run",
    text: "text",
    tab: "tab",
    checkbox: "checkbox",
    hyperlink: "hyperlink",
    noteReference: "noteReference",
    image: "image",
    note: "note",
    commentReference: "commentReference",
    commentRangeStart: "commentRangeStart",
    commentRangeEnd: "commentRangeEnd",
    comment: "comment",
    table: "table",
    tableRow: "tableRow",
    tableCell: "tableCell",
    "break": "break",
    bookmarkStart: "bookmarkStart",
    header: "header",
    footer: "footer"
};

function Document(children, options) {
    options = options || {};
    return {
        type: types.document,
        children: children,
        notes: options.notes || new Notes({}),
        comments: options.comments || [],
        headers: options.headers || [],
        footers: options.footers || []
    };
}

function Paragraph(children, properties) {
    properties = properties || {};
    return {
        type: types.paragraph,
        children: children,
        styleId: properties.styleId || null,
        styleName: properties.styleName || null,
        numbering: properties.numbering || null,
        alignment: properties.alignment || null,
        indent: properties.indent || null,
        spacing: properties.spacing || null
    };
}

function Ins(children, properties) {
    properties = properties || {};
    return {
        type: types.ins,
        children: children,
        authorName: properties.authorName,
        changeId: properties.changeId,
        date: properties.date
    };
}

function Del(children, properties) {
    properties = properties || {};
    return {
        type: types.del,
        children: children,
        authorName: properties.authorName,
        changeId: properties.changeId,
        date: properties.date
    };
}

function Run(children, properties) {
    properties = properties || {};
    return {
        type: types.run,
        children: children,
        styleId: properties.styleId || null,
        styleName: properties.styleName || null,
        isBold: !!properties.isBold,
        isUnderline: !!properties.isUnderline,
        isItalic: !!properties.isItalic,
        isStrikethrough: !!properties.isStrikethrough,
        isAllCaps: !!properties.isAllCaps,
        isSmallCaps: !!properties.isSmallCaps,
        verticalAlignment: properties.verticalAlignment || verticalAlignment.baseline,
        font: properties.font || null,
        fontSize: properties.fontSize || null,
        color: properties.color || null,
        highlight: properties.highlight || null,
        shading: properties.shading || null
    };
}

var verticalAlignment = {
    baseline: "baseline",
    superscript: "superscript",
    subscript: "subscript"
};

function Text(value) {
    return {
        type: types.text,
        value: value
    };
}

function Tab() {
    return {
        type: types.tab
    };
}

function Checkbox(options) {
    return {
        type: types.checkbox,
        checked: options.checked
    };
}

function Hyperlink(children, options) {
    return {
        type: types.hyperlink,
        children: children,
        href: options.href,
        anchor: options.anchor,
        targetFrame: options.targetFrame
    };
}

function NoteReference(options) {
    return {
        type: types.noteReference,
        noteType: options.noteType,
        noteId: options.noteId
    };
}

function Notes(notes) {
    this._notes = _.indexBy(notes, function(note) {
        return noteKey(note.noteType, note.noteId);
    });
}

Notes.prototype.resolve = function(reference) {
    return this.findNoteByKey(noteKey(reference.noteType, reference.noteId));
};

Notes.prototype.findNoteByKey = function(key) {
    return this._notes[key] || null;
};

function Note(options) {
    return {
        type: types.note,
        noteType: options.noteType,
        noteId: options.noteId,
        body: options.body
    };
}

function commentReference(options) {
    return {
        type: types.commentReference,
        commentId: options.commentId
    };
}

function commentRangeStart(options) {
    return {
        type: types.commentRangeStart,
        commentId: options.commentId
    };
}

function commentRangeEnd(options) {
    return {
        type: types.commentRangeEnd,
        commentId: options.commentId
    };
}

function comment(options) {
    return {
        type: types.comment,
        commentId: options.commentId,
        body: options.body,
        authorName: options.authorName,
        authorInitials: options.authorInitials,
        date: options.date
    };
}

function noteKey(noteType, id) {
    return noteType + "-" + id;
}

function Image(options) {
    return {
        type: types.image,
        // `read` is retained for backwards compatibility, but other read
        // methods should be preferred.
        read: function(encoding) {
            if (encoding) {
                return options.readImage(encoding);
            } else {
                return options.readImage().then(function(arrayBuffer) {
                    return Buffer.from(arrayBuffer);
                });
            }
        },
        readAsArrayBuffer: function() {
            return options.readImage();
        },
        readAsBase64String: function() {
            return options.readImage("base64");
        },
        readAsBuffer: function() {
            return options.readImage().then(function(arrayBuffer) {
                return Buffer.from(arrayBuffer);
            });
        },
        altText: options.altText,
        contentType: options.contentType,
        width: options.width,
        height: options.height,
        srcRect: options.srcRect,
        isCropped: options.isCropped,
        imageProperties: options.imageProperties
    };
}

function Table(children, properties) {
    properties = properties || {};
    return {
        type: types.table,
        children: children,
        styleId: properties.styleId || null,
        styleName: properties.styleName || null,
        isBordered: properties.isBordered || false,
        borders: properties.borders || null
    };
}

function TableRow(children, options) {
    options = options || {};
    return {
        type: types.tableRow,
        children: children,
        isHeader: options.isHeader || false
    };
}

function TableCell(children, options) {
    options = options || {};
    return {
        type: types.tableCell,
        children: children,
        width: options.width,
        colSpan: options.colSpan == null ? 1 : options.colSpan,
        rowSpan: options.rowSpan == null ? 1 : options.rowSpan,
        bgColor: options.bgColor,
        borders: options.borders || null
    };
}

function Break(breakType) {
    return {
        type: types["break"],
        breakType: breakType
    };
}

function BookmarkStart(options) {
    return {
        type: types.bookmarkStart,
        name: options.name
    };
}

function Header(children, properties) {
    properties = properties || {};
    return {
        type: types.header,
        children: children,
        headerType: properties.headerType || "default",
        sectionIndex: properties.sectionIndex || 0
    };
}

function Footer(children, properties) {
    properties = properties || {};
    return {
        type: types.footer,
        children: children,
        footerType: properties.footerType || "default",
        sectionIndex: properties.sectionIndex || 0
    };
}

exports.document = exports.Document = Document;
exports.paragraph = exports.Paragraph = Paragraph;
exports.ins = exports.Ins = Ins;
exports.del = exports.Del = Del;
exports.run = exports.Run = Run;
exports.text = exports.Text = Text;
exports.tab = exports.Tab = Tab;
exports.checkbox = exports.Checkbox = Checkbox;
exports.Hyperlink = Hyperlink;
exports.noteReference = exports.NoteReference = NoteReference;
exports.Notes = Notes;
exports.Note = Note;
exports.commentReference = commentReference;
exports.commentRangeStart = commentRangeStart;
exports.commentRangeEnd = commentRangeEnd;
exports.comment = comment;
exports.Image = Image;
exports.Table = Table;
exports.TableRow = TableRow;
exports.TableCell = TableCell;
exports.lineBreak = Break("line");
exports.pageBreak = Break("page");
exports.columnBreak = Break("column");
exports.BookmarkStart = BookmarkStart;
exports.Header = Header;
exports.Footer = Footer;

exports.verticalAlignment = verticalAlignment;
