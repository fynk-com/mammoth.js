interface Mammoth {
    convertToHtml: (input: Input, options?: Options) => Promise<mammoth.ConversionResult>;
    extractRawText: (input: Input) => Promise<Result>;
    embedStyleMap: (input: Input, styleMap: string) => Promise<{
        toArrayBuffer: () => ArrayBuffer,
        toBuffer: () => Buffer,
    }>;
    images: Images;
}

type Input = NodeJsInput | BrowserInput;

type NodeJsInput = PathInput | BufferInput;

interface PathInput {
    path: string;
}

interface BufferInput {
    buffer: Buffer;
}

type BrowserInput = ArrayBufferInput;

interface ArrayBufferInput {
    arrayBuffer: ArrayBuffer;
}

interface Options {
    styleMap?: string | Array<string>;
    includeEmbeddedStyleMap?: boolean;
    includeDefaultStyleMap?: boolean;
    convertImage?: ImageConverter;
    ignoreEmptyParagraphs?: boolean;
    idPrefix?: string;
    externalFileAccess?: boolean;
    transformDocument?: (element: mammoth.Document) => any;
}

interface ImageConverter {
    __mammothBrand: "ImageConverter";
}

interface Image {
    contentType: string;
    readAsArrayBuffer: () => Promise<ArrayBuffer>;
    readAsBase64String: () => Promise<string>;
    readAsBuffer: () => Promise<Buffer>;
    read: ImageRead;
}

type ImageContent = Image;

interface ImageRead {
    (): Promise<Buffer>;
    (encoding: string): Promise<string>;
}

interface ImageAttributes {
    src: string;
}

interface Images {
    dataUri: ImageConverter;
    imgElement: (f: (image: Image) => Promise<ImageAttributes>) => ImageConverter;
}

interface Result {
    value: string;
    messages: Array<Message>;
}

type Message = Warning | Error;

interface Warning {
    type: "warning";
    message: string;
}

interface Error {
    type: "error";
    message: string;
    error: unknown;
}

declare namespace mammoth {
    type ConversionOptions = Options;

    interface ConversionResult extends Result {
        customStyles: {[styleId: string]: CustomStyle};
        numberingStyles: {[abstractNumId: string]: AbstractNumbering};
    }

    interface CustomStyle {
        name: string | null;
        type: string;
        properties: StyleProperties;
    }

    // Formatting resolved through w:basedOn. A property is absent when
    // neither the style nor its base styles set it.
    interface StyleProperties {
        font?: string;
        fontSize?: number;
        color?: string;
        highlight?: string;
        bold?: true;
        italic?: true;
        underline?: true;
        strike?: true;
        alignment?: string;
        spacing?: {before?: string, after?: string, line?: string, lineRule?: string};
        indent?: {left?: string, right?: string, start?: string, end?: string, firstLine?: string, hanging?: string};
        numbering?: {numId?: string, level?: string};
    }

    interface AbstractNumbering {
        levels: {[level: string]: NumberingLevel};
        numStyleLink: string | undefined;
    }

    interface NumberingLevel {
        isOrdered: boolean;
        format: string | undefined;
        delimiter: string | undefined;
        level: string;
        paragraphStyleId: string | undefined;
    }

    interface Document {
        type: "document";
        children: Array<DocumentElement>;
        notes: Notes;
        comments: Array<Comment>;
        headers: Array<Header>;
        footers: Array<Footer>;
    }

    interface Notes {
        resolve(reference: NoteReference): Note | null;
    }

    interface Note {
        type: "note";
        noteType: "footnote" | "endnote";
        noteId: string;
        body: Array<DocumentElement>;
    }

    interface Comment {
        type: "comment";
        commentId: string;
        body: Array<DocumentElement>;
        authorName: string | null;
        authorInitials: string | null;
        date: string | null;
    }

    interface Header {
        type: "header";
        children: Array<DocumentElement>;
    }

    interface Footer {
        type: "footer";
        children: Array<DocumentElement>;
    }

    type DocumentElement =
        | Paragraph
        | Run
        | Text
        | Tab
        | Checkbox
        | Hyperlink
        | NoteReference
        | CommentReference
        | CommentRangeStart
        | CommentRangeEnd
        | Image
        | Table
        | TableRow
        | TableCell
        | Break
        | BookmarkStart
        | Ins
        | Del;

    interface Paragraph {
        type: "paragraph";
        children: Array<DocumentElement>;
        styleId: string | null;
        styleName: string | null;
        numbering: ParagraphNumbering | null;
        alignment: string | null;
        indent: {
            start: string | null,
            end: string | null,
            left: string | null,
            right: string | null,
            firstLine: string | null,
            hanging: string | null,
        };
        spacing: {before: string | null, after: string | null, line: string | null, lineRule: string | null};
    }

    interface ParagraphNumbering extends NumberingLevel {
        numId?: string;
    }

    interface Run {
        type: "run";
        children: Array<DocumentElement>;
        styleId: string | null;
        styleName: string | null;
        isBold: boolean;
        isUnderline: boolean;
        isItalic: boolean;
        isStrikethrough: boolean;
        isAllCaps: boolean;
        isSmallCaps: boolean;
        verticalAlignment: "baseline" | "superscript" | "subscript";
        font: string | null;
        fontSize: number | null;
        color: string | null;
        highlight: string | null;
    }

    interface Text {
        type: "text";
        value: string;
    }

    interface Tab {
        type: "tab";
    }

    interface Checkbox {
        type: "checkbox";
        checked: boolean;
    }

    interface Hyperlink {
        type: "hyperlink";
        children: Array<DocumentElement>;
        href?: string;
        anchor?: string;
        targetFrame?: string | null;
    }

    interface NoteReference {
        type: "noteReference";
        noteType: "footnote" | "endnote";
        noteId: string;
    }

    interface CommentReference {
        type: "commentReference";
        commentId: string;
    }

    interface CommentRangeStart {
        type: "commentRangeStart";
        commentId: string;
    }

    interface CommentRangeEnd {
        type: "commentRangeEnd";
        commentId: string;
    }

    interface Image extends ImageContent {
        type: "image";
        altText: string | undefined;
        width: number | null;
        height: number | null;
        imageProperties: {
            floating: "inline" | "floating",
            wrappingStyle: "wrapSquare" | "wrapTight" | "wrapThrough" | "wrapTopAndBottom" | "wrapNone" | null,
        } | null;
        isCropped: boolean;
    }

    interface Table {
        type: "table";
        children: Array<DocumentElement>;
        styleId: string | null;
        styleName: string | null;
        isBordered: boolean;
    }

    interface TableRow {
        type: "tableRow";
        children: Array<DocumentElement>;
        isHeader: boolean;
    }

    interface TableCell {
        type: "tableCell";
        children: Array<DocumentElement>;
        colSpan: number;
        rowSpan: number;
        width: number | null;
        bgColor: string | null;
    }

    interface Break {
        type: "break";
        breakType: "line" | "page" | "column";
    }

    interface BookmarkStart {
        type: "bookmarkStart";
        name: string;
    }

    interface Ins {
        type: "ins";
        children: Array<DocumentElement>;
        authorName: string | undefined;
        changeId: string | undefined;
        date: string | undefined;
    }

    interface Del {
        type: "del";
        children: Array<DocumentElement>;
        authorName: string | undefined;
        changeId: string | undefined;
        date: string | undefined;
    }
}

declare const mammoth: Mammoth;

export = mammoth;
