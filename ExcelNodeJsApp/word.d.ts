declare module Word {
    /**
     *
     * The Application object.
     *
     * [Api set: WordApi 1.3]
     */
    class Application extends OfficeExtension.ClientObject {
        /**
         *
         * Creates a new document by using a base64 encoded .docx file.
         *
         * @param base64File Optional. The base64 encoded .docx file. The default value is null.
         *
         * [Api set: WordApi 1.3]
         */
        createDocument(base64File?: string): Word.Document;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Create a new instance of Word.Application object
         */
        static newObject(context: OfficeExtension.ClientRequestContext): Word.Application;
        toJSON(): {};
    }
    /**
     *
     * Represents the body of a document or a section.
     *
     * [Api set: WordApi 1.1]
     */
    class Body extends OfficeExtension.ClientObject {
        private m_contentControls;
        private m_font;
        private m_inlinePictures;
        private m_lists;
        private m_paragraphs;
        private m_parentBody;
        private m_parentContentControl;
        private m_parentSection;
        private m_style;
        private m_styleBuiltIn;
        private m_tables;
        private m_text;
        private m_type;
        private m__ReferenceId;
        /**
         *
         * Gets the collection of rich text content control objects in the body. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        contentControls: Word.ContentControlCollection;
        /**
         *
         * Gets the text format of the body. Use this to get and set font name, size, color and other properties. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        font: Word.Font;
        /**
         *
         * Gets the collection of inlinePicture objects in the body. The collection does not include floating images. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        inlinePictures: Word.InlinePictureCollection;
        /**
         *
         * Gets the collection of list objects in the body. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        lists: Word.ListCollection;
        /**
         *
         * Gets the collection of paragraph objects in the body. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        paragraphs: Word.ParagraphCollection;
        /**
         *
         * Gets the parent body of the body. For example, a table cell body's parent body could be a header. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentBody: Word.Body;
        /**
         *
         * Gets the content control that contains the body. Returns a null object if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        parentContentControl: Word.ContentControl;
        /**
         *
         * Gets the parent section of the body. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentSection: Word.Section;
        /**
         *
         * Gets the collection of table objects in the body. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        tables: Word.TableCollection;
        /**
         *
         * Gets or sets the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
         *
         * [Api set: WordApi 1.1]
         */
        style: string;
        /**
         *
         * Gets or sets the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
         *
         * [Api set: WordApi 1.3]
         */
        styleBuiltIn: string;
        /**
         *
         * Gets the text of the body. Use the insertText method to insert text. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        text: string;
        /**
         *
         * Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        type: string;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        /**
         *
         * Clears the contents of the body object. The user can perform the undo operation on the cleared content.
         *
         * [Api set: WordApi 1.1]
         */
        clear(): void;
        /**
         *
         * Gets the HTML representation of the body object.
         *
         * [Api set: WordApi 1.1]
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Gets the OOXML (Office Open XML) representation of the body object.
         *
         * [Api set: WordApi 1.1]
         */
        getOoxml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Gets the whole body, or the starting or ending point of the body, as a range.
         *
         * @param rangeLocation Optional. The range location can be 'Whole', 'Start', 'End', 'After' or 'Content'.
         *
         * [Api set: WordApi 1.3]
         */
        getRange(rangeLocation?: string): Word.Range;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Start' or 'End'.
         *
         * @param breakType Required. The break type to add to the body.
         * @param insertLocation Required. The value can be 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         */
        insertBreak(breakType: string, insertLocation: string): void;
        /**
         *
         * Wraps the body object with a Rich Text content control.
         *
         * [Api set: WordApi 1.1]
         */
        insertContentControl(): Word.ContentControl;
        /**
         *
         * Inserts a document into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param base64File Required. The base64 encoded content of a .docx file.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         */
        insertFileFromBase64(base64File: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param html Required. The HTML to be inserted in the document.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         */
        insertHtml(html: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a picture into the body at the specified location. The insertLocation value can be 'Start' or 'End'.
         *
         * @param base64EncodedImage Required. The base64 encoded image to be inserted in the body.
         * @param insertLocation Required. The value can be 'Start' or 'End'.
         *
         * [Api set: WordApi 1.2]
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string): Word.InlinePicture;
        /**
         *
         * Inserts OOXML at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param ooxml Required. The OOXML to be inserted.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         */
        insertOoxml(ooxml: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Start' or 'End'.
         *
         * @param paragraphText Required. The paragraph text to be inserted.
         * @param insertLocation Required. The value can be 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         */
        insertParagraph(paragraphText: string, insertLocation: string): Word.Paragraph;
        /**
         *
         * Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Start' or 'End'.
         *
         * @param rowCount Required. The number of rows in the table.
         * @param columnCount Required. The number of columns in the table.
         * @param insertLocation Required. The value can be 'Start' or 'End'.
         * @param values Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         *
         * [Api set: WordApi 1.3]
         */
        insertTable(rowCount: number, columnCount: number, insertLocation: string, values?: Array<Array<string>>): Word.Table;
        /**
         *
         * Inserts text into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param text Required. Text to be inserted.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         */
        insertText(text: string, insertLocation: string): Word.Range;
        /**
         *
         * Performs a search with the specified searchOptions on the scope of the body object. The search results are a collection of range objects.
         *
         * @param searchText Required. The search text.
         * @param searchOptions Optional. Options for the search.
         *
         * [Api set: WordApi 1.1]
         */
        search(searchText: string, searchOptions?: Word.SearchOptions | {
            ignorePunct?: boolean;
            ignoreSpace?: boolean;
            matchCase?: boolean;
            matchPrefix?: boolean;
            matchSuffix?: boolean;
            matchWholeWord?: boolean;
            matchWildcards?: boolean;
        }): Word.RangeCollection;
        /**
         *
         * Selects the body and navigates the Word UI to it.
         *
         * @param selectionMode Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.
         *
         * [Api set: WordApi 1.1]
         */
        select(selectionMode?: string): void;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.Body;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.Body;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.Body;
        toJSON(): {
            "font": Font;
            "style": string;
            "styleBuiltIn": string;
            "text": string;
            "type": string;
        };
    }
    /**
     *
     * Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
     *
     * [Api set: WordApi 1.1]
     */
    class ContentControl extends OfficeExtension.ClientObject {
        private m_appearance;
        private m_cannotDelete;
        private m_cannotEdit;
        private m_color;
        private m_contentControls;
        private m_font;
        private m_id;
        private m_inlinePictures;
        private m_lists;
        private m_paragraphs;
        private m_parentBody;
        private m_parentContentControl;
        private m_parentTable;
        private m_parentTableCell;
        private m_placeholderText;
        private m_removeWhenEdited;
        private m_style;
        private m_styleBuiltIn;
        private m_subtype;
        private m_tables;
        private m_tag;
        private m_text;
        private m_title;
        private m_type;
        private m__ReferenceId;
        /**
         *
         * Gets the collection of content control objects in the content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        contentControls: Word.ContentControlCollection;
        /**
         *
         * Gets the text format of the content control. Use this to get and set font name, size, color, and other properties. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        font: Word.Font;
        /**
         *
         * Gets the collection of inlinePicture objects in the content control. The collection does not include floating images. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        inlinePictures: Word.InlinePictureCollection;
        /**
         *
         * Gets the collection of list objects in the content control. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        lists: Word.ListCollection;
        /**
         *
         * Get the collection of paragraph objects in the content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        paragraphs: Word.ParagraphCollection;
        /**
         *
         * Gets the parent body of the content control. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentBody: Word.Body;
        /**
         *
         * Gets the content control that contains the content control. Returns a null object if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        parentContentControl: Word.ContentControl;
        /**
         *
         * Gets the table that contains the content control. Returns a null object if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTable: Word.Table;
        /**
         *
         * Gets the table cell that contains the content control. Returns a null object if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTableCell: Word.TableCell;
        /**
         *
         * Gets the collection of table objects in the content control. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        tables: Word.TableCollection;
        /**
         *
         * Gets or sets the appearance of the content control. The value can be 'boundingBox', 'tags' or 'hidden'.
         *
         * [Api set: WordApi 1.1]
         */
        appearance: string;
        /**
         *
         * Gets or sets a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.
         *
         * [Api set: WordApi 1.1]
         */
        cannotDelete: boolean;
        /**
         *
         * Gets or sets a value that indicates whether the user can edit the contents of the content control.
         *
         * [Api set: WordApi 1.1]
         */
        cannotEdit: boolean;
        /**
         *
         * Gets or sets the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.
         *
         * [Api set: WordApi 1.1]
         */
        color: string;
        /**
         *
         * Gets an integer that represents the content control identifier. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        id: number;
        /**
         *
         * Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
         *
         * [Api set: WordApi 1.1]
         */
        placeholderText: string;
        /**
         *
         * Gets or sets a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.
         *
         * [Api set: WordApi 1.1]
         */
        removeWhenEdited: boolean;
        /**
         *
         * Gets or sets the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
         *
         * [Api set: WordApi 1.1]
         */
        style: string;
        /**
         *
         * Gets or sets the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
         *
         * [Api set: WordApi 1.3]
         */
        styleBuiltIn: string;
        /**
         *
         * Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        subtype: string;
        /**
         *
         * Gets or sets a tag to identify a content control.
         *
         * [Api set: WordApi 1.1]
         */
        tag: string;
        /**
         *
         * Gets the text of the content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        text: string;
        /**
         *
         * Gets or sets the title for a content control.
         *
         * [Api set: WordApi 1.1]
         */
        title: string;
        /**
         *
         * Gets the content control type. Only rich text content controls are supported currently. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        type: string;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        /**
         *
         * Clears the contents of the content control. The user can perform the undo operation on the cleared content.
         *
         * [Api set: WordApi 1.1]
         */
        clear(): void;
        /**
         *
         * Deletes the content control and its content. If keepContent is set to true, the content is not deleted.
         *
         * @param keepContent Required. Indicates whether the content should be deleted with the content control. If keepContent is set to true, the content is not deleted.
         *
         * [Api set: WordApi 1.1]
         */
        delete(keepContent: boolean): void;
        /**
         *
         * Gets the HTML representation of the content control object.
         *
         * [Api set: WordApi 1.1]
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Gets the Office Open XML (OOXML) representation of the content control object.
         *
         * [Api set: WordApi 1.1]
         */
        getOoxml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Gets the whole content control, or the starting or ending point of the content control, as a range.
         *
         * @param rangeLocation Optional. The range location can be 'Whole', 'Before', 'Start', 'End', 'After' or 'Content'.
         *
         * [Api set: WordApi 1.3]
         */
        getRange(rangeLocation?: string): Word.Range;
        /**
         *
         * Gets the text ranges in the content control by using punctuation marks and/or other ending marks.
         *
         * @param endingMarks Required. The punctuation marks and/or other ending marks as an array of strings.
         * @param trimSpacing Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         *
         * [Api set: WordApi 1.3]
         */
        getTextRanges(endingMarks: Array<string>, trimSpacing?: boolean): Word.RangeCollection;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Start', 'End', 'Before' or 'After'. This method cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
         *
         * @param breakType Required. Type of break.
         * @param insertLocation Required. The value can be 'Start', 'End', 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         */
        insertBreak(breakType: string, insertLocation: string): void;
        /**
         *
         * Inserts a document into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param base64File Required. The base64 encoded content of a .docx file.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         *
         * [Api set: WordApi 1.1]
         */
        insertFileFromBase64(base64File: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts HTML into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param html Required. The HTML to be inserted in to the content control.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         *
         * [Api set: WordApi 1.1]
         */
        insertHtml(html: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts an inline picture into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param base64EncodedImage Required. The base64 encoded image to be inserted in the content control.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         *
         * [Api set: WordApi 1.2]
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string): Word.InlinePicture;
        /**
         *
         * Inserts OOXML into the content control at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param ooxml Required. The OOXML to be inserted in to the content control.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         *
         * [Api set: WordApi 1.1]
         */
        insertOoxml(ooxml: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.
         *
         * @param paragraphText Required. The paragrph text to be inserted.
         * @param insertLocation Required. The value can be 'Start', 'End', 'Before' or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
         *
         * [Api set: WordApi 1.1]
         */
        insertParagraph(paragraphText: string, insertLocation: string): Word.Paragraph;
        /**
         *
         * Inserts a table with the specified number of rows and columns into, or next to, a content control. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.
         *
         * @param rowCount Required. The number of rows in the table.
         * @param columnCount Required. The number of columns in the table.
         * @param insertLocation Required. The value can be 'Start', 'End', 'Before' or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
         * @param values Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         *
         * [Api set: WordApi 1.3]
         */
        insertTable(rowCount: number, columnCount: number, insertLocation: string, values?: Array<Array<string>>): Word.Table;
        /**
         *
         * Inserts text into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param text Required. The text to be inserted in to the content control.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         *
         * [Api set: WordApi 1.1]
         */
        insertText(text: string, insertLocation: string): Word.Range;
        /**
         *
         * Performs a search with the specified searchOptions on the scope of the content control object. The search results are a collection of range objects.
         *
         * @param searchText Required. The search text.
         * @param searchOptions Optional. Options for the search.
         *
         * [Api set: WordApi 1.1]
         */
        search(searchText: string, searchOptions?: Word.SearchOptions | {
            ignorePunct?: boolean;
            ignoreSpace?: boolean;
            matchCase?: boolean;
            matchPrefix?: boolean;
            matchSuffix?: boolean;
            matchWholeWord?: boolean;
            matchWildcards?: boolean;
        }): Word.RangeCollection;
        /**
         *
         * Selects the content control. This causes Word to scroll to the selection.
         *
         * @param selectionMode Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.
         *
         * [Api set: WordApi 1.1]
         */
        select(selectionMode?: string): void;
        /**
         *
         * Splits the content control into child ranges by using delimiters.
         *
         * @param delimiters Required. The delimiters as an array of strings.
         * @param multiParagraphs Optional. Indicates whether a returned child range can cover multiple paragraphs. Default is false which indicates that the paragraph boundaries are also used as delimiters.
         * @param trimDelimiters Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.
         * @param trimSpacing Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         *
         * [Api set: WordApi 1.3]
         */
        split(delimiters: Array<string>, multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean): Word.RangeCollection;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.ContentControl;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.ContentControl;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.ContentControl;
        toJSON(): {
            "appearance": string;
            "cannotDelete": boolean;
            "cannotEdit": boolean;
            "color": string;
            "font": Font;
            "id": number;
            "placeholderText": string;
            "removeWhenEdited": boolean;
            "style": string;
            "styleBuiltIn": string;
            "subtype": string;
            "tag": string;
            "text": string;
            "title": string;
            "type": string;
        };
    }
    /**
     *
     * Contains a collection of [contentControl](contentControl.md) objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
     *
     * [Api set: WordApi 1.1]
     */
    class ContentControlCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Word.ContentControl>;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        /**
         *
         * Gets a content control by its identifier.
         *
         * @param id Required. A content control identifier.
         *
         * [Api set: WordApi 1.1]
         */
        getById(id: number): Word.ContentControl;
        /**
         *
         * Gets the content controls that have the specified tag.
         *
         * @param tag Required. A tag set on a content control.
         *
         * [Api set: WordApi 1.1]
         */
        getByTag(tag: string): Word.ContentControlCollection;
        /**
         *
         * Gets the content controls that have the specified title.
         *
         * @param title Required. The title of a content control.
         *
         * [Api set: WordApi 1.1]
         */
        getByTitle(title: string): Word.ContentControlCollection;
        /**
         *
         * Gets the content controls that have the specified types and/or subtypes.
         *
         * @param types Required. An array of content control types and/or subtypes.
         *
         * [Api set: WordApi 1.3]
         */
        getByTypes(types: Array<string>): Word.ContentControlCollection;
        /**
         *
         * Gets the first content control in this collection.
         *
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.ContentControl;
        /**
         *
         * Gets a content control by its index in the collection.
         *
         * @param index The index.
         *
         * [Api set: WordApi 1.1]
         */
        getItem(index: number): Word.ContentControl;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.ContentControlCollection;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.ContentControlCollection;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.ContentControlCollection;
        toJSON(): {};
    }
    /**
     *
     * Represents a custom property.
     *
     * [Api set: WordApi 1.3]
     */
    class CustomProperty extends OfficeExtension.ClientObject {
        private m_key;
        private m_type;
        private m_value;
        private m__ReferenceId;
        /**
         *
         * Gets the key of the custom property. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        key: string;
        /**
         *
         * Gets the value type of the custom property. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        type: string;
        /**
         *
         * Gets or sets the value of the custom property.
         *
         * [Api set: WordApi 1.3]
         */
        value: any;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        /**
         *
         * Deletes the custom property.
         *
         * [Api set: WordApi 1.3]
         */
        delete(): void;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.CustomProperty;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.CustomProperty;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.CustomProperty;
        toJSON(): {
            "key": string;
            "type": string;
            "value": any;
        };
    }
    /**
     *
     * Contains the collection of [customProperty](customProperty.md) objects.
     *
     * [Api set: WordApi 1.3]
     */
    class CustomPropertyCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Word.CustomProperty>;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        /**
         *
         * Deletes all custom properties in this collection.
         *
         * [Api set: WordApi 1.3]
         */
        deleteAll(): void;
        /**
         *
         * Gets the count of custom properties.
         *
         * [Api set: WordApi 1.3]
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         *
         * Gets a custom property object by its key, which is case-insensitive.
         *
         * @param key The key that identifies the custom property object.
         *
         * [Api set: WordApi 1.3]
         */
        getItem(key: string): Word.CustomProperty;
        /**
         *
         * Creates or sets a custom property.
         *
         * @param key Required. The custom property's key, which is case-insensitive.
         * @param value Required. The custom property's value.
         *
         * [Api set: WordApi 1.3]
         */
        set(key: string, value: any): Word.CustomProperty;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.CustomPropertyCollection;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.CustomPropertyCollection;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.CustomPropertyCollection;
        toJSON(): {};
    }
    /**
     *
     * The Document object is the top level object. A Document object contains one or more sections, content controls, and the body that contains the contents of the document.
     *
     * [Api set: WordApi 1.1]
     */
    class Document extends OfficeExtension.ClientObject {
        private m_body;
        private m_contentControls;
        private m_properties;
        private m_saved;
        private m_sections;
        private m_settings;
        private m__ReferenceId;
        /**
         *
         * Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        body: Word.Body;
        /**
         *
         * Gets the collection of content control objects in the current document. This includes content controls in the body of the document, headers, footers, textboxes, etc.. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        contentControls: Word.ContentControlCollection;
        /**
         *
         * Gets the properties of the current document. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        properties: Word.DocumentProperties;
        /**
         *
         * Gets the collection of section objects in the document. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        sections: Word.SectionCollection;
        /**
         *
         * Gets the add-in's settings in the current document. Read-only.
         *
         * [Api set: WordApi 1.4]
         */
        settings: Word.SettingCollection;
        /**
         *
         * Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        saved: boolean;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        /**
         *
         * Deletes a bookmark, if exists, from this document.
         *
         * @param name Required. The bookmark name, which is case-insensitive.
         *
         * [Api set: WordApi 1.4]
         */
        deleteBookmark(name: string): void;
        /**
         *
         * Gets a bookmark's range. Returns a null object if the bookmark does not exist.
         *
         * @param name Required. The bookmark name, which is case-insensitive.
         *
         * [Api set: WordApi 1.4]
         */
        getBookmarkRange(name: string): Word.Range;
        /**
         *
         * Gets the current selection of the document. Multiple selections are not supported.
         *
         * [Api set: WordApi 1.1]
         */
        getSelection(): Word.Range;
        /**
         *
         * Open the document.
         *
         * [Api set: WordApi 1.3]
         */
        open(): void;
        /**
         *
         * Saves the document. This will use the Word default file naming convention if the document has not been saved before.
         *
         * [Api set: WordApi 1.1]
         */
        save(): void;
        _GetObjectByReferenceId(referenceId: string): OfficeExtension.ClientResult<any>;
        _GetObjectTypeNameByReferenceId(referenceId: string): OfficeExtension.ClientResult<string>;
        _KeepReference(): void;
        _RemoveAllReferences(): void;
        _RemoveReference(referenceId: string): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.Document;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.Document;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.Document;
        toJSON(): {
            "body": Body;
            "properties": DocumentProperties;
            "saved": boolean;
        };
    }
    /**
     *
     * Represents document properties.
     *
     * [Api set: WordApi 1.3]
     */
    class DocumentProperties extends OfficeExtension.ClientObject {
        private m_applicationName;
        private m_author;
        private m_category;
        private m_comments;
        private m_company;
        private m_creationDate;
        private m_customProperties;
        private m_format;
        private m_hyperlinkBase;
        private m_keywords;
        private m_lastAuthor;
        private m_lastPrintDate;
        private m_lastSaveTime;
        private m_manager;
        private m_numberOfBytes;
        private m_numberOfCharacters;
        private m_numberOfCharactersWithSpaces;
        private m_numberOfHiddenSlides;
        private m_numberOfLines;
        private m_numberOfMultimediaClips;
        private m_numberOfNotes;
        private m_numberOfPages;
        private m_numberOfParagraphs;
        private m_numberOfSlides;
        private m_numberOfWords;
        private m_revisionNumber;
        private m_security;
        private m_subject;
        private m_template;
        private m_title;
        private m_totalEditingTime;
        private m__ReferenceId;
        /**
         *
         * Gets the collection of custom properties of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        customProperties: Word.CustomPropertyCollection;
        /**
         *
         * Gets the application name of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        applicationName: string;
        /**
         *
         * Gets or sets the author of the document.
         *
         * [Api set: WordApi 1.3]
         */
        author: string;
        /**
         *
         * Gets or sets the category of the document.
         *
         * [Api set: WordApi 1.3]
         */
        category: string;
        /**
         *
         * Gets or sets the comments of the document.
         *
         * [Api set: WordApi 1.3]
         */
        comments: string;
        /**
         *
         * Gets or sets the company of the document.
         *
         * [Api set: WordApi 1.3]
         */
        company: string;
        /**
         *
         * Gets the creation date of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        creationDate: Date;
        /**
         *
         * Gets or sets the format of the document.
         *
         * [Api set: WordApi 1.3]
         */
        format: string;
        /**
         *
         * Gets or sets the hyperlink base of the document.
         *
         * [Api set: WordApi 1.3]
         */
        hyperlinkBase: string;
        /**
         *
         * Gets or sets the keywords of the document.
         *
         * [Api set: WordApi 1.3]
         */
        keywords: string;
        /**
         *
         * Gets or sets the last author of the document.
         *
         * [Api set: WordApi 1.3]
         */
        lastAuthor: string;
        /**
         *
         * Gets the last print date of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        lastPrintDate: Date;
        /**
         *
         * Gets the last save time of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        lastSaveTime: Date;
        /**
         *
         * Gets or sets the manager of the document.
         *
         * [Api set: WordApi 1.3]
         */
        manager: string;
        /**
         *
         * Gets the number of bytes of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        numberOfBytes: number;
        /**
         *
         * Gets the number of characters of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        numberOfCharacters: number;
        /**
         *
         * Gets the number of characters with spaces of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        numberOfCharactersWithSpaces: number;
        /**
         *
         * Gets the number of hidden slides of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        numberOfHiddenSlides: number;
        /**
         *
         * Gets the number of lines of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        numberOfLines: number;
        /**
         *
         * Gets the number of multimedia clips of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        numberOfMultimediaClips: number;
        /**
         *
         * Gets the number of notes of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        numberOfNotes: number;
        /**
         *
         * Gets the number of pages of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        numberOfPages: number;
        /**
         *
         * Gets the number of paragraphs of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        numberOfParagraphs: number;
        /**
         *
         * Gets the number of slides of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        numberOfSlides: number;
        /**
         *
         * Gets the number of words of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        numberOfWords: number;
        /**
         *
         * Gets the revision number of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        revisionNumber: string;
        /**
         *
         * Gets the security of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        security: number;
        /**
         *
         * Gets or sets the subject of the document.
         *
         * [Api set: WordApi 1.3]
         */
        subject: string;
        /**
         *
         * Gets the template of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        template: string;
        /**
         *
         * Gets or sets the title of the document.
         *
         * [Api set: WordApi 1.3]
         */
        title: string;
        /**
         *
         * Gets the total editing time of the document in minutes. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        totalEditingTime: number;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.DocumentProperties;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.DocumentProperties;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.DocumentProperties;
        toJSON(): {
            "applicationName": string;
            "author": string;
            "category": string;
            "comments": string;
            "company": string;
            "creationDate": Date;
            "format": string;
            "hyperlinkBase": string;
            "keywords": string;
            "lastAuthor": string;
            "lastPrintDate": Date;
            "lastSaveTime": Date;
            "manager": string;
            "numberOfBytes": number;
            "numberOfCharacters": number;
            "numberOfCharactersWithSpaces": number;
            "numberOfHiddenSlides": number;
            "numberOfLines": number;
            "numberOfMultimediaClips": number;
            "numberOfNotes": number;
            "numberOfPages": number;
            "numberOfParagraphs": number;
            "numberOfSlides": number;
            "numberOfWords": number;
            "revisionNumber": string;
            "security": number;
            "subject": string;
            "template": string;
            "title": string;
            "totalEditingTime": number;
        };
    }
    /**
     *
     * Represents a font.
     *
     * [Api set: WordApi 1.1]
     */
    class Font extends OfficeExtension.ClientObject {
        private m_bold;
        private m_color;
        private m_doubleStrikeThrough;
        private m_highlightColor;
        private m_italic;
        private m_name;
        private m_size;
        private m_strikeThrough;
        private m_subscript;
        private m_superscript;
        private m_underline;
        private m__ReferenceId;
        /**
         *
         * Gets or sets a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false.
         *
         * [Api set: WordApi 1.1]
         */
        bold: boolean;
        /**
         *
         * Gets or sets the color for the specified font. You can provide the value in the '#RRGGBB' format or the color name.
         *
         * [Api set: WordApi 1.1]
         */
        color: string;
        /**
         *
         * Gets or sets a value that indicates whether the font has a double strike through. True if the font is formatted as double strikethrough text, otherwise, false.
         *
         * [Api set: WordApi 1.1]
         */
        doubleStrikeThrough: boolean;
        /**
         *
         * Gets or sets the highlight color for the specified font. You can provide the value as either in the '#RRGGBB' format or the color name.
         *
         * [Api set: WordApi 1.1]
         */
        highlightColor: string;
        /**
         *
         * Gets or sets a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false.
         *
         * [Api set: WordApi 1.1]
         */
        italic: boolean;
        /**
         *
         * Gets or sets a value that represents the name of the font.
         *
         * [Api set: WordApi 1.1]
         */
        name: string;
        /**
         *
         * Gets or sets a value that represents the font size in points.
         *
         * [Api set: WordApi 1.1]
         */
        size: number;
        /**
         *
         * Gets or sets a value that indicates whether the font has a strike through. True if the font is formatted as strikethrough text, otherwise, false.
         *
         * [Api set: WordApi 1.1]
         */
        strikeThrough: boolean;
        /**
         *
         * Gets or sets a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false.
         *
         * [Api set: WordApi 1.1]
         */
        subscript: boolean;
        /**
         *
         * Gets or sets a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false.
         *
         * [Api set: WordApi 1.1]
         */
        superscript: boolean;
        /**
         *
         * Gets or sets a value that indicates the font's underline type. 'None' if the font is not underlined.
         *
         * [Api set: WordApi 1.1]
         */
        underline: string;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.Font;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.Font;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.Font;
        toJSON(): {
            "bold": boolean;
            "color": string;
            "doubleStrikeThrough": boolean;
            "highlightColor": string;
            "italic": boolean;
            "name": string;
            "size": number;
            "strikeThrough": boolean;
            "subscript": boolean;
            "superscript": boolean;
            "underline": string;
        };
    }
    /**
     *
     * Represents an inline picture.
     *
     * [Api set: WordApi 1.1]
     */
    class InlinePicture extends OfficeExtension.ClientObject {
        private m_altTextDescription;
        private m_altTextTitle;
        private m_height;
        private m_hyperlink;
        private m_imageFormat;
        private m_lockAspectRatio;
        private m_paragraph;
        private m_parentContentControl;
        private m_parentTable;
        private m_parentTableCell;
        private m_width;
        private m__Id;
        private m__ReferenceId;
        /**
         *
         * Gets the parent paragraph that contains the inline image. Read-only.
         *
         * [Api set: WordApi 1.2]
         */
        paragraph: Word.Paragraph;
        /**
         *
         * Gets the content control that contains the inline image. Returns a null object if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        parentContentControl: Word.ContentControl;
        /**
         *
         * Gets the table that contains the inline image. Returns a null object if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTable: Word.Table;
        /**
         *
         * Gets the table cell that contains the inline image. Returns a null object if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTableCell: Word.TableCell;
        /**
         *
         * Gets or sets a string that represents the alternative text associated with the inline image
         *
         * [Api set: WordApi 1.1]
         */
        altTextDescription: string;
        /**
         *
         * Gets or sets a string that contains the title for the inline image.
         *
         * [Api set: WordApi 1.1]
         */
        altTextTitle: string;
        /**
         *
         * Gets or sets a number that describes the height of the inline image.
         *
         * [Api set: WordApi 1.1]
         */
        height: number;
        /**
         *
         * Gets or sets a hyperlink on the image. Use a newline character ('\n') to separate the address part from the optional location part.
         *
         * [Api set: WordApi 1.1]
         */
        hyperlink: string;
        /**
         *
         * Gets the format of the inline image. Read-only.
         *
         * [Api set: WordApi 1.4]
         */
        imageFormat: string;
        /**
         *
         * Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.
         *
         * [Api set: WordApi 1.1]
         */
        lockAspectRatio: boolean;
        /**
         *
         * Gets or sets a number that describes the width of the inline image.
         *
         * [Api set: WordApi 1.1]
         */
        width: number;
        /**
         *
         * ID
         *
         * [Api set: WordApi]
         */
        _Id: number;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        /**
         *
         * Deletes the inline picture from the document.
         *
         * [Api set: WordApi 1.2]
         */
        delete(): void;
        /**
         *
         * Gets the base64 encoded string representation of the inline image.
         *
         * [Api set: WordApi 1.1]
         */
        getBase64ImageSrc(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Gets the next inline image.
         *
         * [Api set: WordApi 1.3]
         */
        getNext(): Word.InlinePicture;
        /**
         *
         * Gets the picture, or the starting or ending point of the picture, as a range.
         *
         * @param rangeLocation Optional. The range location can be 'Whole', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.3]
         */
        getRange(rangeLocation?: string): Word.Range;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.
         *
         * @param breakType Required. The break type to add.
         * @param insertLocation Required. The value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         */
        insertBreak(breakType: string, insertLocation: string): void;
        /**
         *
         * Wraps the inline picture with a rich text content control.
         *
         * [Api set: WordApi 1.1]
         */
        insertContentControl(): Word.ContentControl;
        /**
         *
         * Inserts a document at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * @param base64File Required. The base64 encoded content of a .docx file.
         * @param insertLocation Required. The value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         */
        insertFileFromBase64(base64File: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts HTML at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * @param html Required. The HTML to be inserted.
         * @param insertLocation Required. The value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         */
        insertHtml(html: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts an inline picture at the specified location. The insertLocation value can be 'Replace', 'Before' or 'After'.
         *
         * @param base64EncodedImage Required. The base64 encoded image to be inserted.
         * @param insertLocation Required. The value can be 'Replace', 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string): Word.InlinePicture;
        /**
         *
         * Inserts OOXML at the specified location.  The insertLocation value can be 'Before' or 'After'.
         *
         * @param ooxml Required. The OOXML to be inserted.
         * @param insertLocation Required. The value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         */
        insertOoxml(ooxml: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * @param paragraphText Required. The paragraph text to be inserted.
         * @param insertLocation Required. The value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         */
        insertParagraph(paragraphText: string, insertLocation: string): Word.Paragraph;
        /**
         *
         * Inserts text at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * @param text Required. Text to be inserted.
         * @param insertLocation Required. The value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         */
        insertText(text: string, insertLocation: string): Word.Range;
        /**
         *
         * Selects the inline picture. This causes Word to scroll to the selection.
         *
         * @param selectionMode Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.
         *
         * [Api set: WordApi 1.2]
         */
        select(selectionMode?: string): void;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.InlinePicture;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.InlinePicture;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.InlinePicture;
        toJSON(): {
            "altTextDescription": string;
            "altTextTitle": string;
            "height": number;
            "hyperlink": string;
            "imageFormat": string;
            "lockAspectRatio": boolean;
            "width": number;
        };
    }
    /**
     *
     * Contains a collection of [inlinePicture](inlinePicture.md) objects.
     *
     * [Api set: WordApi 1.1]
     */
    class InlinePictureCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Word.InlinePicture>;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        /**
         *
         * Gets the first inline image in this collection.
         *
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.InlinePicture;
        /**
         *
         * Gets an inline picture object by its index in the collection.
         *
         * @param index A number that identifies the index location of an inline picture object.
         *
         * [Api set: WordApi 1.1]
         */
        _GetItem(index: number): Word.InlinePicture;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.InlinePictureCollection;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.InlinePictureCollection;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.InlinePictureCollection;
        toJSON(): {};
    }
    /**
     *
     * Contains a collection of [paragraph](paragraph.md) objects.
     *
     * [Api set: WordApi 1.3]
     */
    class List extends OfficeExtension.ClientObject {
        private m_id;
        private m_levelExistences;
        private m_levelTypes;
        private m_paragraphs;
        private m__ReferenceId;
        /**
         *
         * Gets paragraphs in the list. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        paragraphs: Word.ParagraphCollection;
        /**
         *
         * Gets the list's id.
         *
         * [Api set: WordApi 1.3]
         */
        id: number;
        /**
         *
         * Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        levelExistences: Array<boolean>;
        /**
         *
         * Gets all 9 level types in the list. Each type can be 'Bullet', 'Number' or 'Picture'. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        levelTypes: Array<string>;
        _ReferenceId: string;
        /**
         *
         * Gets the font of the bullet, number or picture at the specified level in the list.
         *
         * @param level Required. The level in the list.
         *
         * [Api set: WordApi 1.4]
         */
        getLevelFont(level: number): Word.Font;
        /**
         *
         * Gets the paragraphs that occur at the specified level in the list.
         *
         * @param level Required. The level in the list.
         *
         * [Api set: WordApi 1.3]
         */
        getLevelParagraphs(level: number): Word.ParagraphCollection;
        /**
         *
         * Gets the base64 encoded string representation of the picture at the specified level in the list.
         *
         * @param level Required. The level in the list.
         *
         * [Api set: WordApi 1.4]
         */
        getLevelPicture(level: number): OfficeExtension.ClientResult<string>;
        /**
         *
         * Gets the bullet, number or picture at the specified level as a string.
         *
         * @param level Required. The level in the list.
         *
         * [Api set: WordApi 1.3]
         */
        getLevelString(level: number): OfficeExtension.ClientResult<string>;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.
         *
         * @param paragraphText Required. The paragraph text to be inserted.
         * @param insertLocation Required. The value can be 'Start', 'End', 'Before' or 'After'.
         *
         * [Api set: WordApi 1.3]
         */
        insertParagraph(paragraphText: string, insertLocation: string): Word.Paragraph;
        /**
         *
         * Resets the font of the bullet, number or picture at the specified level in the list.
         *
         * @param level Required. The level in the list.
         * @param resetFontName Optional. Indicates whether to reset the font name. Default is false that indicates the font name is kept unchanged.
         *
         * [Api set: WordApi 1.4]
         */
        resetLevelFont(level: number, resetFontName?: boolean): void;
        /**
         *
         * Sets the alignment of the bullet, number or picture at the specified level in the list.
         *
         * @param level Required. The level in the list.
         * @param alignment Required. The level alignment that can be 'left', 'centered' or 'right'.
         *
         * [Api set: WordApi 1.3]
         */
        setLevelAlignment(level: number, alignment: string): void;
        /**
         *
         * Sets the bullet format at the specified level in the list. If the bullet is 'Custom', the charCode is required.
         *
         * @param level Required. The level in the list.
         * @param listBullet Required. The bullet.
         * @param charCode Optional. The bullet character's code value. Used only if the bullet is 'Custom'.
         * @param fontName Optional. The bullet's font name. Used only if the bullet is 'Custom'.
         *
         * [Api set: WordApi 1.3]
         */
        setLevelBullet(level: number, listBullet: string, charCode?: number, fontName?: string): void;
        /**
         *
         * Sets the two indents of the specified level in the list.
         *
         * @param level Required. The level in the list.
         * @param textIndent Required. The text indent in points. It is the same as paragraph left indent.
         * @param textIndent Required. The relative indent, in points, of the bullet, number or picture. It is the same as paragraph first line indent.
         *
         * [Api set: WordApi 1.3]
         */
        setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number): void;
        /**
         *
         * Sets the numbering format at the specified level in the list.
         *
         * @param level Required. The level in the list.
         * @param listNumbering Required. The ordinal format.
         * @param formatString Optional. The numbering string format defined as an array of strings and/or integers. Each integer is a level of number type that is higher than or equal to this level. For example, an array of ["(", level - 1, ".", level, ")"] can define the format of "(2.c)", where 2 is the parent's item number and c is this level's item number.
         *
         * [Api set: WordApi 1.3]
         */
        setLevelNumbering(level: number, listNumbering: string, formatString?: Array<any>): void;
        /**
         *
         * Sets the picture at the specified level in the list.
         *
         * @param level Required. The level in the list.
         * @param base64EncodedImage Optional. The base64 encoded image to be set. If not given, the default picture is set.
         *
         * [Api set: WordApi 1.4]
         */
        setLevelPicture(level: number, base64EncodedImage?: string): void;
        /**
         *
         * Sets the starting number at the specified level in the list. Default value is 1.
         *
         * @param level Required. The level in the list.
         * @param startingNumber Required. The number to start with.
         *
         * [Api set: WordApi 1.3]
         */
        setLevelStartingNumber(level: number, startingNumber: number): void;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.List;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.List;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.List;
        toJSON(): {
            "id": number;
            "levelExistences": boolean[];
            "levelTypes": string[];
        };
    }
    /**
     *
     * Contains a collection of [list](list.md) objects.
     *
     * [Api set: WordApi 1.3]
     */
    class ListCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Word.List>;
        _ReferenceId: string;
        /**
         *
         * Gets a list by its identifier.
         *
         * @param id Required. A list identifier.
         *
         * [Api set: WordApi 1.3]
         */
        getById(id: number): Word.List;
        /**
         *
         * Gets the first list in this collection.
         *
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.List;
        /**
         *
         * Gets a list object by its index in the collection.
         *
         * @param index A number that identifies the index location of a list object.
         *
         * [Api set: WordApi 1.3]
         */
        getItem(index: number): Word.List;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.ListCollection;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.ListCollection;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.ListCollection;
        toJSON(): {};
    }
    /**
     *
     * Represents the paragraph list item format.
     *
     * [Api set: WordApi 1.3]
     */
    class ListItem extends OfficeExtension.ClientObject {
        private m_level;
        private m_listString;
        private m_siblingIndex;
        private m__ReferenceId;
        /**
         *
         * Gets or sets the level of the item in the list.
         *
         * [Api set: WordApi 1.3]
         */
        level: number;
        /**
         *
         * Gets the list item bullet, number or picture as a string. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        listString: string;
        /**
         *
         * Gets the list item order number in relation to its siblings. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        siblingIndex: number;
        _ReferenceId: string;
        /**
         *
         * Gets the list item parent, or the closest ancestor if the parent does not exist.
         *
         * @param parentOnly Optional. Specified only the list item's parent will be returned. The default is false that specifies to get the lowest ancestor.
         *
         * [Api set: WordApi 1.3]
         */
        getAncestor(parentOnly?: boolean): Word.Paragraph;
        /**
         *
         * Gets all descendant list items of the list item.
         *
         * @param directChildrenOnly Optional. Specified only the list item's direct children will be returned. The default is false that indicates to get all descendant items.
         *
         * [Api set: WordApi 1.3]
         */
        getDescendants(directChildrenOnly?: boolean): Word.ParagraphCollection;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.ListItem;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.ListItem;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.ListItem;
        toJSON(): {
            "level": number;
            "listString": string;
            "siblingIndex": number;
        };
    }
    /**
     *
     * Represents a single paragraph in a selection, range, content control, or document body.
     *
     * [Api set: WordApi 1.1]
     */
    class Paragraph extends OfficeExtension.ClientObject {
        private m_alignment;
        private m_contentControls;
        private m_firstLineIndent;
        private m_font;
        private m_inlinePictures;
        private m_isLastParagraph;
        private m_isListItem;
        private m_leftIndent;
        private m_lineSpacing;
        private m_lineUnitAfter;
        private m_lineUnitBefore;
        private m_list;
        private m_listItem;
        private m_outlineLevel;
        private m_parentBody;
        private m_parentContentControl;
        private m_parentTable;
        private m_parentTableCell;
        private m_rightIndent;
        private m_spaceAfter;
        private m_spaceBefore;
        private m_style;
        private m_styleBuiltIn;
        private m_tableNestingLevel;
        private m_text;
        private m__Id;
        private m__ReferenceId;
        /**
         *
         * Gets the collection of content control objects in the paragraph. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        contentControls: Word.ContentControlCollection;
        /**
         *
         * Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        font: Word.Font;
        /**
         *
         * Gets the collection of inlinePicture objects in the paragraph. The collection does not include floating images. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        inlinePictures: Word.InlinePictureCollection;
        /**
         *
         * Gets the List to which this paragraph belongs. Returns a null object if the paragraph is not in a list. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        list: Word.List;
        /**
         *
         * Gets the ListItem for the paragraph. Returns a null object if the paragraph is not part of a list. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        listItem: Word.ListItem;
        /**
         *
         * Gets the parent body of the paragraph. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentBody: Word.Body;
        /**
         *
         * Gets the content control that contains the paragraph. Returns a null object if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        parentContentControl: Word.ContentControl;
        /**
         *
         * Gets the table that contains the paragraph. Returns a null object if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTable: Word.Table;
        /**
         *
         * Gets the table cell that contains the paragraph. Returns a null object if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTableCell: Word.TableCell;
        /**
         *
         * Gets or sets the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.
         *
         * [Api set: WordApi 1.1]
         */
        alignment: string;
        /**
         *
         * Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
         *
         * [Api set: WordApi 1.1]
         */
        firstLineIndent: number;
        /**
         *
         * Indicates the paragraph is the last one inside its parent body. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        isLastParagraph: boolean;
        /**
         *
         * Checks whether the paragraph is a list item. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        isListItem: boolean;
        /**
         *
         * Gets or sets the left indent value, in points, for the paragraph.
         *
         * [Api set: WordApi 1.1]
         */
        leftIndent: number;
        /**
         *
         * Gets or sets the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
         *
         * [Api set: WordApi 1.1]
         */
        lineSpacing: number;
        /**
         *
         * Gets or sets the amount of spacing, in grid lines. after the paragraph.
         *
         * [Api set: WordApi 1.1]
         */
        lineUnitAfter: number;
        /**
         *
         * Gets or sets the amount of spacing, in grid lines, before the paragraph.
         *
         * [Api set: WordApi 1.1]
         */
        lineUnitBefore: number;
        /**
         *
         * Gets or sets the outline level for the paragraph.
         *
         * [Api set: WordApi 1.1]
         */
        outlineLevel: number;
        /**
         *
         * Gets or sets the right indent value, in points, for the paragraph.
         *
         * [Api set: WordApi 1.1]
         */
        rightIndent: number;
        /**
         *
         * Gets or sets the spacing, in points, after the paragraph.
         *
         * [Api set: WordApi 1.1]
         */
        spaceAfter: number;
        /**
         *
         * Gets or sets the spacing, in points, before the paragraph.
         *
         * [Api set: WordApi 1.1]
         */
        spaceBefore: number;
        /**
         *
         * Gets or sets the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
         *
         * [Api set: WordApi 1.1]
         */
        style: string;
        /**
         *
         * Gets or sets the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
         *
         * [Api set: WordApi 1.3]
         */
        styleBuiltIn: string;
        /**
         *
         * Gets the level of the paragraph's table. It returns 0 if the paragraph is not in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        tableNestingLevel: number;
        /**
         *
         * Gets the text of the paragraph. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        text: string;
        /**
         *
         * ID
         *
         * [Api set: WordApi]
         */
        _Id: number;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        /**
         *
         * Lets the paragraph join an existing list at the specified level. Fails if the paragraph cannot join the list or if the paragraph is already a list item.
         *
         * @param listId Required. The ID of an existing list.
         * @param level Required. The level in the list.
         *
         * [Api set: WordApi 1.3]
         */
        attachToList(listId: number, level: number): Word.List;
        /**
         *
         * Clears the contents of the paragraph object. The user can perform the undo operation on the cleared content.
         *
         * [Api set: WordApi 1.1]
         */
        clear(): void;
        /**
         *
         * Deletes the paragraph and its content from the document.
         *
         * [Api set: WordApi 1.1]
         */
        delete(): void;
        /**
         *
         * Moves this paragraph out of its list, if the paragraph is a list item.
         *
         * [Api set: WordApi 1.3]
         */
        detachFromList(): void;
        /**
         *
         * Gets the HTML representation of the paragraph object.
         *
         * [Api set: WordApi 1.1]
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Gets the next paragraph.
         *
         * [Api set: WordApi 1.3]
         */
        getNext(): Word.Paragraph;
        /**
         *
         * Gets the Office Open XML (OOXML) representation of the paragraph object.
         *
         * [Api set: WordApi 1.1]
         */
        getOoxml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Gets the previous paragraph.
         *
         * [Api set: WordApi 1.3]
         */
        getPrevious(): Word.Paragraph;
        /**
         *
         * Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.
         *
         * @param rangeLocation Optional. The range location can be 'Whole', 'Start', 'End', 'After' or 'Content'.
         *
         * [Api set: WordApi 1.3]
         */
        getRange(rangeLocation?: string): Word.Range;
        /**
         *
         * Gets the text ranges in the paragraph by using punctuation marks and/or other ending marks.
         *
         * @param endingMarks Required. The punctuation marks and/or other ending marks as an array of strings.
         * @param trimSpacing Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         *
         * [Api set: WordApi 1.3]
         */
        getTextRanges(endingMarks: Array<string>, trimSpacing?: boolean): Word.RangeCollection;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.
         *
         * @param breakType Required. The break type to add to the document.
         * @param insertLocation Required. The value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         */
        insertBreak(breakType: string, insertLocation: string): void;
        /**
         *
         * Wraps the paragraph object with a rich text content control.
         *
         * [Api set: WordApi 1.1]
         */
        insertContentControl(): Word.ContentControl;
        /**
         *
         * Inserts a document into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param base64File Required. The base64 encoded content of a .docx file.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         */
        insertFileFromBase64(base64File: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts HTML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param html Required. The HTML to be inserted in the paragraph.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         */
        insertHtml(html: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a picture into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param base64EncodedImage Required. The base64 encoded image to be inserted.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string): Word.InlinePicture;
        /**
         *
         * Inserts OOXML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param ooxml Required. The OOXML to be inserted in the paragraph.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         */
        insertOoxml(ooxml: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * @param paragraphText Required. The paragraph text to be inserted.
         * @param insertLocation Required. The value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         */
        insertParagraph(paragraphText: string, insertLocation: string): Word.Paragraph;
        /**
         *
         * Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.
         *
         * @param rowCount Required. The number of rows in the table.
         * @param columnCount Required. The number of columns in the table.
         * @param insertLocation Required. The value can be 'Before' or 'After'.
         * @param values Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         *
         * [Api set: WordApi 1.3]
         */
        insertTable(rowCount: number, columnCount: number, insertLocation: string, values?: Array<Array<string>>): Word.Table;
        /**
         *
         * Inserts text into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * @param text Required. Text to be inserted.
         * @param insertLocation Required. The value can be 'Replace', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         */
        insertText(text: string, insertLocation: string): Word.Range;
        /**
         *
         * Performs a search with the specified searchOptions on the scope of the paragraph object. The search results are a collection of range objects.
         *
         * @param searchText Required. The search text.
         * @param searchOptions Optional. Options for the search.
         *
         * [Api set: WordApi 1.1]
         */
        search(searchText: string, searchOptions?: Word.SearchOptions | {
            ignorePunct?: boolean;
            ignoreSpace?: boolean;
            matchCase?: boolean;
            matchPrefix?: boolean;
            matchSuffix?: boolean;
            matchWholeWord?: boolean;
            matchWildcards?: boolean;
        }): Word.RangeCollection;
        /**
         *
         * Selects and navigates the Word UI to the paragraph.
         *
         * @param selectionMode Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.
         *
         * [Api set: WordApi 1.1]
         */
        select(selectionMode?: string): void;
        /**
         *
         * Splits the paragraph into child ranges by using delimiters.
         *
         * @param delimiters Required. The delimiters as an array of strings.
         * @param trimDelimiters Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.
         * @param trimSpacing Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         *
         * [Api set: WordApi 1.3]
         */
        split(delimiters: Array<string>, trimDelimiters?: boolean, trimSpacing?: boolean): Word.RangeCollection;
        /**
         *
         * Starts a new list with this paragraph. Fails if the paragraph is already a list item.
         *
         * [Api set: WordApi 1.3]
         */
        startNewList(): Word.List;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.Paragraph;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.Paragraph;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.Paragraph;
        toJSON(): {
            "alignment": string;
            "firstLineIndent": number;
            "font": Font;
            "isLastParagraph": boolean;
            "isListItem": boolean;
            "leftIndent": number;
            "lineSpacing": number;
            "lineUnitAfter": number;
            "lineUnitBefore": number;
            "listItem": ListItem;
            "outlineLevel": number;
            "rightIndent": number;
            "spaceAfter": number;
            "spaceBefore": number;
            "style": string;
            "styleBuiltIn": string;
            "tableNestingLevel": number;
            "text": string;
        };
    }
    /**
     *
     * Contains a collection of [paragraph](paragraph.md) objects.
     *
     * [Api set: WordApi 1.1]
     */
    class ParagraphCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Word.Paragraph>;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        /**
         *
         * Gets the first paragraph in this collection.
         *
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.Paragraph;
        /**
         *
         * Gets the last paragraph in this collection.
         *
         * [Api set: WordApi 1.3]
         */
        getLast(): Word.Paragraph;
        /**
         *
         * Gets a paragraph object by its index in the collection.
         *
         * @param index A number that identifies the index location of a paragraph object.
         *
         * [Api set: WordApi 1.1]
         */
        _GetItem(index: number): Word.Paragraph;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.ParagraphCollection;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.ParagraphCollection;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.ParagraphCollection;
        toJSON(): {};
    }
    /**
     *
     * Represents a contiguous area in a document.
     *
     * [Api set: WordApi 1.1]
     */
    class Range extends OfficeExtension.ClientObject {
        private m_contentControls;
        private m_font;
        private m_hyperlink;
        private m_inlinePictures;
        private m_isEmpty;
        private m_lists;
        private m_paragraphs;
        private m_parentBody;
        private m_parentContentControl;
        private m_parentTable;
        private m_parentTableCell;
        private m_style;
        private m_styleBuiltIn;
        private m_tables;
        private m_text;
        private m__Id;
        private m__ReferenceId;
        /**
         *
         * Gets the collection of content control objects in the range. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        contentControls: Word.ContentControlCollection;
        /**
         *
         * Gets the text format of the range. Use this to get and set font name, size, color, and other properties. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        font: Word.Font;
        /**
         *
         * Gets the collection of inline picture objects in the range. Read-only.
         *
         * [Api set: WordApi 1.2]
         */
        inlinePictures: Word.InlinePictureCollection;
        /**
         *
         * Gets the collection of list objects in the range. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        lists: Word.ListCollection;
        /**
         *
         * Gets the collection of paragraph objects in the range. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        paragraphs: Word.ParagraphCollection;
        /**
         *
         * Gets the parent body of the range. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentBody: Word.Body;
        /**
         *
         * Gets the content control that contains the range. Returns a null object if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        parentContentControl: Word.ContentControl;
        /**
         *
         * Gets the table that contains the range. Returns null if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTable: Word.Table;
        /**
         *
         * Gets the table cell that contains the range. Returns a null object if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTableCell: Word.TableCell;
        /**
         *
         * Gets the collection of table objects in the range. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        tables: Word.TableCollection;
        /**
         *
         * Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a newline character ('\n') to separate the address part from the optional location part.
         *
         * [Api set: WordApi 1.3]
         */
        hyperlink: string;
        /**
         *
         * Checks whether the range length is zero. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        isEmpty: boolean;
        /**
         *
         * Gets or sets the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
         *
         * [Api set: WordApi 1.1]
         */
        style: string;
        /**
         *
         * Gets or sets the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
         *
         * [Api set: WordApi 1.3]
         */
        styleBuiltIn: string;
        /**
         *
         * Gets the text of the range. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        text: string;
        /**
         *
         * ID
         *
         * [Api set: WordApi]
         */
        _Id: number;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        /**
         *
         * Clears the contents of the range object. The user can perform the undo operation on the cleared content.
         *
         * [Api set: WordApi 1.1]
         */
        clear(): void;
        /**
         *
         * Compares this range's location with another range's location.
         *
         * @param range Required. The range to compare with this range.
         *
         * [Api set: WordApi 1.3]
         */
        compareLocationWith(range: Word.Range): OfficeExtension.ClientResult<string>;
        /**
         *
         * Deletes the range and its content from the document.
         *
         * [Api set: WordApi 1.1]
         */
        delete(): void;
        /**
         *
         * Returns a new range that extends from this range in either direction to cover another range. This range is not changed.
         *
         * @param range Required. Another range.
         *
         * [Api set: WordApi 1.3]
         */
        expandTo(range: Word.Range): Word.Range;
        /**
         *
         * Gets the names all bookmarks in or overlapping the range. A bookmark is hidden if its name starts with an underscore character.
         *
         * @param includeHidden Optional. Indicates whether to include hidden bookmarks. Default is false which indicates that the hidden bookmarks are excluded.
         * @param includeAdjacent Optional. Indicates whether to include bookmarks that are adjacent to the range. Default is false which indicates that the adjacent bookmarks are excluded.
         *
         * [Api set: WordApi 1.4]
         */
        getBookmarks(includeHidden?: boolean, includeAdjacent?: boolean): OfficeExtension.ClientResult<Array<string>>;
        /**
         *
         * Gets the HTML representation of the range object.
         *
         * [Api set: WordApi 1.1]
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Gets hyperlink child ranges within the range.
         *
         * [Api set: WordApi 1.3]
         */
        getHyperlinkRanges(): Word.RangeCollection;
        /**
         *
         * Gets the next text range by using punctuation marks and/or other ending marks.
         *
         * @param endingMarks Required. The punctuation marks and/or other ending marks as an array of strings.
         * @param trimSpacing Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the returned range. Default is false which indicates that spacing characters at the start and end of the range are included.
         *
         * [Api set: WordApi 1.3]
         */
        getNextTextRange(endingMarks: Array<string>, trimSpacing?: boolean): Word.Range;
        /**
         *
         * Gets the OOXML representation of the range object.
         *
         * [Api set: WordApi 1.1]
         */
        getOoxml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Clones the range, or gets the starting or ending point of the range as a new range.
         *
         * @param rangeLocation Optional. The range location can be 'Whole', 'Start', 'End', 'After' or 'Content'.
         *
         * [Api set: WordApi 1.3]
         */
        getRange(rangeLocation?: string): Word.Range;
        /**
         *
         * Gets the text child ranges in the range by using punctuation marks and/or other ending marks.
         *
         * @param endingMarks Required. The punctuation marks and/or other ending marks as an array of strings.
         * @param trimSpacing Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         *
         * [Api set: WordApi 1.3]
         */
        getTextRanges(endingMarks: Array<string>, trimSpacing?: boolean): Word.RangeCollection;
        /**
         *
         * Inserts a bookmark on the range. If a bookmark of the same name exists, it is replaced.
         *
         * @param name Required. The bookmark name, which is case-insensitive. If the name starts with an underscore character, the bookmark is an hidden one.
         *
         * [Api set: WordApi 1.4]
         */
        insertBookmark(name: string): void;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.
         *
         * @param breakType Required. The break type to add.
         * @param insertLocation Required. The value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         */
        insertBreak(breakType: string, insertLocation: string): void;
        /**
         *
         * Wraps the range object with a rich text content control.
         *
         * [Api set: WordApi 1.1]
         */
        insertContentControl(): Word.ContentControl;
        /**
         *
         * Inserts a document at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         *
         * @param base64File Required. The base64 encoded content of a .docx file.
         * @param insertLocation Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         */
        insertFileFromBase64(base64File: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         *
         * @param html Required. The HTML to be inserted.
         * @param insertLocation Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         */
        insertHtml(html: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a picture at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         *
         * @param base64EncodedImage Required. The base64 encoded image to be inserted.
         * @param insertLocation Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string): Word.InlinePicture;
        /**
         *
         * Inserts OOXML at the specified location.  The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         *
         * @param ooxml Required. The OOXML to be inserted.
         * @param insertLocation Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         */
        insertOoxml(ooxml: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * @param paragraphText Required. The paragraph text to be inserted.
         * @param insertLocation Required. The value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         */
        insertParagraph(paragraphText: string, insertLocation: string): Word.Paragraph;
        /**
         *
         * Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.
         *
         * @param rowCount Required. The number of rows in the table.
         * @param columnCount Required. The number of columns in the table.
         * @param insertLocation Required. The value can be 'Before' or 'After'.
         * @param values Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         *
         * [Api set: WordApi 1.3]
         */
        insertTable(rowCount: number, columnCount: number, insertLocation: string, values?: Array<Array<string>>): Word.Table;
        /**
         *
         * Inserts text at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         *
         * @param text Required. Text to be inserted.
         * @param insertLocation Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         */
        insertText(text: string, insertLocation: string): Word.Range;
        /**
         *
         * Returns a new range as the intersection of this range with another range. This range is not changed.
         *
         * @param range Required. Another range.
         *
         * [Api set: WordApi 1.3]
         */
        intersectWith(range: Word.Range): Word.Range;
        /**
         *
         * Performs a search with the specified searchOptions on the scope of the range object. The search results are a collection of range objects.
         *
         * @param searchText Required. The search text.
         * @param searchOptions Optional. Options for the search.
         *
         * [Api set: WordApi 1.1]
         */
        search(searchText: string, searchOptions?: Word.SearchOptions | {
            ignorePunct?: boolean;
            ignoreSpace?: boolean;
            matchCase?: boolean;
            matchPrefix?: boolean;
            matchSuffix?: boolean;
            matchWholeWord?: boolean;
            matchWildcards?: boolean;
        }): Word.RangeCollection;
        /**
         *
         * Selects and navigates the Word UI to the range.
         *
         * @param selectionMode Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.
         *
         * [Api set: WordApi 1.1]
         */
        select(selectionMode?: string): void;
        /**
         *
         * Splits the range into child ranges by using delimiters.
         *
         * @param delimiters Required. The delimiters as an array of strings.
         * @param multiParagraphs Optional. Indicates whether a returned child range can cover multiple paragraphs. Default is false which indicates that the paragraph boundaries are also used as delimiters.
         * @param trimDelimiters Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.
         * @param trimSpacing Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         *
         * [Api set: WordApi 1.3]
         */
        split(delimiters: Array<string>, multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean): Word.RangeCollection;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.Range;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.Range;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.Range;
        toJSON(): {
            "font": Font;
            "hyperlink": string;
            "isEmpty": boolean;
            "style": string;
            "styleBuiltIn": string;
            "text": string;
        };
    }
    /**
     *
     * Contains a collection of [range](range.md) objects.
     *
     * [Api set: WordApi 1.3]
     */
    class RangeCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Word.Range>;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        /**
         *
         * Gets the first range in this collection.
         *
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.Range;
        /**
         *
         * Gets a range object by its index in the collection.
         *
         * @param index A number that identifies the index location of a range object.
         *
         * [Api set: WordApi 1.3]
         */
        _GetItem(index: number): Word.Range;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.RangeCollection;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.RangeCollection;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.RangeCollection;
        toJSON(): {};
    }
    /**
     *
     * Specifies the options to be included in a search operation.
     *
     * [Api set: WordApi 1.1]
     */
    class SearchOptions extends OfficeExtension.ClientObject {
        private m_ignorePunct;
        private m_ignoreSpace;
        private m_matchCase;
        private m_matchPrefix;
        private m_matchSuffix;
        private m_matchWholeWord;
        private m_matchWildcards;
        matchWildCards: boolean;
        /**
         *
         * Gets or sets a value that indicates whether to ignore all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.
         *
         * [Api set: WordApi 1.1]
         */
        ignorePunct: boolean;
        /**
         *
         * Gets or sets a value that indicates whether to ignore all whitespace between words. Corresponds to the Ignore whitespace characters check box in the Find and Replace dialog box.
         *
         * [Api set: WordApi 1.1]
         */
        ignoreSpace: boolean;
        /**
         *
         * Gets or sets a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box (Edit menu).
         *
         * [Api set: WordApi 1.1]
         */
        matchCase: boolean;
        /**
         *
         * Gets or sets a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.
         *
         * [Api set: WordApi 1.1]
         */
        matchPrefix: boolean;
        /**
         *
         * Gets or sets a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.
         *
         * [Api set: WordApi 1.1]
         */
        matchSuffix: boolean;
        /**
         *
         * Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.
         *
         * [Api set: WordApi 1.1]
         */
        matchWholeWord: boolean;
        /**
         *
         * Gets or sets a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box.
         *
         * [Api set: WordApi 1.1]
         */
        matchWildcards: boolean;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.SearchOptions;
        /**
         * Create a new instance of Word.SearchOptions object
         */
        static newObject(context: OfficeExtension.ClientRequestContext): Word.SearchOptions;
        toJSON(): {
            "ignorePunct": boolean;
            "ignoreSpace": boolean;
            "matchCase": boolean;
            "matchPrefix": boolean;
            "matchSuffix": boolean;
            "matchWholeWord": boolean;
            "matchWildcards": boolean;
        };
    }
    /**
     *
     * Represents a section in a Word document.
     *
     * [Api set: WordApi 1.1]
     */
    class Section extends OfficeExtension.ClientObject {
        private m_body;
        private m__Id;
        private m__ReferenceId;
        /**
         *
         * Gets the body object of the section. This does not include the header/footer and other section metadata. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        body: Word.Body;
        /**
         *
         * ID
         *
         * [Api set: WordApi]
         */
        _Id: number;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        /**
         *
         * Gets one of the section's footers.
         *
         * @param type Required. The type of footer to return. This value can be: 'primary', 'firstPage' or 'evenPages'.
         *
         * [Api set: WordApi 1.1]
         */
        getFooter(type: string): Word.Body;
        /**
         *
         * Gets one of the section's headers.
         *
         * @param type Required. The type of header to return. This value can be: 'primary', 'firstPage' or 'evenPages'.
         *
         * [Api set: WordApi 1.1]
         */
        getHeader(type: string): Word.Body;
        /**
         *
         * Gets the next section.
         *
         * [Api set: WordApi 1.3]
         */
        getNext(): Word.Section;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.Section;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.Section;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.Section;
        toJSON(): {
            "body": Body;
        };
    }
    /**
     *
     * Contains the collection of the document's [section](section.md) objects.
     *
     * [Api set: WordApi 1.1]
     */
    class SectionCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Word.Section>;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        /**
         *
         * Gets the first section in this collection.
         *
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.Section;
        /**
         *
         * Gets a section object by its index in the collection.
         *
         * @param index A number that identifies the index location of a section object.
         *
         * [Api set: WordApi 1.1]
         */
        _GetItem(index: number): Word.Section;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.SectionCollection;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.SectionCollection;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.SectionCollection;
        toJSON(): {};
    }
    /**
     *
     * Represents a setting of the add-in.
     *
     * [Api set: WordApi 1.4]
     */
    class Setting extends OfficeExtension.ClientObject {
        private m_key;
        private m_value;
        private m__ReferenceId;
        /**
         *
         * Gets the key of the setting. Read only.
         *
         * [Api set: WordApi 1.4]
         */
        key: string;
        /**
         *
         * Gets or sets the value of the setting.
         *
         * [Api set: WordApi 1.4]
         */
        value: any;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        /**
         *
         * Deletes the setting.
         *
         * [Api set: WordApi 1.4]
         */
        delete(): void;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.Setting;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.Setting;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.Setting;
        toJSON(): {
            "key": string;
            "value": any;
        };
    }
    /**
     *
     * Contains the collection of [setting](setting.md) objects.
     *
     * [Api set: WordApi 1.4]
     */
    class SettingCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Word.Setting>;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        /**
         *
         * Deletes all settings in this add-in.
         *
         * [Api set: WordApi 1.4]
         */
        deleteAll(): void;
        /**
         *
         * Gets the count of settings.
         *
         * [Api set: WordApi 1.4]
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         *
         * Gets a setting object by its key, which is case-sensitive.
         *
         * @param key The key that identifies the setting object.
         *
         * [Api set: WordApi 1.4]
         */
        getItem(key: string): Word.Setting;
        /**
         *
         * Creates or sets a setting.
         *
         * @param key Required. The setting's key, which is case-sensitive.
         * @param value Required. The setting's value.
         *
         * [Api set: WordApi 1.4]
         */
        set(key: string, value: any): Word.Setting;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.SettingCollection;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.SettingCollection;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.SettingCollection;
        toJSON(): {};
    }
    /**
     *
     * Represents a table in a Word document.
     *
     * [Api set: WordApi 1.3]
     */
    class Table extends OfficeExtension.ClientObject {
        private m_font;
        private m_headerRowCount;
        private m_height;
        private m_horizontalAlignment;
        private m_isUniform;
        private m_nestingLevel;
        private m_paragraphAfter;
        private m_paragraphBefore;
        private m_parentBody;
        private m_parentContentControl;
        private m_parentTable;
        private m_parentTableCell;
        private m_rowCount;
        private m_rows;
        private m_shadingColor;
        private m_style;
        private m_styleBandedColumns;
        private m_styleBandedRows;
        private m_styleBuiltIn;
        private m_styleFirstColumn;
        private m_styleLastColumn;
        private m_styleTotalRow;
        private m_tables;
        private m_values;
        private m_verticalAlignment;
        private m_width;
        private m__Id;
        private m__ReferenceId;
        /**
         *
         * Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        font: Word.Font;
        /**
         *
         * Gets the paragraph after the table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        paragraphAfter: Word.Paragraph;
        /**
         *
         * Gets the paragraph before the table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        paragraphBefore: Word.Paragraph;
        /**
         *
         * Gets the parent body of the table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentBody: Word.Body;
        /**
         *
         * Gets the content control that contains the table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentContentControl: Word.ContentControl;
        /**
         *
         * Gets the table that contains this table. Returns a null object if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTable: Word.Table;
        /**
         *
         * Gets the table cell that contains this table. Returns a null object if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTableCell: Word.TableCell;
        /**
         *
         * Gets all of the table rows. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        rows: Word.TableRowCollection;
        /**
         *
         * Gets the child tables nested one level deeper. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        tables: Word.TableCollection;
        /**
         *
         * Gets and sets the number of header rows.
         *
         * [Api set: WordApi 1.3]
         */
        headerRowCount: number;
        /**
         *
         * Gets the height of the table in points. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        height: number;
        /**
         *
         * Gets and sets the horizontal alignment of every cell in the table. The value can be 'left', 'centered', 'right', or 'justified'.
         *
         * [Api set: WordApi 1.3]
         */
        horizontalAlignment: string;
        /**
         *
         * Indicates whether all of the table rows are uniform. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        isUniform: boolean;
        /**
         *
         * Gets the nesting level of the table. Top-level tables have level 1. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        nestingLevel: number;
        /**
         *
         * Gets the number of rows in the table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        rowCount: number;
        /**
         *
         * Gets and sets the shading color.
         *
         * [Api set: WordApi 1.3]
         */
        shadingColor: string;
        /**
         *
         * Gets or sets the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
         *
         * [Api set: WordApi 1.3]
         */
        style: string;
        /**
         *
         * Gets and sets whether the table has banded columns.
         *
         * [Api set: WordApi 1.3]
         */
        styleBandedColumns: boolean;
        /**
         *
         * Gets and sets whether the table has banded rows.
         *
         * [Api set: WordApi 1.3]
         */
        styleBandedRows: boolean;
        /**
         *
         * Gets or sets the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
         *
         * [Api set: WordApi 1.3]
         */
        styleBuiltIn: string;
        /**
         *
         * Gets and sets whether the table has a first column with a special style.
         *
         * [Api set: WordApi 1.3]
         */
        styleFirstColumn: boolean;
        /**
         *
         * Gets and sets whether the table has a last column with a special style.
         *
         * [Api set: WordApi 1.3]
         */
        styleLastColumn: boolean;
        /**
         *
         * Gets and sets whether the table has a total (last) row with a special style.
         *
         * [Api set: WordApi 1.3]
         */
        styleTotalRow: boolean;
        /**
         *
         * Gets and sets the text values in the table, as a 2D Javascript array.
         *
         * [Api set: WordApi 1.3]
         */
        values: Array<Array<string>>;
        /**
         *
         * Gets and sets the vertical alignment of every cell in the table. The value can be 'top', 'center' or 'bottom'.
         *
         * [Api set: WordApi 1.3]
         */
        verticalAlignment: string;
        /**
         *
         * Gets and sets the width of the table in points.
         *
         * [Api set: WordApi 1.3]
         */
        width: number;
        /**
         *
         * ID
         *
         * [Api set: WordApi]
         */
        _Id: number;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        /**
         *
         * Adds columns to the start or end of the table, using the first or last existing column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.
         *
         * @param insertLocation Required. It can be 'Start' or 'End', corresponding to the appropriate side of the table.
         * @param columnCount Required. Number of columns to add.
         * @param values Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         *
         * [Api set: WordApi 1.3]
         */
        addColumns(insertLocation: string, columnCount: number, values?: Array<Array<string>>): void;
        /**
         *
         * Adds rows to the start or end of the table, using the first or last existing row as a template. The string values, if specified, are set in the newly inserted rows.
         *
         * @param insertLocation Required. It can be 'Start' or 'End'.
         * @param rowCount Required. Number of rows to add.
         * @param values Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         *
         * [Api set: WordApi 1.3]
         */
        addRows(insertLocation: string, rowCount: number, values?: Array<Array<string>>): Word.TableRowCollection;
        /**
         *
         * Autofits the table columns to the width of their contents.
         *
         * [Api set: WordApi 1.3]
         */
        autoFitContents(): void;
        /**
         *
         * Autofits the table columns to the width of the window.
         *
         * [Api set: WordApi 1.3]
         */
        autoFitWindow(): void;
        /**
         *
         * Clears the contents of the table.
         *
         * [Api set: WordApi 1.3]
         */
        clear(): void;
        /**
         *
         * Deletes the entire table.
         *
         * [Api set: WordApi 1.3]
         */
        delete(): void;
        /**
         *
         * Deletes specific columns. This is applicable to uniform tables.
         *
         * @param columnIndex Required. The first column to delete.
         * @param columnCount Optional. The number of columns to delete. Default 1.
         *
         * [Api set: WordApi 1.3]
         */
        deleteColumns(columnIndex: number, columnCount?: number): void;
        /**
         *
         * Deletes specific rows.
         *
         * @param rowIndex Required. The first row to delete.
         * @param rowCount Optional. The number of rows to delete. Default 1.
         *
         * [Api set: WordApi 1.3]
         */
        deleteRows(rowIndex: number, rowCount?: number): void;
        /**
         *
         * Distributes the column widths evenly.
         *
         * [Api set: WordApi 1.3]
         */
        distributeColumns(): void;
        /**
         *
         * Distributes the row heights evenly.
         *
         * [Api set: WordApi 1.3]
         */
        distributeRows(): void;
        /**
         *
         * Gets the border style for the specified border.
         *
         * @param borderLocation Required. The border location.
         *
         * [Api set: WordApi 1.3]
         */
        getBorder(borderLocation: string): Word.TableBorder;
        /**
         *
         * Gets the table cell at a specified row and column.
         *
         * @param rowIndex Required. The index of the row.
         * @param cellIndex Required. The index of the cell in the row.
         *
         * [Api set: WordApi 1.3]
         */
        getCell(rowIndex: number, cellIndex: number): Word.TableCell;
        /**
         *
         * Gets cell padding in points.
         *
         * @param cellPaddingLocation Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.
         *
         * [Api set: WordApi 1.3]
         */
        getCellPadding(cellPaddingLocation: string): OfficeExtension.ClientResult<number>;
        /**
         *
         * Gets the next table.
         *
         * [Api set: WordApi 1.3]
         */
        getNext(): Word.Table;
        /**
         *
         * Gets the range that contains this table, or the range at the start or end of the table.
         *
         * @param rangeLocation Optional. The range location can be 'Whole', 'Start', 'End' or 'After'.
         *
         * [Api set: WordApi 1.3]
         */
        getRange(rangeLocation?: string): Word.Range;
        /**
         *
         * Inserts a content control on the table.
         *
         * [Api set: WordApi 1.3]
         */
        insertContentControl(): Word.ContentControl;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * @param paragraphText Required. The paragraph text to be inserted.
         * @param insertLocation Required. The value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.3]
         */
        insertParagraph(paragraphText: string, insertLocation: string): Word.Paragraph;
        /**
         *
         * Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.
         *
         * @param rowCount Required. The number of rows in the table.
         * @param columnCount Required. The number of columns in the table.
         * @param insertLocation Required. The value can be 'Before' or 'After'.
         * @param values Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         *
         * [Api set: WordApi 1.3]
         */
        insertTable(rowCount: number, columnCount: number, insertLocation: string, values?: Array<Array<string>>): Word.Table;
        /**
         *
         * Merges the cells bounded inclusively by a first and last cell.
         *
         * @param topRow Required. The row of the first cell
         * @param firstCell Required. The index of the first cell in its row
         * @param bottomRow Required. The row of the last cell
         * @param lastCell Required. The index of the last cell in its row
         *
         * [Api set: WordApi 1.4]
         */
        mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number): Word.TableCell;
        /**
         *
         * Performs a search with the specified searchOptions on the scope of the table object. The search results are a collection of range objects.
         *
         * @param searchText Required. The search text.
         * @param searchOptions Optional. Options for the search.
         *
         * [Api set: WordApi 1.3]
         */
        search(searchText: string, searchOptions?: Word.SearchOptions | {
            ignorePunct?: boolean;
            ignoreSpace?: boolean;
            matchCase?: boolean;
            matchPrefix?: boolean;
            matchSuffix?: boolean;
            matchWholeWord?: boolean;
            matchWildcards?: boolean;
        }): Word.RangeCollection;
        /**
         *
         * Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.
         *
         * @param selectionMode Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.
         *
         * [Api set: WordApi 1.3]
         */
        select(selectionMode?: string): void;
        /**
         *
         * Sets cell padding in points.
         *
         * @param cellPaddingLocation Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.
         *
         * [Api set: WordApi 1.3]
         */
        setCellPadding(cellPaddingLocation: string, cellPadding: number): void;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.Table;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.Table;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.Table;
        toJSON(): {
            "font": Font;
            "headerRowCount": number;
            "height": number;
            "horizontalAlignment": string;
            "isUniform": boolean;
            "nestingLevel": number;
            "rowCount": number;
            "shadingColor": string;
            "style": string;
            "styleBandedColumns": boolean;
            "styleBandedRows": boolean;
            "styleBuiltIn": string;
            "styleFirstColumn": boolean;
            "styleLastColumn": boolean;
            "styleTotalRow": boolean;
            "values": string[][];
            "verticalAlignment": string;
            "width": number;
        };
    }
    /**
     *
     * Contains the collection of the document's Table objects.
     *
     * [Api set: WordApi 1.3]
     */
    class TableCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Word.Table>;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        /**
         *
         * Gets the first table in this collection.
         *
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.Table;
        /**
         *
         * Gets a table object by its index in the collection.
         *
         * @param index A number that identifies the index location of a table object.
         *
         * [Api set: WordApi 1.3]
         */
        _GetItem(index: number): Word.Table;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.TableCollection;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.TableCollection;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.TableCollection;
        toJSON(): {};
    }
    /**
     *
     * Represents a row in a Word document.
     *
     * [Api set: WordApi 1.3]
     */
    class TableRow extends OfficeExtension.ClientObject {
        private m_cellCount;
        private m_cells;
        private m_font;
        private m_horizontalAlignment;
        private m_isHeader;
        private m_parentTable;
        private m_preferredHeight;
        private m_rowIndex;
        private m_shadingColor;
        private m_values;
        private m_verticalAlignment;
        private m__Id;
        private m__ReferenceId;
        /**
         *
         * Gets cells. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        cells: Word.TableCellCollection;
        /**
         *
         * Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        font: Word.Font;
        /**
         *
         * Gets parent table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTable: Word.Table;
        /**
         *
         * Gets the number of cells in the row. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        cellCount: number;
        /**
         *
         * Gets and sets the horizontal alignment of every cell in the row. The value can be 'left', 'centered', 'right', or 'justified'.
         *
         * [Api set: WordApi 1.3]
         */
        horizontalAlignment: string;
        /**
         *
         * Checks whether the row is a header row. Read-only. To set the number of header rows, use HeaderRowCount on the Table object.
         *
         * [Api set: WordApi 1.3]
         */
        isHeader: boolean;
        /**
         *
         * Gets and sets the preferred height of the row in points.
         *
         * [Api set: WordApi 1.3]
         */
        preferredHeight: number;
        /**
         *
         * Gets the index of the row in its parent table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        rowIndex: number;
        /**
         *
         * Gets and sets the shading color.
         *
         * [Api set: WordApi 1.3]
         */
        shadingColor: string;
        /**
         *
         * Gets and sets the text values in the row, as a 1D Javascript array.
         *
         * [Api set: WordApi 1.3]
         */
        values: Array<string>;
        /**
         *
         * Gets and sets the vertical alignment of the cells in the row. The value can be 'top', 'center' or 'bottom'.
         *
         * [Api set: WordApi 1.3]
         */
        verticalAlignment: string;
        /**
         *
         * ID
         *
         * [Api set: WordApi]
         */
        _Id: number;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        /**
         *
         * Clears the contents of the row.
         *
         * [Api set: WordApi 1.3]
         */
        clear(): void;
        /**
         *
         * Deletes the entire row.
         *
         * [Api set: WordApi 1.3]
         */
        delete(): void;
        /**
         *
         * Gets the border style of the cells in the row.
         *
         * @param borderLocation Required. The border location.
         *
         * [Api set: WordApi 1.3]
         */
        getBorder(borderLocation: string): Word.TableBorder;
        /**
         *
         * Gets cell padding in points.
         *
         * @param cellPaddingLocation Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.
         *
         * [Api set: WordApi 1.3]
         */
        getCellPadding(cellPaddingLocation: string): OfficeExtension.ClientResult<number>;
        /**
         *
         * Gets the next row.
         *
         * [Api set: WordApi 1.3]
         */
        getNext(): Word.TableRow;
        /**
         *
         * Inserts rows using this row as a template. If values are specified, inserts the values into the new rows.
         *
         * @param insertLocation Required. Where the new rows should be inserted, relative to the current row. It can be 'Before' or 'After'.
         * @param rowCount Required. Number of rows to add
         * @param values Optional. Strings to insert in the new rows, specified as a 2D array. The number of cells in each row must not exceed the number of cells in the existing row.
         *
         * [Api set: WordApi 1.3]
         */
        insertRows(insertLocation: string, rowCount: number, values?: Array<Array<string>>): Word.TableRowCollection;
        /**
         *
         * Merges the row into one cell.
         *
         * [Api set: WordApi 1.4]
         */
        merge(): Word.TableCell;
        /**
         *
         * Performs a search with the specified searchOptions on the scope of the row. The search results are a collection of range objects.
         *
         * @param searchText Required. The search text.
         * @param searchOptions Optional. Options for the search.
         *
         * [Api set: WordApi 1.3]
         */
        search(searchText: string, searchOptions?: Word.SearchOptions | {
            ignorePunct?: boolean;
            ignoreSpace?: boolean;
            matchCase?: boolean;
            matchPrefix?: boolean;
            matchSuffix?: boolean;
            matchWholeWord?: boolean;
            matchWildcards?: boolean;
        }): Word.RangeCollection;
        /**
         *
         * Selects the row and navigates the Word UI to it.
         *
         * @param selectionMode Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.
         *
         * [Api set: WordApi 1.3]
         */
        select(selectionMode?: string): void;
        /**
         *
         * Sets cell padding in points.
         *
         * @param cellPaddingLocation Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.
         *
         * [Api set: WordApi 1.3]
         */
        setCellPadding(cellPaddingLocation: string, cellPadding: number): void;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.TableRow;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.TableRow;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.TableRow;
        toJSON(): {
            "cellCount": number;
            "font": Font;
            "horizontalAlignment": string;
            "isHeader": boolean;
            "preferredHeight": number;
            "rowIndex": number;
            "shadingColor": string;
            "values": string[];
            "verticalAlignment": string;
        };
    }
    /**
     *
     * Contains the collection of the document's TableRow objects.
     *
     * [Api set: WordApi 1.3]
     */
    class TableRowCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Word.TableRow>;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        /**
         *
         * Gets the first row in this collection.
         *
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.TableRow;
        /**
         *
         * Gets a table row object by its index in the collection.
         *
         * @param index A number that identifies the index location of a table row object.
         *
         * [Api set: WordApi 1.3]
         */
        _GetItem(index: number): Word.TableRow;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.TableRowCollection;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.TableRowCollection;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.TableRowCollection;
        toJSON(): {};
    }
    /**
     *
     * Represents a table cell in a Word document.
     *
     * [Api set: WordApi 1.3]
     */
    class TableCell extends OfficeExtension.ClientObject {
        private m_body;
        private m_cellIndex;
        private m_columnWidth;
        private m_horizontalAlignment;
        private m_parentRow;
        private m_parentTable;
        private m_rowIndex;
        private m_shadingColor;
        private m_value;
        private m_verticalAlignment;
        private m_width;
        private m__Id;
        private m__ReferenceId;
        /**
         *
         * Gets the body object of the cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        body: Word.Body;
        /**
         *
         * Gets the parent row of the cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentRow: Word.TableRow;
        /**
         *
         * Gets the parent table of the cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTable: Word.Table;
        /**
         *
         * Gets the index of the cell in its row. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        cellIndex: number;
        /**
         *
         * Gets and sets the width of the cell's column in points. This is applicable to uniform tables.
         *
         * [Api set: WordApi 1.3]
         */
        columnWidth: number;
        /**
         *
         * Gets and sets the horizontal alignment of the cell. The value can be 'left', 'centered', 'right', or 'justified'.
         *
         * [Api set: WordApi 1.3]
         */
        horizontalAlignment: string;
        /**
         *
         * Gets the index of the cell's row in the table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        rowIndex: number;
        /**
         *
         * Gets or sets the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.
         *
         * [Api set: WordApi 1.3]
         */
        shadingColor: string;
        /**
         *
         * Gets and sets the text of the cell.
         *
         * [Api set: WordApi 1.3]
         */
        value: string;
        /**
         *
         * Gets and sets the vertical alignment of the cell. The value can be 'top', 'center' or 'bottom'.
         *
         * [Api set: WordApi 1.3]
         */
        verticalAlignment: string;
        /**
         *
         * Gets the width of the cell in points. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        width: number;
        /**
         *
         * ID
         *
         * [Api set: WordApi]
         */
        _Id: number;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        /**
         *
         * Deletes the column containing this cell. This is applicable to uniform tables.
         *
         * [Api set: WordApi 1.3]
         */
        deleteColumn(): void;
        /**
         *
         * Deletes the row containing this cell.
         *
         * [Api set: WordApi 1.3]
         */
        deleteRow(): void;
        /**
         *
         * Gets the border style for the specified border.
         *
         * @param borderLocation Required. The border location.
         *
         * [Api set: WordApi 1.3]
         */
        getBorder(borderLocation: string): Word.TableBorder;
        /**
         *
         * Gets cell padding in points.
         *
         * @param cellPaddingLocation Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.
         *
         * [Api set: WordApi 1.3]
         */
        getCellPadding(cellPaddingLocation: string): OfficeExtension.ClientResult<number>;
        /**
         *
         * Gets the next cell.
         *
         * [Api set: WordApi 1.3]
         */
        getNext(): Word.TableCell;
        /**
         *
         * Adds columns to the left or right of the cell, using the cell's column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.
         *
         * @param insertLocation Required. It can be 'Before' or 'After'.
         * @param columnCount Required. Number of columns to add
         * @param values Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         *
         * [Api set: WordApi 1.3]
         */
        insertColumns(insertLocation: string, columnCount: number, values?: Array<Array<string>>): void;
        /**
         *
         * Inserts rows above or below the cell, using the cell's row as a template. The string values, if specified, are set in the newly inserted rows.
         *
         * @param insertLocation Required. It can be 'Before' or 'After'.
         * @param rowCount Required. Number of rows to add.
         * @param values Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         *
         * [Api set: WordApi 1.3]
         */
        insertRows(insertLocation: string, rowCount: number, values?: Array<Array<string>>): Word.TableRowCollection;
        /**
         *
         * Sets cell padding in points.
         *
         * @param cellPaddingLocation Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.
         *
         * [Api set: WordApi 1.3]
         */
        setCellPadding(cellPaddingLocation: string, cellPadding: number): void;
        /**
         *
         * Adds columns to the left or right of the cell, using the existing column as a template. The string values, if specified, are set in the newly inserted rows.
         *
         * @param rowCount Required. The number of rows to split into. Must be a divisor of the number of underlying rows.
         * @param columnCount Required. The number of columns to split into.
         *
         * [Api set: WordApi 1.4]
         */
        split(rowCount: number, columnCount: number): void;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.TableCell;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.TableCell;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.TableCell;
        toJSON(): {
            "body": Body;
            "cellIndex": number;
            "columnWidth": number;
            "horizontalAlignment": string;
            "rowIndex": number;
            "shadingColor": string;
            "value": string;
            "verticalAlignment": string;
            "width": number;
        };
    }
    /**
     *
     * Contains the collection of the document's TableCell objects.
     *
     * [Api set: WordApi 1.3]
     */
    class TableCellCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Word.TableCell>;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        /**
         *
         * Gets the first table cell in this collection.
         *
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.TableCell;
        /**
         *
         * Gets a table cell object by its index in the collection.
         *
         * @param index A number that identifies the index location of a table cell object.
         *
         * [Api set: WordApi 1.3]
         */
        _GetItem(index: number): Word.TableCell;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.TableCellCollection;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.TableCellCollection;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.TableCellCollection;
        toJSON(): {};
    }
    /**
     *
     * Specifies the border style
     *
     * [Api set: WordApi 1.3]
     */
    class TableBorder extends OfficeExtension.ClientObject {
        private m_color;
        private m_type;
        private m_width;
        private m__ReferenceId;
        /**
         *
         * Gets or sets the table border color, as a hex value or name.
         *
         * [Api set: WordApi 1.3]
         */
        color: string;
        /**
         *
         * Gets or sets the type of the table border.
         *
         * [Api set: WordApi 1.3]
         */
        type: string;
        /**
         *
         * Gets or sets the width, in points, of the table border. Not applicable to table border types that have fixed widths.
         *
         * [Api set: WordApi 1.3]
         */
        width: number;
        /**
         *
         * ReferenceId
         *
         * [Api set: WordApi]
         */
        _ReferenceId: string;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.TableBorder;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.TableBorder;
        /**
         * Release the memory associated with this object, if has previous been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.TableBorder;
        toJSON(): {
            "color": string;
            "type": string;
            "width": number;
        };
    }
    /**
     *
     * Specifies supported content control types and subtypes.
     *
     * [Api set: WordApi]
     */
    module ContentControlType {
        var unknown: string;
        var richTextInline: string;
        var richTextParagraphs: string;
        var richTextTableCell: string;
        var richTextTableRow: string;
        var richTextTable: string;
        var plainTextInline: string;
        var plainTextParagraph: string;
        var picture: string;
        var buildingBlockGallery: string;
        var checkBox: string;
        var comboBox: string;
        var dropDownList: string;
        var datePicker: string;
        var repeatingSection: string;
        var richText: string;
        var plainText: string;
    }
    /**
     *
     * ContentControl appearance
     *
     * [Api set: WordApi]
     */
    module ContentControlAppearance {
        var boundingBox: string;
        var tags: string;
        var hidden: string;
    }
    /**
     *
     * Underline types
     *
     * [Api set: WordApi]
     */
    module UnderlineType {
        var mixed: string;
        var none: string;
        /**
         *
         * @deprecated Hidden is no longer supported.
         */
        var hidden: string;
        /**
         *
         * @deprecated DotLine is no longer supported.
         */
        var dotLine: string;
        var single: string;
        var word: string;
        var double: string;
        var thick: string;
        var dotted: string;
        var dottedHeavy: string;
        var dashLine: string;
        var dashLineHeavy: string;
        var dashLineLong: string;
        var dashLineLongHeavy: string;
        var dotDashLine: string;
        var dotDashLineHeavy: string;
        var twoDotDashLine: string;
        var twoDotDashLineHeavy: string;
        var wave: string;
        var waveHeavy: string;
        var waveDouble: string;
    }
    /**
     *
     * Page break, line break, and four section breaks
     *
     * [Api set: WordApi]
     */
    module BreakType {
        /**
         *
         * Page break.
         *
         */
        var page: string;
        /**
         *
         * @deprecated Use sectionNext instead.
         */
        var next: string;
        /**
         *
         * Section break, with the new section starting on the next page.
         *
         */
        var sectionNext: string;
        /**
         *
         * Section break, with the new section starting on the same page.
         *
         */
        var sectionContinuous: string;
        /**
         *
         * Section break, with the new section starting on the next even-numbered page.
         *
         */
        var sectionEven: string;
        /**
         *
         * Section break, with the new section starting on the next odd-numbered page.
         *
         */
        var sectionOdd: string;
        /**
         *
         * Line break.
         *
         */
        var line: string;
    }
    /**
     *
     * The insertion location types
     *
     * [Api set: WordApi]
     */
    module InsertLocation {
        var before: string;
        var after: string;
        var start: string;
        var end: string;
        var replace: string;
    }
    /**
     * [Api set: WordApi]
     */
    module Alignment {
        var mixed: string;
        var unknown: string;
        var left: string;
        var centered: string;
        var right: string;
        var justified: string;
    }
    /**
     * [Api set: WordApi]
     */
    module HeaderFooterType {
        var primary: string;
        var firstPage: string;
        var evenPages: string;
    }
    /**
     * [Api set: WordApi]
     */
    module BodyType {
        var unknown: string;
        var mainDoc: string;
        var section: string;
        var header: string;
        var footer: string;
        var tableCell: string;
    }
    /**
     * [Api set: WordApi]
     */
    module SelectionMode {
        var select: string;
        var start: string;
        var end: string;
    }
    /**
     * [Api set: WordApi]
     */
    module ImageFormat {
        var unsupported: string;
        var undefined: string;
        var bmp: string;
        var jpeg: string;
        var gif: string;
        var tiff: string;
        var png: string;
        var icon: string;
        var exif: string;
        var wmf: string;
        var emf: string;
        var pict: string;
        var pdf: string;
        var svg: string;
    }
    /**
     * [Api set: WordApi]
     */
    module RangeLocation {
        var whole: string;
        var start: string;
        var end: string;
        var before: string;
        var after: string;
        var content: string;
    }
    /**
     * [Api set: WordApi]
     */
    module LocationRelation {
        var unrelated: string;
        var equal: string;
        var containsStart: string;
        var containsEnd: string;
        var contains: string;
        var insideStart: string;
        var insideEnd: string;
        var inside: string;
        var adjacentBefore: string;
        var overlapsBefore: string;
        var before: string;
        var adjacentAfter: string;
        var overlapsAfter: string;
        var after: string;
    }
    /**
     * [Api set: WordApi]
     */
    module BorderLocation {
        var top: string;
        var left: string;
        var bottom: string;
        var right: string;
        var insideHorizontal: string;
        var insideVertical: string;
        var inside: string;
        var outside: string;
        var all: string;
    }
    /**
     * [Api set: WordApi]
     */
    module CellPaddingLocation {
        var top: string;
        var left: string;
        var bottom: string;
        var right: string;
    }
    /**
     * [Api set: WordApi]
     */
    module BorderType {
        var mixed: string;
        var none: string;
        var single: string;
        var double: string;
        var dotted: string;
        var dashed: string;
        var dotDashed: string;
        var dot2Dashed: string;
        var triple: string;
        var thinThickSmall: string;
        var thickThinSmall: string;
        var thinThickThinSmall: string;
        var thinThickMed: string;
        var thickThinMed: string;
        var thinThickThinMed: string;
        var thinThickLarge: string;
        var thickThinLarge: string;
        var thinThickThinLarge: string;
        var wave: string;
        var doubleWave: string;
        var dashedSmall: string;
        var dashDotStroked: string;
        var threeDEmboss: string;
        var threeDEngrave: string;
    }
    /**
     * [Api set: WordApi]
     */
    module VerticalAlignment {
        var mixed: string;
        var top: string;
        var center: string;
        var bottom: string;
    }
    /**
     * [Api set: WordApi]
     */
    module ListLevelType {
        var bullet: string;
        var number: string;
        var picture: string;
    }
    /**
     * [Api set: WordApi]
     */
    module ListBullet {
        var custom: string;
        var solid: string;
        var hollow: string;
        var square: string;
        var diamonds: string;
        var arrow: string;
        var checkmark: string;
    }
    /**
     * [Api set: WordApi]
     */
    module ListNumbering {
        var none: string;
        var arabic: string;
        var upperRoman: string;
        var lowerRoman: string;
        var upperLetter: string;
        var lowerLetter: string;
    }
    /**
     * [Api set: WordApi]
     */
    module Style {
        /**
         *
         * Mixed styles or other style not in this list.
         *
         */
        var other: string;
        /**
         *
         * Reset character and paragraph style to default.
         *
         */
        var normal: string;
        var heading1: string;
        var heading2: string;
        var heading3: string;
        var heading4: string;
        var heading5: string;
        var heading6: string;
        var heading7: string;
        var heading8: string;
        var heading9: string;
        /**
         *
         * Table-of-content level 1.
         *
         */
        var toc1: string;
        /**
         *
         * Table-of-content level 2.
         *
         */
        var toc2: string;
        /**
         *
         * Table-of-content level 3.
         *
         */
        var toc3: string;
        /**
         *
         * Table-of-content level 4.
         *
         */
        var toc4: string;
        /**
         *
         * Table-of-content level 5.
         *
         */
        var toc5: string;
        /**
         *
         * Table-of-content level 6.
         *
         */
        var toc6: string;
        /**
         *
         * Table-of-content level 7.
         *
         */
        var toc7: string;
        /**
         *
         * Table-of-content level 8.
         *
         */
        var toc8: string;
        /**
         *
         * Table-of-content level 9.
         *
         */
        var toc9: string;
        var footnoteText: string;
        var header: string;
        var footer: string;
        var caption: string;
        var footnoteReference: string;
        var endnoteReference: string;
        var endnoteText: string;
        var title: string;
        var subtitle: string;
        var hyperlink: string;
        var strong: string;
        var emphasis: string;
        var noSpacing: string;
        var listParagraph: string;
        var quote: string;
        var intenseQuote: string;
        var subtleEmphasis: string;
        var intenseEmphasis: string;
        var subtleReference: string;
        var intenseReference: string;
        var bookTitle: string;
        var bibliography: string;
        /**
         *
         * Table-of-content heading.
         *
         */
        var tocHeading: string;
        var tableGrid: string;
        var plainTable1: string;
        var plainTable2: string;
        var plainTable3: string;
        var plainTable4: string;
        var plainTable5: string;
        var tableGridLight: string;
        var gridTable1Light: string;
        var gridTable1Light_Accent1: string;
        var gridTable1Light_Accent2: string;
        var gridTable1Light_Accent3: string;
        var gridTable1Light_Accent4: string;
        var gridTable1Light_Accent5: string;
        var gridTable1Light_Accent6: string;
        var gridTable2: string;
        var gridTable2_Accent1: string;
        var gridTable2_Accent2: string;
        var gridTable2_Accent3: string;
        var gridTable2_Accent4: string;
        var gridTable2_Accent5: string;
        var gridTable2_Accent6: string;
        var gridTable3: string;
        var gridTable3_Accent1: string;
        var gridTable3_Accent2: string;
        var gridTable3_Accent3: string;
        var gridTable3_Accent4: string;
        var gridTable3_Accent5: string;
        var gridTable3_Accent6: string;
        var gridTable4: string;
        var gridTable4_Accent1: string;
        var gridTable4_Accent2: string;
        var gridTable4_Accent3: string;
        var gridTable4_Accent4: string;
        var gridTable4_Accent5: string;
        var gridTable4_Accent6: string;
        var gridTable5Dark: string;
        var gridTable5Dark_Accent1: string;
        var gridTable5Dark_Accent2: string;
        var gridTable5Dark_Accent3: string;
        var gridTable5Dark_Accent4: string;
        var gridTable5Dark_Accent5: string;
        var gridTable5Dark_Accent6: string;
        var gridTable6Colorful: string;
        var gridTable6Colorful_Accent1: string;
        var gridTable6Colorful_Accent2: string;
        var gridTable6Colorful_Accent3: string;
        var gridTable6Colorful_Accent4: string;
        var gridTable6Colorful_Accent5: string;
        var gridTable6Colorful_Accent6: string;
        var gridTable7Colorful: string;
        var gridTable7Colorful_Accent1: string;
        var gridTable7Colorful_Accent2: string;
        var gridTable7Colorful_Accent3: string;
        var gridTable7Colorful_Accent4: string;
        var gridTable7Colorful_Accent5: string;
        var gridTable7Colorful_Accent6: string;
        var listTable1Light: string;
        var listTable1Light_Accent1: string;
        var listTable1Light_Accent2: string;
        var listTable1Light_Accent3: string;
        var listTable1Light_Accent4: string;
        var listTable1Light_Accent5: string;
        var listTable1Light_Accent6: string;
        var listTable2: string;
        var listTable2_Accent1: string;
        var listTable2_Accent2: string;
        var listTable2_Accent3: string;
        var listTable2_Accent4: string;
        var listTable2_Accent5: string;
        var listTable2_Accent6: string;
        var listTable3: string;
        var listTable3_Accent1: string;
        var listTable3_Accent2: string;
        var listTable3_Accent3: string;
        var listTable3_Accent4: string;
        var listTable3_Accent5: string;
        var listTable3_Accent6: string;
        var listTable4: string;
        var listTable4_Accent1: string;
        var listTable4_Accent2: string;
        var listTable4_Accent3: string;
        var listTable4_Accent4: string;
        var listTable4_Accent5: string;
        var listTable4_Accent6: string;
        var listTable5Dark: string;
        var listTable5Dark_Accent1: string;
        var listTable5Dark_Accent2: string;
        var listTable5Dark_Accent3: string;
        var listTable5Dark_Accent4: string;
        var listTable5Dark_Accent5: string;
        var listTable5Dark_Accent6: string;
        var listTable6Colorful: string;
        var listTable6Colorful_Accent1: string;
        var listTable6Colorful_Accent2: string;
        var listTable6Colorful_Accent3: string;
        var listTable6Colorful_Accent4: string;
        var listTable6Colorful_Accent5: string;
        var listTable6Colorful_Accent6: string;
        var listTable7Colorful: string;
        var listTable7Colorful_Accent1: string;
        var listTable7Colorful_Accent2: string;
        var listTable7Colorful_Accent3: string;
        var listTable7Colorful_Accent4: string;
        var listTable7Colorful_Accent5: string;
        var listTable7Colorful_Accent6: string;
    }
    /**
     * [Api set: WordApi]
     */
    module DocumentPropertyType {
        var string: string;
        var number: string;
        var date: string;
        var boolean: string;
    }
    module ErrorCodes {
        var accessDenied: string;
        var generalException: string;
        var invalidArgument: string;
        var itemNotFound: string;
        var notImplemented: string;
    }
}
declare module Word {
    /**
     * The RequestContext object facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the request context is required to get access to the Word object model from the add-in.
     */
    class RequestContext extends OfficeExtension.ClientRequestContext {
        private m_document;
        private m_application;
        constructor(url?: string);
        document: Document;
        application: Application;
    }
    /**
     * Executes a batch script that performs actions on the Word object model, using a new RequestContext. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the RequestContext is required to get access to the Word object model from the add-in.
     */
    function run<T>(batch: (context: Word.RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
    /**
     * Executes a batch script that performs actions on the Word object model, using the RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param object - A previously-created API object. The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the RequestContext is required to get access to the Word object model from the add-in.
     */
    function run<T>(object: OfficeExtension.ClientObject, batch: (context: Word.RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
    /**
     * Executes a batch script that performs actions on the Word object model, using the RequestContext of previously-created API objects.
     * @param objects - An array of previously-created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared RequestContext, which means that any changes applied to these objects will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the RequestContext is required to get access to the Word object model from the add-in.
     */
    function run<T>(objects: OfficeExtension.ClientObject[], batch: (context: Word.RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
}
