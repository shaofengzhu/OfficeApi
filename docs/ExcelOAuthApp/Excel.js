var __extends = this.__extends || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    __.prototype = b.prototype;
    d.prototype = new __();
};
var Excel;
(function (Excel) {
    function lowerCaseFirst(str) {
        return str[0].toLowerCase() + str.slice(1);
    }
    var iconSets = ["ThreeArrows", "ThreeArrowsGray", "ThreeFlags", "ThreeTrafficLights1", "ThreeTrafficLights2", "ThreeSigns", "ThreeSymbols", "ThreeSymbols2", "FourArrows", "FourArrowsGray", "FourRedToBlack", "FourRating", "FourTrafficLights", "FiveArrows", "FiveArrowsGray", "FiveRating", "FiveQuarters", "ThreeStars", "ThreeTriangles", "FiveBoxes"];
    var iconNames = [["RedDownArrow", "YellowSideArrow", "GreenUpArrow"], ["GrayDownArrow", "GraySideArrow", "GrayUpArrow"], ["RedFlag", "YellowFlag", "GreenFlag"], ["RedCircleWithBorder", "YellowCircle", "GreenCircle"], ["RedTrafficLight", "YellowTrafficLight", "GreenTrafficLight"], ["RedDiamond", "YellowTriangle", "GreenCircle"], ["RedCrossSymbol", "YellowExclamationSymbol", "GreenCheckSymbol"], ["RedCross", "YellowExclamation", "GreenCheck"], ["RedDownArrow", "YellowDownInclineArrow", "YellowUpInclineArrow", "GreenUpArrow"], ["GrayDownArrow", "GrayDownInclineArrow", "GrayUpInclineArrow", "GrayUpArrow"], ["BlackCircle", "GrayCircle", "PinkCircle", "RedCircle"], ["OneBar", "TwoBars", "ThreeBars", "FourBars"], ["BlackCircleWithBorder", "RedCircleWithBorder", "YellowCircle", "GreenCircle"], ["RedDownArrow", "YellowDownInclineArrow", "YellowSideArrow", "YellowUpInclineArrow", "GreenUpArrow"], ["GrayDownArrow", "GrayDownInclineArrow", "GraySideArrow", "GrayUpInclineArrow", "GrayUpArrow"], ["NoBars", "OneBar", "TwoBars", "ThreeBars", "FourBars"], ["WhiteCircleAllWhiteQuarters", "CircleWithThreeWhiteQuarters", "CircleWithTwoWhiteQuarters", "CircleWithOneWhiteQuarter", "BlackCircle"], ["SilverStar", "HalfGoldStar", "GoldStar"], ["RedDownTriangle", "YellowDash", "GreenUpTriangle"], ["NoFilledBoxes", "OneFilledBox", "TwoFilledBoxes", "ThreeFilledBoxes", "FourFilledBoxes"],];
    Excel.icons = {};
    iconSets.map(function (title, i) {
        var camelTitle = lowerCaseFirst(title);
        Excel.icons[camelTitle] = [];
        iconNames[i].map(function (iconName, j) {
            iconName = lowerCaseFirst(iconName);
            var obj = { set: title, index: j };
            Excel.icons[camelTitle].push(obj);
            Excel.icons[camelTitle][iconName] = obj;
        });
    });
    function setRangePropertiesInBulk(range, propertyName, values) {
        var maxCellCount = 1500;
        if (Array.isArray(values) && values.length > 0 && Array.isArray(values[0]) && (values.length * values[0].length > maxCellCount) && isExcel1_3OrAbove()) {
            var maxRowCount = Math.max(1, Math.round(maxCellCount / values[0].length));
            range._ValidateArraySize(values.length, values[0].length);
            for (var startRowIndex = 0; startRowIndex < values.length; startRowIndex += maxRowCount) {
                var rowCount = maxRowCount;
                if (startRowIndex + rowCount > values.length) {
                    rowCount = values.length - startRowIndex;
                }
                var chunk = range.getRow(startRowIndex).getBoundingRect(range.getRow(startRowIndex + rowCount - 1));
                var valueSlice = values.slice(startRowIndex, startRowIndex + rowCount);
                _createSetPropertyAction(chunk.context, chunk, propertyName, valueSlice);
            }
            return true;
        }
        return false;
    }
    function isExcel1_3OrAbove() {
        if (typeof (window) !== "undefined" && window.Office && window.Office.context && window.Office.context.requirements) {
            return window.Office.context.requirements.isSetSupported("ExcelApi", 1.3);
        }
        return true;
    }
    var _createPropertyObjectPath = OfficeExtension.ObjectPathFactory.createPropertyObjectPath;
    var _createMethodObjectPath = OfficeExtension.ObjectPathFactory.createMethodObjectPath;
    var _createIndexerObjectPath = OfficeExtension.ObjectPathFactory.createIndexerObjectPath;
    var _createNewObjectObjectPath = OfficeExtension.ObjectPathFactory.createNewObjectObjectPath;
    var _createChildItemObjectPathUsingIndexer = OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer;
    var _createChildItemObjectPathUsingGetItemAt = OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt;
    var _createChildItemObjectPathUsingIndexerOrGetItemAt = OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt;
    var _createMethodAction = OfficeExtension.ActionFactory.createMethodAction;
    var _createSetPropertyAction = OfficeExtension.ActionFactory.createSetPropertyAction;
    var _isNullOrUndefined = OfficeExtension.Utility.isNullOrUndefined;
    var _isUndefined = OfficeExtension.Utility.isUndefined;
    var _throwIfNotLoaded = OfficeExtension.Utility.throwIfNotLoaded;
    var _load = OfficeExtension.Utility.load;
    var _fixObjectPathIfNecessary = OfficeExtension.Utility.fixObjectPathIfNecessary;
    var _addActionResultHandler = OfficeExtension.Utility._addActionResultHandler;
    var _handleNavigationPropertyResults = OfficeExtension.Utility._handleNavigationPropertyResults;
    var _adjustToDateTime = OfficeExtension.Utility.adjustToDateTime;
    var Application = (function (_super) {
        __extends(Application, _super);
        function Application() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Application.prototype, "calculationMode", {
            get: function () {
                _throwIfNotLoaded("calculationMode", this.m_calculationMode, "Application", this._isNull);
                return this.m_calculationMode;
            },
            enumerable: true,
            configurable: true
        });
        Application.prototype.calculate = function (calculationType) {
            _createMethodAction(this.context, this, "Calculate", 0 /* Default */, [calculationType]);
        };
        Application.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["CalculationMode"])) {
                this.m_calculationMode = obj["CalculationMode"];
            }
        };
        Application.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return Application;
    })(OfficeExtension.ClientObject);
    Excel.Application = Application;
    var Workbook = (function (_super) {
        __extends(Workbook, _super);
        function Workbook() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Workbook.prototype, "application", {
            get: function () {
                if (!this.m_application) {
                    this.m_application = new Excel.Application(this.context, _createPropertyObjectPath(this.context, this, "Application", false, false));
                }
                return this.m_application;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "bindings", {
            get: function () {
                if (!this.m_bindings) {
                    this.m_bindings = new Excel.BindingCollection(this.context, _createPropertyObjectPath(this.context, this, "Bindings", true, false));
                }
                return this.m_bindings;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "functions", {
            get: function () {
                if (!this.m_functions) {
                    this.m_functions = new Excel.Functions(this.context, _createPropertyObjectPath(this.context, this, "Functions", false, false));
                }
                return this.m_functions;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "names", {
            get: function () {
                if (!this.m_names) {
                    this.m_names = new Excel.NamedItemCollection(this.context, _createPropertyObjectPath(this.context, this, "Names", true, false));
                }
                return this.m_names;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "pivotTables", {
            get: function () {
                if (!this.m_pivotTables) {
                    this.m_pivotTables = new Excel.PivotTableCollection(this.context, _createPropertyObjectPath(this.context, this, "PivotTables", true, false));
                }
                return this.m_pivotTables;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "settings", {
            get: function () {
                if (!this.m_settings) {
                    this.m_settings = new Excel.SettingCollection(this.context, _createPropertyObjectPath(this.context, this, "Settings", true, false));
                }
                return this.m_settings;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "tables", {
            get: function () {
                if (!this.m_tables) {
                    this.m_tables = new Excel.TableCollection(this.context, _createPropertyObjectPath(this.context, this, "Tables", true, false));
                }
                return this.m_tables;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "worksheets", {
            get: function () {
                if (!this.m_worksheets) {
                    this.m_worksheets = new Excel.WorksheetCollection(this.context, _createPropertyObjectPath(this.context, this, "Worksheets", true, false));
                }
                return this.m_worksheets;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "_V1Api", {
            get: function () {
                if (!this.m__V1Api) {
                    this.m__V1Api = new Excel._V1Api(this.context, _createPropertyObjectPath(this.context, this, "_V1Api", false, false));
                }
                return this.m__V1Api;
            },
            enumerable: true,
            configurable: true
        });
        Workbook.prototype.getSelectedRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetSelectedRange", 1 /* Read */, [], false, true, null));
        };
        Workbook.prototype._GetObjectByReferenceId = function (bstrReferenceId) {
            var action = _createMethodAction(this.context, this, "_GetObjectByReferenceId", 1 /* Read */, [bstrReferenceId]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        Workbook.prototype._GetObjectTypeNameByReferenceId = function (bstrReferenceId) {
            var action = _createMethodAction(this.context, this, "_GetObjectTypeNameByReferenceId", 1 /* Read */, [bstrReferenceId]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        Workbook.prototype._GetReferenceCount = function () {
            var action = _createMethodAction(this.context, this, "_GetReferenceCount", 1 /* Read */, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        Workbook.prototype._RemoveAllReferences = function () {
            _createMethodAction(this.context, this, "_RemoveAllReferences", 1 /* Read */, []);
        };
        Workbook.prototype._RemoveReference = function (bstrReferenceId) {
            _createMethodAction(this.context, this, "_RemoveReference", 1 /* Read */, [bstrReferenceId]);
        };
        Workbook.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["application", "Application", "bindings", "Bindings", "functions", "Functions", "names", "Names", "pivotTables", "PivotTables", "settings", "Settings", "tables", "Tables", "worksheets", "Worksheets", "_V1Api", "_V1Api"]);
        };
        Workbook.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        Object.defineProperty(Workbook.prototype, "onSelectionChanged", {
            get: function () {
                var _this = this;
                if (!this.m_selectionChanged) {
                    this.m_selectionChanged = new OfficeExtension.EventHandlers(this.context, this, "SelectionChanged", {
                        registerFunc: function (handlerCallback) {
                            return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handlerCallback, callback); });
                        },
                        unregisterFunc: function (handlerCallback) {
                            return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: handlerCallback }, callback); });
                        },
                        eventArgsTransformFunc: function (args) {
                            return OfficeExtension.Utility._createPromiseFromResult({ workbook: _this });
                        }
                    });
                }
                return this.m_selectionChanged;
            },
            enumerable: true,
            configurable: true
        });
        return Workbook;
    })(OfficeExtension.ClientObject);
    Excel.Workbook = Workbook;
    var Worksheet = (function (_super) {
        __extends(Worksheet, _super);
        function Worksheet() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Worksheet.prototype, "charts", {
            get: function () {
                if (!this.m_charts) {
                    this.m_charts = new Excel.ChartCollection(this.context, _createPropertyObjectPath(this.context, this, "Charts", true, false));
                }
                return this.m_charts;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "pivotTables", {
            get: function () {
                if (!this.m_pivotTables) {
                    this.m_pivotTables = new Excel.PivotTableCollection(this.context, _createPropertyObjectPath(this.context, this, "PivotTables", true, false));
                }
                return this.m_pivotTables;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "protection", {
            get: function () {
                if (!this.m_protection) {
                    this.m_protection = new Excel.WorksheetProtection(this.context, _createPropertyObjectPath(this.context, this, "Protection", false, false));
                }
                return this.m_protection;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "tables", {
            get: function () {
                if (!this.m_tables) {
                    this.m_tables = new Excel.TableCollection(this.context, _createPropertyObjectPath(this.context, this, "Tables", true, false));
                }
                return this.m_tables;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this.m_id, "Worksheet", this._isNull);
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name, "Worksheet", this._isNull);
                return this.m_name;
            },
            set: function (value) {
                this.m_name = value;
                _createSetPropertyAction(this.context, this, "Name", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "position", {
            get: function () {
                _throwIfNotLoaded("position", this.m_position, "Worksheet", this._isNull);
                return this.m_position;
            },
            set: function (value) {
                this.m_position = value;
                _createSetPropertyAction(this.context, this, "Position", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "visibility", {
            get: function () {
                _throwIfNotLoaded("visibility", this.m_visibility, "Worksheet", this._isNull);
                return this.m_visibility;
            },
            set: function (value) {
                this.m_visibility = value;
                _createSetPropertyAction(this.context, this, "Visibility", value);
            },
            enumerable: true,
            configurable: true
        });
        Worksheet.prototype.activate = function () {
            _createMethodAction(this.context, this, "Activate", 1 /* Read */, []);
        };
        Worksheet.prototype.delete = function () {
            _createMethodAction(this.context, this, "Delete", 0 /* Default */, []);
        };
        Worksheet.prototype.getCell = function (row, column) {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetCell", 1 /* Read */, [row, column], false, true, null));
        };
        Worksheet.prototype.getRange = function (address) {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1 /* Read */, [address], false, true, null));
        };
        Worksheet.prototype.getUsedRange = function (valuesOnly) {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetUsedRange", 1 /* Read */, [valuesOnly], false, true, null));
        };
        Worksheet.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Id"])) {
                this.m_id = obj["Id"];
            }
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            if (!_isUndefined(obj["Position"])) {
                this.m_position = obj["Position"];
            }
            if (!_isUndefined(obj["Visibility"])) {
                this.m_visibility = obj["Visibility"];
            }
            _handleNavigationPropertyResults(this, obj, ["charts", "Charts", "pivotTables", "PivotTables", "protection", "Protection", "tables", "Tables"]);
        };
        Worksheet.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        Worksheet.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["Id"])) {
                this.m_id = value["Id"];
            }
        };
        return Worksheet;
    })(OfficeExtension.ClientObject);
    Excel.Worksheet = Worksheet;
    var WorksheetCollection = (function (_super) {
        __extends(WorksheetCollection, _super);
        function WorksheetCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(WorksheetCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "WorksheetCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        WorksheetCollection.prototype.add = function (name) {
            return new Excel.Worksheet(this.context, _createMethodObjectPath(this.context, this, "Add", 0 /* Default */, [name], false, true, null));
        };
        WorksheetCollection.prototype.getActiveWorksheet = function () {
            return new Excel.Worksheet(this.context, _createMethodObjectPath(this.context, this, "GetActiveWorksheet", 1 /* Read */, [], false, false, null));
        };
        WorksheetCollection.prototype.getItem = function (key) {
            return new Excel.Worksheet(this.context, _createIndexerObjectPath(this.context, this, [key]));
        };
        WorksheetCollection.prototype.getItemOrNull = function (key) {
            return new Excel.Worksheet(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNull", 1 /* Read */, [key], false, false, null));
        };
        WorksheetCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.Worksheet(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        WorksheetCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return WorksheetCollection;
    })(OfficeExtension.ClientObject);
    Excel.WorksheetCollection = WorksheetCollection;
    var WorksheetProtection = (function (_super) {
        __extends(WorksheetProtection, _super);
        function WorksheetProtection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(WorksheetProtection.prototype, "options", {
            get: function () {
                _throwIfNotLoaded("options", this.m_options, "WorksheetProtection", this._isNull);
                return this.m_options;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(WorksheetProtection.prototype, "protected", {
            get: function () {
                _throwIfNotLoaded("protected", this.m_protected, "WorksheetProtection", this._isNull);
                return this.m_protected;
            },
            enumerable: true,
            configurable: true
        });
        WorksheetProtection.prototype.protect = function (options) {
            _createMethodAction(this.context, this, "Protect", 0 /* Default */, [options]);
        };
        WorksheetProtection.prototype.unprotect = function () {
            _createMethodAction(this.context, this, "Unprotect", 0 /* Default */, []);
        };
        WorksheetProtection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Options"])) {
                this.m_options = obj["Options"];
            }
            if (!_isUndefined(obj["Protected"])) {
                this.m_protected = obj["Protected"];
            }
        };
        WorksheetProtection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return WorksheetProtection;
    })(OfficeExtension.ClientObject);
    Excel.WorksheetProtection = WorksheetProtection;
    var Range = (function (_super) {
        __extends(Range, _super);
        function Range() {
            _super.apply(this, arguments);
        }
        Range.prototype._ensureInteger = function (num, methodName) {
            if (!(typeof num === "number" && isFinite(num) && Math.floor(num) === num)) {
                throw new OfficeExtension.Utility.throwError(Excel.ErrorCodes.invalidArgument, num, methodName);
            }
        };
        Range.prototype._getAdjacentRange = function (functionName, count, referenceRange, rowDirection, columnDirection) {
            if (count == null) {
                count = 1;
            }
            this._ensureInteger(count, functionName);
            var startRange;
            var rowOffset = 0;
            var columnOffset = 0;
            if (count > 0) {
                startRange = referenceRange.getOffsetRange(rowDirection, columnDirection);
            }
            else {
                startRange = referenceRange;
                rowOffset = rowDirection;
                columnOffset = columnDirection;
            }
            if (Math.abs(count) == 1) {
                return startRange;
            }
            return startRange.getBoundingRect(referenceRange.getOffsetRange(rowDirection * count + rowOffset, columnDirection * count + columnOffset));
        };
        Object.defineProperty(Range.prototype, "conditionalFormats", {
            get: function () {
                if (!this.m_conditionalFormats) {
                    this.m_conditionalFormats = new Excel.ConditionalFormatCollection(this.context, _createPropertyObjectPath(this.context, this, "ConditionalFormats", true, false));
                }
                return this.m_conditionalFormats;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.RangeFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "sort", {
            get: function () {
                if (!this.m_sort) {
                    this.m_sort = new Excel.RangeSort(this.context, _createPropertyObjectPath(this.context, this, "Sort", false, false));
                }
                return this.m_sort;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "worksheet", {
            get: function () {
                if (!this.m_worksheet) {
                    this.m_worksheet = new Excel.Worksheet(this.context, _createPropertyObjectPath(this.context, this, "Worksheet", false, false));
                }
                return this.m_worksheet;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "address", {
            get: function () {
                _throwIfNotLoaded("address", this.m_address, "Range", this._isNull);
                return this.m_address;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "addressLocal", {
            get: function () {
                _throwIfNotLoaded("addressLocal", this.m_addressLocal, "Range", this._isNull);
                return this.m_addressLocal;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "cellCount", {
            get: function () {
                _throwIfNotLoaded("cellCount", this.m_cellCount, "Range", this._isNull);
                return this.m_cellCount;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "columnCount", {
            get: function () {
                _throwIfNotLoaded("columnCount", this.m_columnCount, "Range", this._isNull);
                return this.m_columnCount;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "columnHidden", {
            get: function () {
                _throwIfNotLoaded("columnHidden", this.m_columnHidden, "Range", this._isNull);
                return this.m_columnHidden;
            },
            set: function (value) {
                this.m_columnHidden = value;
                _createSetPropertyAction(this.context, this, "ColumnHidden", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "columnIndex", {
            get: function () {
                _throwIfNotLoaded("columnIndex", this.m_columnIndex, "Range", this._isNull);
                return this.m_columnIndex;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "formulas", {
            get: function () {
                _throwIfNotLoaded("formulas", this.m_formulas, "Range", this._isNull);
                return this.m_formulas;
            },
            set: function (value) {
                this.m_formulas = value;
                if (setRangePropertiesInBulk(this, "Formulas", value)) {
                    return;
                }
                this.m_formulas = value;
                _createSetPropertyAction(this.context, this, "Formulas", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "formulasLocal", {
            get: function () {
                _throwIfNotLoaded("formulasLocal", this.m_formulasLocal, "Range", this._isNull);
                return this.m_formulasLocal;
            },
            set: function (value) {
                this.m_formulasLocal = value;
                if (setRangePropertiesInBulk(this, "FormulasLocal", value)) {
                    return;
                }
                this.m_formulasLocal = value;
                _createSetPropertyAction(this.context, this, "FormulasLocal", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "formulasR1C1", {
            get: function () {
                _throwIfNotLoaded("formulasR1C1", this.m_formulasR1C1, "Range", this._isNull);
                return this.m_formulasR1C1;
            },
            set: function (value) {
                this.m_formulasR1C1 = value;
                if (setRangePropertiesInBulk(this, "FormulasR1C1", value)) {
                    return;
                }
                this.m_formulasR1C1 = value;
                _createSetPropertyAction(this.context, this, "FormulasR1C1", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "hidden", {
            get: function () {
                _throwIfNotLoaded("hidden", this.m_hidden, "Range", this._isNull);
                return this.m_hidden;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "numberFormat", {
            get: function () {
                _throwIfNotLoaded("numberFormat", this.m_numberFormat, "Range", this._isNull);
                return this.m_numberFormat;
            },
            set: function (value) {
                this.m_numberFormat = value;
                if (setRangePropertiesInBulk(this, "NumberFormat", value)) {
                    return;
                }
                this.m_numberFormat = value;
                _createSetPropertyAction(this.context, this, "NumberFormat", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "rowCount", {
            get: function () {
                _throwIfNotLoaded("rowCount", this.m_rowCount, "Range", this._isNull);
                return this.m_rowCount;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "rowHidden", {
            get: function () {
                _throwIfNotLoaded("rowHidden", this.m_rowHidden, "Range", this._isNull);
                return this.m_rowHidden;
            },
            set: function (value) {
                this.m_rowHidden = value;
                _createSetPropertyAction(this.context, this, "RowHidden", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "rowIndex", {
            get: function () {
                _throwIfNotLoaded("rowIndex", this.m_rowIndex, "Range", this._isNull);
                return this.m_rowIndex;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "text", {
            get: function () {
                _throwIfNotLoaded("text", this.m_text, "Range", this._isNull);
                return this.m_text;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "valueTypes", {
            get: function () {
                _throwIfNotLoaded("valueTypes", this.m_valueTypes, "Range", this._isNull);
                return this.m_valueTypes;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "values", {
            get: function () {
                _throwIfNotLoaded("values", this.m_values, "Range", this._isNull);
                return this.m_values;
            },
            set: function (value) {
                this.m_values = value;
                if (setRangePropertiesInBulk(this, "Values", value)) {
                    return;
                }
                this.m_values = value;
                _createSetPropertyAction(this.context, this, "Values", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "_ReferenceId", {
            get: function () {
                _throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "Range", this._isNull);
                return this.m__ReferenceId;
            },
            enumerable: true,
            configurable: true
        });
        Range.prototype.clear = function (applyTo) {
            _createMethodAction(this.context, this, "Clear", 0 /* Default */, [applyTo]);
        };
        Range.prototype.delete = function (shift) {
            _createMethodAction(this.context, this, "Delete", 0 /* Default */, [shift]);
        };
        Range.prototype.getBoundingRect = function (anotherRange) {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetBoundingRect", 1 /* Read */, [anotherRange], false, true, null));
        };
        Range.prototype.getCell = function (row, column) {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetCell", 1 /* Read */, [row, column], false, true, null));
        };
        Range.prototype.getColumn = function (column) {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetColumn", 1 /* Read */, [column], false, true, null));
        };
        Range.prototype.getColumnsAfter = function (count) {
            if (!isExcel1_3OrAbove()) {
                if (count == null) {
                    count = 1;
                }
                this._ensureInteger(count, "RowsAbove");
                if (count == 0) {
                    throw new OfficeExtension.Utility.throwError(Excel.ErrorCodes.invalidArgument, "count", "RowsAbove");
                }
                return this._getAdjacentRange("getColumnsAfter", count, this.getLastColumn(), 0, 1);
            }
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetColumnsAfter", 1 /* Read */, [count], false, true, null));
        };
        Range.prototype.getColumnsBefore = function (count) {
            if (!isExcel1_3OrAbove()) {
                if (count == null) {
                    count = 1;
                }
                this._ensureInteger(count, "RowsAbove");
                if (count == 0) {
                    throw new OfficeExtension.Utility.throwError(Excel.ErrorCodes.invalidArgument, "count", "RowsAbove");
                }
                return this._getAdjacentRange("getColumnsBefore", count, this.getColumn(0), 0, -1);
            }
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetColumnsBefore", 1 /* Read */, [count], false, true, null));
        };
        Range.prototype.getEntireColumn = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetEntireColumn", 1 /* Read */, [], false, true, null));
        };
        Range.prototype.getEntireRow = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetEntireRow", 1 /* Read */, [], false, true, null));
        };
        Range.prototype.getIntersection = function (anotherRange) {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetIntersection", 1 /* Read */, [anotherRange], false, true, null));
        };
        Range.prototype.getIntersectionOrNull = function (anotherRange) {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetIntersectionOrNull", 1 /* Read */, [anotherRange], false, true, null));
        };
        Range.prototype.getLastCell = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetLastCell", 1 /* Read */, [], false, true, null));
        };
        Range.prototype.getLastColumn = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetLastColumn", 1 /* Read */, [], false, true, null));
        };
        Range.prototype.getLastRow = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetLastRow", 1 /* Read */, [], false, true, null));
        };
        Range.prototype.getOffsetRange = function (rowOffset, columnOffset) {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetOffsetRange", 1 /* Read */, [rowOffset, columnOffset], false, true, null));
        };
        Range.prototype.getResizedRange = function (deltaRows, deltaColumns) {
            if (!isExcel1_3OrAbove()) {
                this._ensureInteger(deltaRows, "getResizedRange");
                this._ensureInteger(deltaColumns, "getResizedRange");
                var referenceRange = (deltaRows >= 0 && deltaColumns >= 0) ? this : this.getCell(0, 0);
                return referenceRange.getBoundingRect(this.getLastCell().getOffsetRange(deltaRows, deltaColumns));
            }
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetResizedRange", 1 /* Read */, [deltaRows, deltaColumns], false, true, null));
        };
        Range.prototype.getRow = function (row) {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRow", 1 /* Read */, [row], false, true, null));
        };
        Range.prototype.getRowsAbove = function (count) {
            if (!isExcel1_3OrAbove()) {
                if (count == null) {
                    count = 1;
                }
                this._ensureInteger(count, "RowsAbove");
                if (count == 0) {
                    throw new OfficeExtension.Utility.throwError(Excel.ErrorCodes.invalidArgument, "count", "RowsAbove");
                }
                return this._getAdjacentRange("getRowsAbove", count, this.getRow(0), -1, 0);
            }
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRowsAbove", 1 /* Read */, [count], false, true, null));
        };
        Range.prototype.getRowsBelow = function (count) {
            if (!isExcel1_3OrAbove()) {
                if (count == null) {
                    count = 1;
                }
                this._ensureInteger(count, "RowsAbove");
                if (count == 0) {
                    throw new OfficeExtension.Utility.throwError(Excel.ErrorCodes.invalidArgument, "count", "RowsAbove");
                }
                return this._getAdjacentRange("getRowsBelow", count, this.getLastRow(), 1, 0);
            }
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRowsBelow", 1 /* Read */, [count], false, true, null));
        };
        Range.prototype.getUsedRange = function (valuesOnly) {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetUsedRange", 1 /* Read */, [valuesOnly], false, true, null));
        };
        Range.prototype.getVisibleView = function () {
            return new Excel.RangeView(this.context, _createMethodObjectPath(this.context, this, "GetVisibleView", 1 /* Read */, [], false, false, null));
        };
        Range.prototype.insert = function (shift) {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "Insert", 0 /* Default */, [shift], false, true, null));
        };
        Range.prototype.merge = function (across) {
            _createMethodAction(this.context, this, "Merge", 0 /* Default */, [across]);
        };
        Range.prototype.select = function () {
            _createMethodAction(this.context, this, "Select", 1 /* Read */, []);
        };
        Range.prototype.unmerge = function () {
            _createMethodAction(this.context, this, "Unmerge", 0 /* Default */, []);
        };
        Range.prototype._KeepReference = function () {
            _createMethodAction(this.context, this, "_KeepReference", 1 /* Read */, []);
        };
        Range.prototype._ValidateArraySize = function (rows, columns) {
            _createMethodAction(this.context, this, "_ValidateArraySize", 1 /* Read */, [rows, columns]);
        };
        Range.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Address"])) {
                this.m_address = obj["Address"];
            }
            if (!_isUndefined(obj["AddressLocal"])) {
                this.m_addressLocal = obj["AddressLocal"];
            }
            if (!_isUndefined(obj["CellCount"])) {
                this.m_cellCount = obj["CellCount"];
            }
            if (!_isUndefined(obj["ColumnCount"])) {
                this.m_columnCount = obj["ColumnCount"];
            }
            if (!_isUndefined(obj["ColumnHidden"])) {
                this.m_columnHidden = obj["ColumnHidden"];
            }
            if (!_isUndefined(obj["ColumnIndex"])) {
                this.m_columnIndex = obj["ColumnIndex"];
            }
            if (!_isUndefined(obj["Formulas"])) {
                this.m_formulas = obj["Formulas"];
            }
            if (!_isUndefined(obj["FormulasLocal"])) {
                this.m_formulasLocal = obj["FormulasLocal"];
            }
            if (!_isUndefined(obj["FormulasR1C1"])) {
                this.m_formulasR1C1 = obj["FormulasR1C1"];
            }
            if (!_isUndefined(obj["Hidden"])) {
                this.m_hidden = obj["Hidden"];
            }
            if (!_isUndefined(obj["NumberFormat"])) {
                this.m_numberFormat = obj["NumberFormat"];
            }
            if (!_isUndefined(obj["RowCount"])) {
                this.m_rowCount = obj["RowCount"];
            }
            if (!_isUndefined(obj["RowHidden"])) {
                this.m_rowHidden = obj["RowHidden"];
            }
            if (!_isUndefined(obj["RowIndex"])) {
                this.m_rowIndex = obj["RowIndex"];
            }
            if (!_isUndefined(obj["Text"])) {
                this.m_text = obj["Text"];
            }
            if (!_isUndefined(obj["ValueTypes"])) {
                this.m_valueTypes = obj["ValueTypes"];
            }
            if (!_isUndefined(obj["Values"])) {
                this.m_values = obj["Values"];
            }
            if (!_isUndefined(obj["_ReferenceId"])) {
                this.m__ReferenceId = obj["_ReferenceId"];
            }
            _handleNavigationPropertyResults(this, obj, ["conditionalFormats", "ConditionalFormats", "format", "Format", "sort", "Sort", "worksheet", "Worksheet"]);
        };
        Range.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        Range.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["_ReferenceId"])) {
                this.m__ReferenceId = value["_ReferenceId"];
            }
        };
        return Range;
    })(OfficeExtension.ClientObject);
    Excel.Range = Range;
    var RangeView = (function (_super) {
        __extends(RangeView, _super);
        function RangeView() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(RangeView.prototype, "rows", {
            get: function () {
                if (!this.m_rows) {
                    this.m_rows = new Excel.RangeViewCollection(this.context, _createPropertyObjectPath(this.context, this, "Rows", true, false));
                }
                return this.m_rows;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeView.prototype, "cellAddresses", {
            get: function () {
                _throwIfNotLoaded("cellAddresses", this.m_cellAddresses, "RangeView", this._isNull);
                return this.m_cellAddresses;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeView.prototype, "columnCount", {
            get: function () {
                _throwIfNotLoaded("columnCount", this.m_columnCount, "RangeView", this._isNull);
                return this.m_columnCount;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeView.prototype, "formulas", {
            get: function () {
                _throwIfNotLoaded("formulas", this.m_formulas, "RangeView", this._isNull);
                return this.m_formulas;
            },
            set: function (value) {
                this.m_formulas = value;
                _createSetPropertyAction(this.context, this, "Formulas", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeView.prototype, "formulasLocal", {
            get: function () {
                _throwIfNotLoaded("formulasLocal", this.m_formulasLocal, "RangeView", this._isNull);
                return this.m_formulasLocal;
            },
            set: function (value) {
                this.m_formulasLocal = value;
                _createSetPropertyAction(this.context, this, "FormulasLocal", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeView.prototype, "formulasR1C1", {
            get: function () {
                _throwIfNotLoaded("formulasR1C1", this.m_formulasR1C1, "RangeView", this._isNull);
                return this.m_formulasR1C1;
            },
            set: function (value) {
                this.m_formulasR1C1 = value;
                _createSetPropertyAction(this.context, this, "FormulasR1C1", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeView.prototype, "index", {
            get: function () {
                _throwIfNotLoaded("index", this.m_index, "RangeView", this._isNull);
                return this.m_index;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeView.prototype, "numberFormat", {
            get: function () {
                _throwIfNotLoaded("numberFormat", this.m_numberFormat, "RangeView", this._isNull);
                return this.m_numberFormat;
            },
            set: function (value) {
                this.m_numberFormat = value;
                _createSetPropertyAction(this.context, this, "NumberFormat", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeView.prototype, "rowCount", {
            get: function () {
                _throwIfNotLoaded("rowCount", this.m_rowCount, "RangeView", this._isNull);
                return this.m_rowCount;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeView.prototype, "text", {
            get: function () {
                _throwIfNotLoaded("text", this.m_text, "RangeView", this._isNull);
                return this.m_text;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeView.prototype, "valueTypes", {
            get: function () {
                _throwIfNotLoaded("valueTypes", this.m_valueTypes, "RangeView", this._isNull);
                return this.m_valueTypes;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeView.prototype, "values", {
            get: function () {
                _throwIfNotLoaded("values", this.m_values, "RangeView", this._isNull);
                return this.m_values;
            },
            set: function (value) {
                this.m_values = value;
                _createSetPropertyAction(this.context, this, "Values", value);
            },
            enumerable: true,
            configurable: true
        });
        RangeView.prototype.getRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1 /* Read */, [], false, true, null));
        };
        RangeView.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["CellAddresses"])) {
                this.m_cellAddresses = obj["CellAddresses"];
            }
            if (!_isUndefined(obj["ColumnCount"])) {
                this.m_columnCount = obj["ColumnCount"];
            }
            if (!_isUndefined(obj["Formulas"])) {
                this.m_formulas = obj["Formulas"];
            }
            if (!_isUndefined(obj["FormulasLocal"])) {
                this.m_formulasLocal = obj["FormulasLocal"];
            }
            if (!_isUndefined(obj["FormulasR1C1"])) {
                this.m_formulasR1C1 = obj["FormulasR1C1"];
            }
            if (!_isUndefined(obj["Index"])) {
                this.m_index = obj["Index"];
            }
            if (!_isUndefined(obj["NumberFormat"])) {
                this.m_numberFormat = obj["NumberFormat"];
            }
            if (!_isUndefined(obj["RowCount"])) {
                this.m_rowCount = obj["RowCount"];
            }
            if (!_isUndefined(obj["Text"])) {
                this.m_text = obj["Text"];
            }
            if (!_isUndefined(obj["ValueTypes"])) {
                this.m_valueTypes = obj["ValueTypes"];
            }
            if (!_isUndefined(obj["Values"])) {
                this.m_values = obj["Values"];
            }
            _handleNavigationPropertyResults(this, obj, ["rows", "Rows"]);
        };
        RangeView.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return RangeView;
    })(OfficeExtension.ClientObject);
    Excel.RangeView = RangeView;
    var RangeViewCollection = (function (_super) {
        __extends(RangeViewCollection, _super);
        function RangeViewCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(RangeViewCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "RangeViewCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        RangeViewCollection.prototype.getItemAt = function (index) {
            return new Excel.RangeView(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [index], false, false, null));
        };
        RangeViewCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.RangeView(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(false, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        RangeViewCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return RangeViewCollection;
    })(OfficeExtension.ClientObject);
    Excel.RangeViewCollection = RangeViewCollection;
    var SettingCollection = (function (_super) {
        __extends(SettingCollection, _super);
        function SettingCollection() {
            _super.apply(this, arguments);
        }
        SettingCollection.prototype.add = function (key, value) {
            return this._Add(key, JSON.stringify(value));
        };
        Object.defineProperty(SettingCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "SettingCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        SettingCollection.prototype.getItem = function (key) {
            return new Excel.Setting(this.context, _createIndexerObjectPath(this.context, this, [key]));
        };
        SettingCollection.prototype.getItemOrNull = function (key) {
            return new Excel.Setting(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNull", 1 /* Read */, [key], false, false, null));
        };
        SettingCollection.prototype._Add = function (key, value) {
            return new Excel.Setting(this.context, _createMethodObjectPath(this.context, this, "_Add", 0 /* Default */, [key, value], false, false, null));
        };
        SettingCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.Setting(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        SettingCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        Object.defineProperty(SettingCollection.prototype, "onSettingsChanged", {
            get: function () {
                var _this = this;
                if (!this.m_settingsChanged) {
                    this.m_settingsChanged = new OfficeExtension.EventHandlers(this.context, this, "SettingsChanged", {
                        registerFunc: function (handlerCallback) {
                            return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.settings.addHandlerAsync(Office.EventType.SettingsChanged, handlerCallback, callback); });
                        },
                        unregisterFunc: function (handlerCallback) {
                            return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.settings.removeHandlerAsync(Office.EventType.SettingsChanged, { handler: handlerCallback }, callback); });
                        },
                        eventArgsTransformFunc: function (args) {
                            return OfficeExtension.Utility._createPromiseFromResult({ settings: _this });
                        }
                    });
                }
                return this.m_settingsChanged;
            },
            enumerable: true,
            configurable: true
        });
        return SettingCollection;
    })(OfficeExtension.ClientObject);
    Excel.SettingCollection = SettingCollection;
    var Setting = (function (_super) {
        __extends(Setting, _super);
        function Setting() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Setting.prototype, "value", {
            get: function () {
                return JSON.parse(this._Value);
            },
            set: function (value) {
                this._Value = JSON.stringify(value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Setting.prototype, "key", {
            get: function () {
                _throwIfNotLoaded("key", this.m_key, "Setting", this._isNull);
                return this.m_key;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Setting.prototype, "_Value", {
            get: function () {
                _throwIfNotLoaded("_Value", this.m__Value, "Setting", this._isNull);
                return this.m__Value;
            },
            set: function (value) {
                this.m__Value = value;
                _createSetPropertyAction(this.context, this, "_Value", value);
            },
            enumerable: true,
            configurable: true
        });
        Setting.prototype.delete = function () {
            _createMethodAction(this.context, this, "Delete", 0 /* Default */, []);
        };
        Setting.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Key"])) {
                this.m_key = obj["Key"];
            }
            if (!_isUndefined(obj["_Value"])) {
                this.m__Value = obj["_Value"];
            }
        };
        Setting.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return Setting;
    })(OfficeExtension.ClientObject);
    Excel.Setting = Setting;
    var NamedItemCollection = (function (_super) {
        __extends(NamedItemCollection, _super);
        function NamedItemCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(NamedItemCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "NamedItemCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        NamedItemCollection.prototype.getItem = function (name) {
            return new Excel.NamedItem(this.context, _createIndexerObjectPath(this.context, this, [name]));
        };
        NamedItemCollection.prototype.getItemOrNull = function (name) {
            return new Excel.NamedItem(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNull", 1 /* Read */, [name], false, false, null));
        };
        NamedItemCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.NamedItem(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        NamedItemCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return NamedItemCollection;
    })(OfficeExtension.ClientObject);
    Excel.NamedItemCollection = NamedItemCollection;
    var NamedItem = (function (_super) {
        __extends(NamedItem, _super);
        function NamedItem() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(NamedItem.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name, "NamedItem", this._isNull);
                return this.m_name;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(NamedItem.prototype, "type", {
            get: function () {
                _throwIfNotLoaded("type", this.m_type, "NamedItem", this._isNull);
                return this.m_type;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(NamedItem.prototype, "value", {
            get: function () {
                _throwIfNotLoaded("value", this.m_value, "NamedItem", this._isNull);
                return this.m_value;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(NamedItem.prototype, "visible", {
            get: function () {
                _throwIfNotLoaded("visible", this.m_visible, "NamedItem", this._isNull);
                return this.m_visible;
            },
            set: function (value) {
                this.m_visible = value;
                _createSetPropertyAction(this.context, this, "Visible", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(NamedItem.prototype, "_Id", {
            get: function () {
                _throwIfNotLoaded("_Id", this.m__Id, "NamedItem", this._isNull);
                return this.m__Id;
            },
            enumerable: true,
            configurable: true
        });
        NamedItem.prototype.getRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1 /* Read */, [], false, true, null));
        };
        NamedItem.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            if (!_isUndefined(obj["Type"])) {
                this.m_type = obj["Type"];
            }
            if (!_isUndefined(obj["Value"])) {
                this.m_value = obj["Value"];
            }
            if (!_isUndefined(obj["Visible"])) {
                this.m_visible = obj["Visible"];
            }
            if (!_isUndefined(obj["_Id"])) {
                this.m__Id = obj["_Id"];
            }
        };
        NamedItem.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        NamedItem.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["_Id"])) {
                this.m__Id = value["_Id"];
            }
        };
        return NamedItem;
    })(OfficeExtension.ClientObject);
    Excel.NamedItem = NamedItem;
    var Binding = (function (_super) {
        __extends(Binding, _super);
        function Binding() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Binding.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this.m_id, "Binding", this._isNull);
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Binding.prototype, "type", {
            get: function () {
                _throwIfNotLoaded("type", this.m_type, "Binding", this._isNull);
                return this.m_type;
            },
            enumerable: true,
            configurable: true
        });
        Binding.prototype.delete = function () {
            _createMethodAction(this.context, this, "Delete", 0 /* Default */, []);
        };
        Binding.prototype.getRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1 /* Read */, [], false, false, null));
        };
        Binding.prototype.getTable = function () {
            return new Excel.Table(this.context, _createMethodObjectPath(this.context, this, "GetTable", 1 /* Read */, [], false, false, null));
        };
        Binding.prototype.getText = function () {
            var action = _createMethodAction(this.context, this, "GetText", 1 /* Read */, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        Binding.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Id"])) {
                this.m_id = obj["Id"];
            }
            if (!_isUndefined(obj["Type"])) {
                this.m_type = obj["Type"];
            }
        };
        Binding.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        Binding.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["Id"])) {
                this.m_id = value["Id"];
            }
        };
        Object.defineProperty(Binding.prototype, "onDataChanged", {
            get: function () {
                var _this = this;
                if (!this.m_dataChanged) {
                    this.m_dataChanged = new OfficeExtension.EventHandlers(this.context, this, "DataChanged", {
                        registerFunc: function (handlerCallback) {
                            return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(_this.id, callback); }).then(function (officeBinding) {
                                return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.addHandlerAsync(Office.EventType.BindingDataChanged, handlerCallback, callback); });
                            });
                        },
                        unregisterFunc: function (handlerCallback) {
                            return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(_this.id, callback); }).then(function (officeBinding) {
                                return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.removeHandlerAsync(Office.EventType.BindingDataChanged, { handler: handlerCallback }, callback); });
                            });
                        },
                        eventArgsTransformFunc: function (args) {
                            var evt = {
                                binding: _this
                            };
                            return OfficeExtension.Utility._createPromiseFromResult(evt);
                        }
                    });
                }
                return this.m_dataChanged;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Binding.prototype, "onSelectionChanged", {
            get: function () {
                var _this = this;
                if (!this.m_selectionChanged) {
                    this.m_selectionChanged = new OfficeExtension.EventHandlers(this.context, this, "SelectionChanged", {
                        registerFunc: function (handlerCallback) {
                            return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(_this.id, callback); }).then(function (officeBinding) {
                                return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.addHandlerAsync(Office.EventType.BindingSelectionChanged, handlerCallback, callback); });
                            });
                        },
                        unregisterFunc: function (handlerCallback) {
                            return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(_this.id, callback); }).then(function (officeBinding) {
                                return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.removeHandlerAsync(Office.EventType.BindingSelectionChanged, { handler: handlerCallback }, callback); });
                            });
                        },
                        eventArgsTransformFunc: function (args) {
                            var evt = {
                                binding: _this,
                                columnCount: args.columnCount,
                                rowCount: args.rowCount,
                                startColumn: args.startColumn,
                                startRow: args.startRow
                            };
                            return OfficeExtension.Utility._createPromiseFromResult(evt);
                        }
                    });
                }
                return this.m_selectionChanged;
            },
            enumerable: true,
            configurable: true
        });
        return Binding;
    })(OfficeExtension.ClientObject);
    Excel.Binding = Binding;
    var BindingCollection = (function (_super) {
        __extends(BindingCollection, _super);
        function BindingCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(BindingCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "BindingCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(BindingCollection.prototype, "count", {
            get: function () {
                _throwIfNotLoaded("count", this.m_count, "BindingCollection", this._isNull);
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });
        BindingCollection.prototype.add = function (range, bindingType, id) {
            return new Excel.Binding(this.context, _createMethodObjectPath(this.context, this, "Add", 0 /* Default */, [range, bindingType, id], false, true, null));
        };
        BindingCollection.prototype.addFromNamedItem = function (name, bindingType, id) {
            return new Excel.Binding(this.context, _createMethodObjectPath(this.context, this, "AddFromNamedItem", 0 /* Default */, [name, bindingType, id], false, false, null));
        };
        BindingCollection.prototype.addFromSelection = function (bindingType, id) {
            return new Excel.Binding(this.context, _createMethodObjectPath(this.context, this, "AddFromSelection", 0 /* Default */, [bindingType, id], false, false, null));
        };
        BindingCollection.prototype.getItem = function (id) {
            return new Excel.Binding(this.context, _createIndexerObjectPath(this.context, this, [id]));
        };
        BindingCollection.prototype.getItemAt = function (index) {
            return new Excel.Binding(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [index], false, false, null));
        };
        BindingCollection.prototype.getItemOrNull = function (id) {
            return new Excel.Binding(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNull", 1 /* Read */, [id], false, false, null));
        };
        BindingCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.Binding(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        BindingCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return BindingCollection;
    })(OfficeExtension.ClientObject);
    Excel.BindingCollection = BindingCollection;
    var TableCollection = (function (_super) {
        __extends(TableCollection, _super);
        function TableCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(TableCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "TableCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableCollection.prototype, "count", {
            get: function () {
                _throwIfNotLoaded("count", this.m_count, "TableCollection", this._isNull);
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });
        TableCollection.prototype.add = function (address, hasHeaders) {
            return new Excel.Table(this.context, _createMethodObjectPath(this.context, this, "Add", 0 /* Default */, [address, hasHeaders], false, true, null));
        };
        TableCollection.prototype.getItem = function (key) {
            return new Excel.Table(this.context, _createIndexerObjectPath(this.context, this, [key]));
        };
        TableCollection.prototype.getItemAt = function (index) {
            return new Excel.Table(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [index], false, false, null));
        };
        TableCollection.prototype.getItemOrNull = function (key) {
            return new Excel.Table(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNull", 1 /* Read */, [key], false, false, null));
        };
        TableCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.Table(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        TableCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return TableCollection;
    })(OfficeExtension.ClientObject);
    Excel.TableCollection = TableCollection;
    var Table = (function (_super) {
        __extends(Table, _super);
        function Table() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Table.prototype, "columns", {
            get: function () {
                if (!this.m_columns) {
                    this.m_columns = new Excel.TableColumnCollection(this.context, _createPropertyObjectPath(this.context, this, "Columns", true, false));
                }
                return this.m_columns;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "rows", {
            get: function () {
                if (!this.m_rows) {
                    this.m_rows = new Excel.TableRowCollection(this.context, _createPropertyObjectPath(this.context, this, "Rows", true, false));
                }
                return this.m_rows;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "sort", {
            get: function () {
                if (!this.m_sort) {
                    this.m_sort = new Excel.TableSort(this.context, _createPropertyObjectPath(this.context, this, "Sort", false, false));
                }
                return this.m_sort;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "worksheet", {
            get: function () {
                if (!this.m_worksheet) {
                    this.m_worksheet = new Excel.Worksheet(this.context, _createPropertyObjectPath(this.context, this, "Worksheet", false, false));
                }
                return this.m_worksheet;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "highlightFirstColumn", {
            get: function () {
                _throwIfNotLoaded("highlightFirstColumn", this.m_highlightFirstColumn, "Table", this._isNull);
                return this.m_highlightFirstColumn;
            },
            set: function (value) {
                this.m_highlightFirstColumn = value;
                _createSetPropertyAction(this.context, this, "HighlightFirstColumn", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "highlightLastColumn", {
            get: function () {
                _throwIfNotLoaded("highlightLastColumn", this.m_highlightLastColumn, "Table", this._isNull);
                return this.m_highlightLastColumn;
            },
            set: function (value) {
                this.m_highlightLastColumn = value;
                _createSetPropertyAction(this.context, this, "HighlightLastColumn", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this.m_id, "Table", this._isNull);
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name, "Table", this._isNull);
                return this.m_name;
            },
            set: function (value) {
                this.m_name = value;
                _createSetPropertyAction(this.context, this, "Name", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "showBandedColumns", {
            get: function () {
                _throwIfNotLoaded("showBandedColumns", this.m_showBandedColumns, "Table", this._isNull);
                return this.m_showBandedColumns;
            },
            set: function (value) {
                this.m_showBandedColumns = value;
                _createSetPropertyAction(this.context, this, "ShowBandedColumns", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "showBandedRows", {
            get: function () {
                _throwIfNotLoaded("showBandedRows", this.m_showBandedRows, "Table", this._isNull);
                return this.m_showBandedRows;
            },
            set: function (value) {
                this.m_showBandedRows = value;
                _createSetPropertyAction(this.context, this, "ShowBandedRows", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "showFilterButton", {
            get: function () {
                _throwIfNotLoaded("showFilterButton", this.m_showFilterButton, "Table", this._isNull);
                return this.m_showFilterButton;
            },
            set: function (value) {
                this.m_showFilterButton = value;
                _createSetPropertyAction(this.context, this, "ShowFilterButton", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "showHeaders", {
            get: function () {
                _throwIfNotLoaded("showHeaders", this.m_showHeaders, "Table", this._isNull);
                return this.m_showHeaders;
            },
            set: function (value) {
                this.m_showHeaders = value;
                _createSetPropertyAction(this.context, this, "ShowHeaders", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "showTotals", {
            get: function () {
                _throwIfNotLoaded("showTotals", this.m_showTotals, "Table", this._isNull);
                return this.m_showTotals;
            },
            set: function (value) {
                this.m_showTotals = value;
                _createSetPropertyAction(this.context, this, "ShowTotals", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Table.prototype, "style", {
            get: function () {
                _throwIfNotLoaded("style", this.m_style, "Table", this._isNull);
                return this.m_style;
            },
            set: function (value) {
                this.m_style = value;
                _createSetPropertyAction(this.context, this, "Style", value);
            },
            enumerable: true,
            configurable: true
        });
        Table.prototype.clearFilters = function () {
            _createMethodAction(this.context, this, "ClearFilters", 0 /* Default */, []);
        };
        Table.prototype.convertToRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "ConvertToRange", 0 /* Default */, [], false, true, null));
        };
        Table.prototype.delete = function () {
            _createMethodAction(this.context, this, "Delete", 0 /* Default */, []);
        };
        Table.prototype.getDataBodyRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetDataBodyRange", 1 /* Read */, [], false, true, null));
        };
        Table.prototype.getHeaderRowRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetHeaderRowRange", 1 /* Read */, [], false, true, null));
        };
        Table.prototype.getRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1 /* Read */, [], false, true, null));
        };
        Table.prototype.getTotalRowRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetTotalRowRange", 1 /* Read */, [], false, true, null));
        };
        Table.prototype.reapplyFilters = function () {
            _createMethodAction(this.context, this, "ReapplyFilters", 0 /* Default */, []);
        };
        Table.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["HighlightFirstColumn"])) {
                this.m_highlightFirstColumn = obj["HighlightFirstColumn"];
            }
            if (!_isUndefined(obj["HighlightLastColumn"])) {
                this.m_highlightLastColumn = obj["HighlightLastColumn"];
            }
            if (!_isUndefined(obj["Id"])) {
                this.m_id = obj["Id"];
            }
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            if (!_isUndefined(obj["ShowBandedColumns"])) {
                this.m_showBandedColumns = obj["ShowBandedColumns"];
            }
            if (!_isUndefined(obj["ShowBandedRows"])) {
                this.m_showBandedRows = obj["ShowBandedRows"];
            }
            if (!_isUndefined(obj["ShowFilterButton"])) {
                this.m_showFilterButton = obj["ShowFilterButton"];
            }
            if (!_isUndefined(obj["ShowHeaders"])) {
                this.m_showHeaders = obj["ShowHeaders"];
            }
            if (!_isUndefined(obj["ShowTotals"])) {
                this.m_showTotals = obj["ShowTotals"];
            }
            if (!_isUndefined(obj["Style"])) {
                this.m_style = obj["Style"];
            }
            _handleNavigationPropertyResults(this, obj, ["columns", "Columns", "rows", "Rows", "sort", "Sort", "worksheet", "Worksheet"]);
        };
        Table.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        Table.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["Id"])) {
                this.m_id = value["Id"];
            }
        };
        return Table;
    })(OfficeExtension.ClientObject);
    Excel.Table = Table;
    var TableColumnCollection = (function (_super) {
        __extends(TableColumnCollection, _super);
        function TableColumnCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(TableColumnCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "TableColumnCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableColumnCollection.prototype, "count", {
            get: function () {
                _throwIfNotLoaded("count", this.m_count, "TableColumnCollection", this._isNull);
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });
        TableColumnCollection.prototype.add = function (index, values) {
            return new Excel.TableColumn(this.context, _createMethodObjectPath(this.context, this, "Add", 0 /* Default */, [index, values], false, true, null));
        };
        TableColumnCollection.prototype.getItem = function (key) {
            return new Excel.TableColumn(this.context, _createIndexerObjectPath(this.context, this, [key]));
        };
        TableColumnCollection.prototype.getItemAt = function (index) {
            return new Excel.TableColumn(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [index], false, false, null));
        };
        TableColumnCollection.prototype.getItemOrNull = function (key) {
            return new Excel.TableColumn(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNull", 1 /* Read */, [key], false, false, null));
        };
        TableColumnCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.TableColumn(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        TableColumnCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return TableColumnCollection;
    })(OfficeExtension.ClientObject);
    Excel.TableColumnCollection = TableColumnCollection;
    var TableColumn = (function (_super) {
        __extends(TableColumn, _super);
        function TableColumn() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(TableColumn.prototype, "filter", {
            get: function () {
                if (!this.m_filter) {
                    this.m_filter = new Excel.Filter(this.context, _createPropertyObjectPath(this.context, this, "Filter", false, false));
                }
                return this.m_filter;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableColumn.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this.m_id, "TableColumn", this._isNull);
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableColumn.prototype, "index", {
            get: function () {
                _throwIfNotLoaded("index", this.m_index, "TableColumn", this._isNull);
                return this.m_index;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableColumn.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name, "TableColumn", this._isNull);
                return this.m_name;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableColumn.prototype, "values", {
            get: function () {
                _throwIfNotLoaded("values", this.m_values, "TableColumn", this._isNull);
                return this.m_values;
            },
            set: function (value) {
                this.m_values = value;
                _createSetPropertyAction(this.context, this, "Values", value);
            },
            enumerable: true,
            configurable: true
        });
        TableColumn.prototype.delete = function () {
            _createMethodAction(this.context, this, "Delete", 0 /* Default */, []);
        };
        TableColumn.prototype.getDataBodyRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetDataBodyRange", 1 /* Read */, [], false, true, null));
        };
        TableColumn.prototype.getHeaderRowRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetHeaderRowRange", 1 /* Read */, [], false, true, null));
        };
        TableColumn.prototype.getRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1 /* Read */, [], false, true, null));
        };
        TableColumn.prototype.getTotalRowRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetTotalRowRange", 1 /* Read */, [], false, true, null));
        };
        TableColumn.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Id"])) {
                this.m_id = obj["Id"];
            }
            if (!_isUndefined(obj["Index"])) {
                this.m_index = obj["Index"];
            }
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            if (!_isUndefined(obj["Values"])) {
                this.m_values = obj["Values"];
            }
            _handleNavigationPropertyResults(this, obj, ["filter", "Filter"]);
        };
        TableColumn.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        TableColumn.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["Id"])) {
                this.m_id = value["Id"];
            }
        };
        return TableColumn;
    })(OfficeExtension.ClientObject);
    Excel.TableColumn = TableColumn;
    var TableRowCollection = (function (_super) {
        __extends(TableRowCollection, _super);
        function TableRowCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(TableRowCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "TableRowCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableRowCollection.prototype, "count", {
            get: function () {
                _throwIfNotLoaded("count", this.m_count, "TableRowCollection", this._isNull);
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });
        TableRowCollection.prototype.add = function (index, values) {
            return new Excel.TableRow(this.context, _createMethodObjectPath(this.context, this, "Add", 0 /* Default */, [index, values], false, true, null));
        };
        TableRowCollection.prototype.getItemAt = function (index) {
            return new Excel.TableRow(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [index], false, false, null));
        };
        TableRowCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.TableRow(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(false, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        TableRowCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return TableRowCollection;
    })(OfficeExtension.ClientObject);
    Excel.TableRowCollection = TableRowCollection;
    var TableRow = (function (_super) {
        __extends(TableRow, _super);
        function TableRow() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(TableRow.prototype, "index", {
            get: function () {
                _throwIfNotLoaded("index", this.m_index, "TableRow", this._isNull);
                return this.m_index;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableRow.prototype, "values", {
            get: function () {
                _throwIfNotLoaded("values", this.m_values, "TableRow", this._isNull);
                return this.m_values;
            },
            set: function (value) {
                this.m_values = value;
                _createSetPropertyAction(this.context, this, "Values", value);
            },
            enumerable: true,
            configurable: true
        });
        TableRow.prototype.delete = function () {
            _createMethodAction(this.context, this, "Delete", 0 /* Default */, []);
        };
        TableRow.prototype.getRange = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1 /* Read */, [], false, true, null));
        };
        TableRow.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Index"])) {
                this.m_index = obj["Index"];
            }
            if (!_isUndefined(obj["Values"])) {
                this.m_values = obj["Values"];
            }
        };
        TableRow.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return TableRow;
    })(OfficeExtension.ClientObject);
    Excel.TableRow = TableRow;
    var RangeFormat = (function (_super) {
        __extends(RangeFormat, _super);
        function RangeFormat() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(RangeFormat.prototype, "borders", {
            get: function () {
                if (!this.m_borders) {
                    this.m_borders = new Excel.RangeBorderCollection(this.context, _createPropertyObjectPath(this.context, this, "Borders", true, false));
                }
                return this.m_borders;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFormat.prototype, "fill", {
            get: function () {
                if (!this.m_fill) {
                    this.m_fill = new Excel.RangeFill(this.context, _createPropertyObjectPath(this.context, this, "Fill", false, false));
                }
                return this.m_fill;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFormat.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Excel.RangeFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFormat.prototype, "protection", {
            get: function () {
                if (!this.m_protection) {
                    this.m_protection = new Excel.FormatProtection(this.context, _createPropertyObjectPath(this.context, this, "Protection", false, false));
                }
                return this.m_protection;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFormat.prototype, "columnWidth", {
            get: function () {
                _throwIfNotLoaded("columnWidth", this.m_columnWidth, "RangeFormat", this._isNull);
                return this.m_columnWidth;
            },
            set: function (value) {
                this.m_columnWidth = value;
                _createSetPropertyAction(this.context, this, "ColumnWidth", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFormat.prototype, "horizontalAlignment", {
            get: function () {
                _throwIfNotLoaded("horizontalAlignment", this.m_horizontalAlignment, "RangeFormat", this._isNull);
                return this.m_horizontalAlignment;
            },
            set: function (value) {
                this.m_horizontalAlignment = value;
                _createSetPropertyAction(this.context, this, "HorizontalAlignment", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFormat.prototype, "rowHeight", {
            get: function () {
                _throwIfNotLoaded("rowHeight", this.m_rowHeight, "RangeFormat", this._isNull);
                return this.m_rowHeight;
            },
            set: function (value) {
                this.m_rowHeight = value;
                _createSetPropertyAction(this.context, this, "RowHeight", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFormat.prototype, "verticalAlignment", {
            get: function () {
                _throwIfNotLoaded("verticalAlignment", this.m_verticalAlignment, "RangeFormat", this._isNull);
                return this.m_verticalAlignment;
            },
            set: function (value) {
                this.m_verticalAlignment = value;
                _createSetPropertyAction(this.context, this, "VerticalAlignment", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFormat.prototype, "wrapText", {
            get: function () {
                _throwIfNotLoaded("wrapText", this.m_wrapText, "RangeFormat", this._isNull);
                return this.m_wrapText;
            },
            set: function (value) {
                this.m_wrapText = value;
                _createSetPropertyAction(this.context, this, "WrapText", value);
            },
            enumerable: true,
            configurable: true
        });
        RangeFormat.prototype.autofitColumns = function () {
            _createMethodAction(this.context, this, "AutofitColumns", 0 /* Default */, []);
        };
        RangeFormat.prototype.autofitRows = function () {
            _createMethodAction(this.context, this, "AutofitRows", 0 /* Default */, []);
        };
        RangeFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["ColumnWidth"])) {
                this.m_columnWidth = obj["ColumnWidth"];
            }
            if (!_isUndefined(obj["HorizontalAlignment"])) {
                this.m_horizontalAlignment = obj["HorizontalAlignment"];
            }
            if (!_isUndefined(obj["RowHeight"])) {
                this.m_rowHeight = obj["RowHeight"];
            }
            if (!_isUndefined(obj["VerticalAlignment"])) {
                this.m_verticalAlignment = obj["VerticalAlignment"];
            }
            if (!_isUndefined(obj["WrapText"])) {
                this.m_wrapText = obj["WrapText"];
            }
            _handleNavigationPropertyResults(this, obj, ["borders", "Borders", "fill", "Fill", "font", "Font", "protection", "Protection"]);
        };
        RangeFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return RangeFormat;
    })(OfficeExtension.ClientObject);
    Excel.RangeFormat = RangeFormat;
    var FormatProtection = (function (_super) {
        __extends(FormatProtection, _super);
        function FormatProtection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(FormatProtection.prototype, "formulaHidden", {
            get: function () {
                _throwIfNotLoaded("formulaHidden", this.m_formulaHidden, "FormatProtection", this._isNull);
                return this.m_formulaHidden;
            },
            set: function (value) {
                this.m_formulaHidden = value;
                _createSetPropertyAction(this.context, this, "FormulaHidden", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(FormatProtection.prototype, "locked", {
            get: function () {
                _throwIfNotLoaded("locked", this.m_locked, "FormatProtection", this._isNull);
                return this.m_locked;
            },
            set: function (value) {
                this.m_locked = value;
                _createSetPropertyAction(this.context, this, "Locked", value);
            },
            enumerable: true,
            configurable: true
        });
        FormatProtection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["FormulaHidden"])) {
                this.m_formulaHidden = obj["FormulaHidden"];
            }
            if (!_isUndefined(obj["Locked"])) {
                this.m_locked = obj["Locked"];
            }
        };
        FormatProtection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return FormatProtection;
    })(OfficeExtension.ClientObject);
    Excel.FormatProtection = FormatProtection;
    var RangeFill = (function (_super) {
        __extends(RangeFill, _super);
        function RangeFill() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(RangeFill.prototype, "color", {
            get: function () {
                _throwIfNotLoaded("color", this.m_color, "RangeFill", this._isNull);
                return this.m_color;
            },
            set: function (value) {
                this.m_color = value;
                _createSetPropertyAction(this.context, this, "Color", value);
            },
            enumerable: true,
            configurable: true
        });
        RangeFill.prototype.clear = function () {
            _createMethodAction(this.context, this, "Clear", 0 /* Default */, []);
        };
        RangeFill.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Color"])) {
                this.m_color = obj["Color"];
            }
        };
        RangeFill.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return RangeFill;
    })(OfficeExtension.ClientObject);
    Excel.RangeFill = RangeFill;
    var RangeBorder = (function (_super) {
        __extends(RangeBorder, _super);
        function RangeBorder() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(RangeBorder.prototype, "color", {
            get: function () {
                _throwIfNotLoaded("color", this.m_color, "RangeBorder", this._isNull);
                return this.m_color;
            },
            set: function (value) {
                this.m_color = value;
                _createSetPropertyAction(this.context, this, "Color", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeBorder.prototype, "sideIndex", {
            get: function () {
                _throwIfNotLoaded("sideIndex", this.m_sideIndex, "RangeBorder", this._isNull);
                return this.m_sideIndex;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeBorder.prototype, "style", {
            get: function () {
                _throwIfNotLoaded("style", this.m_style, "RangeBorder", this._isNull);
                return this.m_style;
            },
            set: function (value) {
                this.m_style = value;
                _createSetPropertyAction(this.context, this, "Style", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeBorder.prototype, "weight", {
            get: function () {
                _throwIfNotLoaded("weight", this.m_weight, "RangeBorder", this._isNull);
                return this.m_weight;
            },
            set: function (value) {
                this.m_weight = value;
                _createSetPropertyAction(this.context, this, "Weight", value);
            },
            enumerable: true,
            configurable: true
        });
        RangeBorder.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Color"])) {
                this.m_color = obj["Color"];
            }
            if (!_isUndefined(obj["SideIndex"])) {
                this.m_sideIndex = obj["SideIndex"];
            }
            if (!_isUndefined(obj["Style"])) {
                this.m_style = obj["Style"];
            }
            if (!_isUndefined(obj["Weight"])) {
                this.m_weight = obj["Weight"];
            }
        };
        RangeBorder.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return RangeBorder;
    })(OfficeExtension.ClientObject);
    Excel.RangeBorder = RangeBorder;
    var RangeBorderCollection = (function (_super) {
        __extends(RangeBorderCollection, _super);
        function RangeBorderCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(RangeBorderCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "RangeBorderCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeBorderCollection.prototype, "count", {
            get: function () {
                _throwIfNotLoaded("count", this.m_count, "RangeBorderCollection", this._isNull);
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });
        RangeBorderCollection.prototype.getItem = function (index) {
            return new Excel.RangeBorder(this.context, _createIndexerObjectPath(this.context, this, [index]));
        };
        RangeBorderCollection.prototype.getItemAt = function (index) {
            return new Excel.RangeBorder(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [index], false, false, null));
        };
        RangeBorderCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.RangeBorder(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        RangeBorderCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return RangeBorderCollection;
    })(OfficeExtension.ClientObject);
    Excel.RangeBorderCollection = RangeBorderCollection;
    var RangeFont = (function (_super) {
        __extends(RangeFont, _super);
        function RangeFont() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(RangeFont.prototype, "bold", {
            get: function () {
                _throwIfNotLoaded("bold", this.m_bold, "RangeFont", this._isNull);
                return this.m_bold;
            },
            set: function (value) {
                this.m_bold = value;
                _createSetPropertyAction(this.context, this, "Bold", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFont.prototype, "color", {
            get: function () {
                _throwIfNotLoaded("color", this.m_color, "RangeFont", this._isNull);
                return this.m_color;
            },
            set: function (value) {
                this.m_color = value;
                _createSetPropertyAction(this.context, this, "Color", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFont.prototype, "italic", {
            get: function () {
                _throwIfNotLoaded("italic", this.m_italic, "RangeFont", this._isNull);
                return this.m_italic;
            },
            set: function (value) {
                this.m_italic = value;
                _createSetPropertyAction(this.context, this, "Italic", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFont.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name, "RangeFont", this._isNull);
                return this.m_name;
            },
            set: function (value) {
                this.m_name = value;
                _createSetPropertyAction(this.context, this, "Name", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFont.prototype, "size", {
            get: function () {
                _throwIfNotLoaded("size", this.m_size, "RangeFont", this._isNull);
                return this.m_size;
            },
            set: function (value) {
                this.m_size = value;
                _createSetPropertyAction(this.context, this, "Size", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeFont.prototype, "underline", {
            get: function () {
                _throwIfNotLoaded("underline", this.m_underline, "RangeFont", this._isNull);
                return this.m_underline;
            },
            set: function (value) {
                this.m_underline = value;
                _createSetPropertyAction(this.context, this, "Underline", value);
            },
            enumerable: true,
            configurable: true
        });
        RangeFont.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Bold"])) {
                this.m_bold = obj["Bold"];
            }
            if (!_isUndefined(obj["Color"])) {
                this.m_color = obj["Color"];
            }
            if (!_isUndefined(obj["Italic"])) {
                this.m_italic = obj["Italic"];
            }
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            if (!_isUndefined(obj["Size"])) {
                this.m_size = obj["Size"];
            }
            if (!_isUndefined(obj["Underline"])) {
                this.m_underline = obj["Underline"];
            }
        };
        RangeFont.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return RangeFont;
    })(OfficeExtension.ClientObject);
    Excel.RangeFont = RangeFont;
    var ChartCollection = (function (_super) {
        __extends(ChartCollection, _super);
        function ChartCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "ChartCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartCollection.prototype, "count", {
            get: function () {
                _throwIfNotLoaded("count", this.m_count, "ChartCollection", this._isNull);
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });
        ChartCollection.prototype.add = function (type, sourceData, seriesBy) {
            if (!(sourceData instanceof Range)) {
                throw OfficeExtension.Utility.createRuntimeError(OfficeExtension.ResourceStrings.invalidArgument, "sourceData", "Charts.Add");
            }
            return new Excel.Chart(this.context, _createMethodObjectPath(this.context, this, "Add", 0 /* Default */, [type, sourceData, seriesBy], false, true, null));
        };
        ChartCollection.prototype.getItem = function (name) {
            return new Excel.Chart(this.context, _createMethodObjectPath(this.context, this, "GetItem", 1 /* Read */, [name], false, false, null));
        };
        ChartCollection.prototype.getItemAt = function (index) {
            return new Excel.Chart(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [index], false, false, null));
        };
        ChartCollection.prototype.getItemOrNull = function (name) {
            return new Excel.Chart(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNull", 1 /* Read */, [name], false, false, null));
        };
        ChartCollection.prototype._GetItem = function (key) {
            return new Excel.Chart(this.context, _createIndexerObjectPath(this.context, this, [key]));
        };
        ChartCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.Chart(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        ChartCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartCollection;
    })(OfficeExtension.ClientObject);
    Excel.ChartCollection = ChartCollection;
    var Chart = (function (_super) {
        __extends(Chart, _super);
        function Chart() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Chart.prototype, "axes", {
            get: function () {
                if (!this.m_axes) {
                    this.m_axes = new Excel.ChartAxes(this.context, _createPropertyObjectPath(this.context, this, "Axes", false, false));
                }
                return this.m_axes;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "dataLabels", {
            get: function () {
                if (!this.m_dataLabels) {
                    this.m_dataLabels = new Excel.ChartDataLabels(this.context, _createPropertyObjectPath(this.context, this, "DataLabels", false, false));
                }
                return this.m_dataLabels;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ChartAreaFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "legend", {
            get: function () {
                if (!this.m_legend) {
                    this.m_legend = new Excel.ChartLegend(this.context, _createPropertyObjectPath(this.context, this, "Legend", false, false));
                }
                return this.m_legend;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "series", {
            get: function () {
                if (!this.m_series) {
                    this.m_series = new Excel.ChartSeriesCollection(this.context, _createPropertyObjectPath(this.context, this, "Series", true, false));
                }
                return this.m_series;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "title", {
            get: function () {
                if (!this.m_title) {
                    this.m_title = new Excel.ChartTitle(this.context, _createPropertyObjectPath(this.context, this, "Title", false, false));
                }
                return this.m_title;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "worksheet", {
            get: function () {
                if (!this.m_worksheet) {
                    this.m_worksheet = new Excel.Worksheet(this.context, _createPropertyObjectPath(this.context, this, "Worksheet", false, false));
                }
                return this.m_worksheet;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "height", {
            get: function () {
                _throwIfNotLoaded("height", this.m_height, "Chart", this._isNull);
                return this.m_height;
            },
            set: function (value) {
                this.m_height = value;
                _createSetPropertyAction(this.context, this, "Height", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "left", {
            get: function () {
                _throwIfNotLoaded("left", this.m_left, "Chart", this._isNull);
                return this.m_left;
            },
            set: function (value) {
                this.m_left = value;
                _createSetPropertyAction(this.context, this, "Left", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name, "Chart", this._isNull);
                return this.m_name;
            },
            set: function (value) {
                this.m_name = value;
                _createSetPropertyAction(this.context, this, "Name", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "top", {
            get: function () {
                _throwIfNotLoaded("top", this.m_top, "Chart", this._isNull);
                return this.m_top;
            },
            set: function (value) {
                this.m_top = value;
                _createSetPropertyAction(this.context, this, "Top", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "width", {
            get: function () {
                _throwIfNotLoaded("width", this.m_width, "Chart", this._isNull);
                return this.m_width;
            },
            set: function (value) {
                this.m_width = value;
                _createSetPropertyAction(this.context, this, "Width", value);
            },
            enumerable: true,
            configurable: true
        });
        Chart.prototype.delete = function () {
            _createMethodAction(this.context, this, "Delete", 0 /* Default */, []);
        };
        Chart.prototype.getImage = function (width, height, fittingMode) {
            var action = _createMethodAction(this.context, this, "GetImage", 1 /* Read */, [width, height, fittingMode]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        Chart.prototype.setData = function (sourceData, seriesBy) {
            if (!(sourceData instanceof Range)) {
                throw OfficeExtension.Utility.createRuntimeError(OfficeExtension.ResourceStrings.invalidArgument, "sourceData", "Chart.setData");
            }
            _createMethodAction(this.context, this, "SetData", 0 /* Default */, [sourceData, seriesBy]);
        };
        Chart.prototype.setPosition = function (startCell, endCell) {
            _createMethodAction(this.context, this, "SetPosition", 0 /* Default */, [startCell, endCell]);
        };
        Chart.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Height"])) {
                this.m_height = obj["Height"];
            }
            if (!_isUndefined(obj["Left"])) {
                this.m_left = obj["Left"];
            }
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            if (!_isUndefined(obj["Top"])) {
                this.m_top = obj["Top"];
            }
            if (!_isUndefined(obj["Width"])) {
                this.m_width = obj["Width"];
            }
            _handleNavigationPropertyResults(this, obj, ["axes", "Axes", "dataLabels", "DataLabels", "format", "Format", "legend", "Legend", "series", "Series", "title", "Title", "worksheet", "Worksheet"]);
        };
        Chart.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return Chart;
    })(OfficeExtension.ClientObject);
    Excel.Chart = Chart;
    var ChartAreaFormat = (function (_super) {
        __extends(ChartAreaFormat, _super);
        function ChartAreaFormat() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartAreaFormat.prototype, "fill", {
            get: function () {
                if (!this.m_fill) {
                    this.m_fill = new Excel.ChartFill(this.context, _createPropertyObjectPath(this.context, this, "Fill", false, false));
                }
                return this.m_fill;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAreaFormat.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Excel.ChartFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });
        ChartAreaFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["fill", "Fill", "font", "Font"]);
        };
        ChartAreaFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartAreaFormat;
    })(OfficeExtension.ClientObject);
    Excel.ChartAreaFormat = ChartAreaFormat;
    var ChartSeriesCollection = (function (_super) {
        __extends(ChartSeriesCollection, _super);
        function ChartSeriesCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartSeriesCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "ChartSeriesCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartSeriesCollection.prototype, "count", {
            get: function () {
                _throwIfNotLoaded("count", this.m_count, "ChartSeriesCollection", this._isNull);
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });
        ChartSeriesCollection.prototype.getItemAt = function (index) {
            return new Excel.ChartSeries(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [index], false, false, null));
        };
        ChartSeriesCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.ChartSeries(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(false, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        ChartSeriesCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartSeriesCollection;
    })(OfficeExtension.ClientObject);
    Excel.ChartSeriesCollection = ChartSeriesCollection;
    var ChartSeries = (function (_super) {
        __extends(ChartSeries, _super);
        function ChartSeries() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartSeries.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ChartSeriesFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartSeries.prototype, "points", {
            get: function () {
                if (!this.m_points) {
                    this.m_points = new Excel.ChartPointsCollection(this.context, _createPropertyObjectPath(this.context, this, "Points", true, false));
                }
                return this.m_points;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartSeries.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name, "ChartSeries", this._isNull);
                return this.m_name;
            },
            set: function (value) {
                this.m_name = value;
                _createSetPropertyAction(this.context, this, "Name", value);
            },
            enumerable: true,
            configurable: true
        });
        ChartSeries.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            _handleNavigationPropertyResults(this, obj, ["format", "Format", "points", "Points"]);
        };
        ChartSeries.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartSeries;
    })(OfficeExtension.ClientObject);
    Excel.ChartSeries = ChartSeries;
    var ChartSeriesFormat = (function (_super) {
        __extends(ChartSeriesFormat, _super);
        function ChartSeriesFormat() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartSeriesFormat.prototype, "fill", {
            get: function () {
                if (!this.m_fill) {
                    this.m_fill = new Excel.ChartFill(this.context, _createPropertyObjectPath(this.context, this, "Fill", false, false));
                }
                return this.m_fill;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartSeriesFormat.prototype, "line", {
            get: function () {
                if (!this.m_line) {
                    this.m_line = new Excel.ChartLineFormat(this.context, _createPropertyObjectPath(this.context, this, "Line", false, false));
                }
                return this.m_line;
            },
            enumerable: true,
            configurable: true
        });
        ChartSeriesFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["fill", "Fill", "line", "Line"]);
        };
        ChartSeriesFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartSeriesFormat;
    })(OfficeExtension.ClientObject);
    Excel.ChartSeriesFormat = ChartSeriesFormat;
    var ChartPointsCollection = (function (_super) {
        __extends(ChartPointsCollection, _super);
        function ChartPointsCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartPointsCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "ChartPointsCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartPointsCollection.prototype, "count", {
            get: function () {
                _throwIfNotLoaded("count", this.m_count, "ChartPointsCollection", this._isNull);
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });
        ChartPointsCollection.prototype.getItemAt = function (index) {
            return new Excel.ChartPoint(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [index], false, false, null));
        };
        ChartPointsCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.ChartPoint(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(false, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        ChartPointsCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartPointsCollection;
    })(OfficeExtension.ClientObject);
    Excel.ChartPointsCollection = ChartPointsCollection;
    var ChartPoint = (function (_super) {
        __extends(ChartPoint, _super);
        function ChartPoint() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartPoint.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ChartPointFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartPoint.prototype, "value", {
            get: function () {
                _throwIfNotLoaded("value", this.m_value, "ChartPoint", this._isNull);
                return this.m_value;
            },
            enumerable: true,
            configurable: true
        });
        ChartPoint.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Value"])) {
                this.m_value = obj["Value"];
            }
            _handleNavigationPropertyResults(this, obj, ["format", "Format"]);
        };
        ChartPoint.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartPoint;
    })(OfficeExtension.ClientObject);
    Excel.ChartPoint = ChartPoint;
    var ChartPointFormat = (function (_super) {
        __extends(ChartPointFormat, _super);
        function ChartPointFormat() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartPointFormat.prototype, "fill", {
            get: function () {
                if (!this.m_fill) {
                    this.m_fill = new Excel.ChartFill(this.context, _createPropertyObjectPath(this.context, this, "Fill", false, false));
                }
                return this.m_fill;
            },
            enumerable: true,
            configurable: true
        });
        ChartPointFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["fill", "Fill"]);
        };
        ChartPointFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartPointFormat;
    })(OfficeExtension.ClientObject);
    Excel.ChartPointFormat = ChartPointFormat;
    var ChartAxes = (function (_super) {
        __extends(ChartAxes, _super);
        function ChartAxes() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartAxes.prototype, "categoryAxis", {
            get: function () {
                if (!this.m_categoryAxis) {
                    this.m_categoryAxis = new Excel.ChartAxis(this.context, _createPropertyObjectPath(this.context, this, "CategoryAxis", false, false));
                }
                return this.m_categoryAxis;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxes.prototype, "seriesAxis", {
            get: function () {
                if (!this.m_seriesAxis) {
                    this.m_seriesAxis = new Excel.ChartAxis(this.context, _createPropertyObjectPath(this.context, this, "SeriesAxis", false, false));
                }
                return this.m_seriesAxis;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxes.prototype, "valueAxis", {
            get: function () {
                if (!this.m_valueAxis) {
                    this.m_valueAxis = new Excel.ChartAxis(this.context, _createPropertyObjectPath(this.context, this, "ValueAxis", false, false));
                }
                return this.m_valueAxis;
            },
            enumerable: true,
            configurable: true
        });
        ChartAxes.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["categoryAxis", "CategoryAxis", "seriesAxis", "SeriesAxis", "valueAxis", "ValueAxis"]);
        };
        ChartAxes.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartAxes;
    })(OfficeExtension.ClientObject);
    Excel.ChartAxes = ChartAxes;
    var ChartAxis = (function (_super) {
        __extends(ChartAxis, _super);
        function ChartAxis() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartAxis.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ChartAxisFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxis.prototype, "majorGridlines", {
            get: function () {
                if (!this.m_majorGridlines) {
                    this.m_majorGridlines = new Excel.ChartGridlines(this.context, _createPropertyObjectPath(this.context, this, "MajorGridlines", false, false));
                }
                return this.m_majorGridlines;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxis.prototype, "minorGridlines", {
            get: function () {
                if (!this.m_minorGridlines) {
                    this.m_minorGridlines = new Excel.ChartGridlines(this.context, _createPropertyObjectPath(this.context, this, "MinorGridlines", false, false));
                }
                return this.m_minorGridlines;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxis.prototype, "title", {
            get: function () {
                if (!this.m_title) {
                    this.m_title = new Excel.ChartAxisTitle(this.context, _createPropertyObjectPath(this.context, this, "Title", false, false));
                }
                return this.m_title;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxis.prototype, "majorUnit", {
            get: function () {
                _throwIfNotLoaded("majorUnit", this.m_majorUnit, "ChartAxis", this._isNull);
                return this.m_majorUnit;
            },
            set: function (value) {
                this.m_majorUnit = value;
                _createSetPropertyAction(this.context, this, "MajorUnit", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxis.prototype, "maximum", {
            get: function () {
                _throwIfNotLoaded("maximum", this.m_maximum, "ChartAxis", this._isNull);
                return this.m_maximum;
            },
            set: function (value) {
                this.m_maximum = value;
                _createSetPropertyAction(this.context, this, "Maximum", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxis.prototype, "minimum", {
            get: function () {
                _throwIfNotLoaded("minimum", this.m_minimum, "ChartAxis", this._isNull);
                return this.m_minimum;
            },
            set: function (value) {
                this.m_minimum = value;
                _createSetPropertyAction(this.context, this, "Minimum", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxis.prototype, "minorUnit", {
            get: function () {
                _throwIfNotLoaded("minorUnit", this.m_minorUnit, "ChartAxis", this._isNull);
                return this.m_minorUnit;
            },
            set: function (value) {
                this.m_minorUnit = value;
                _createSetPropertyAction(this.context, this, "MinorUnit", value);
            },
            enumerable: true,
            configurable: true
        });
        ChartAxis.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["MajorUnit"])) {
                this.m_majorUnit = obj["MajorUnit"];
            }
            if (!_isUndefined(obj["Maximum"])) {
                this.m_maximum = obj["Maximum"];
            }
            if (!_isUndefined(obj["Minimum"])) {
                this.m_minimum = obj["Minimum"];
            }
            if (!_isUndefined(obj["MinorUnit"])) {
                this.m_minorUnit = obj["MinorUnit"];
            }
            _handleNavigationPropertyResults(this, obj, ["format", "Format", "majorGridlines", "MajorGridlines", "minorGridlines", "MinorGridlines", "title", "Title"]);
        };
        ChartAxis.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartAxis;
    })(OfficeExtension.ClientObject);
    Excel.ChartAxis = ChartAxis;
    var ChartAxisFormat = (function (_super) {
        __extends(ChartAxisFormat, _super);
        function ChartAxisFormat() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartAxisFormat.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Excel.ChartFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxisFormat.prototype, "line", {
            get: function () {
                if (!this.m_line) {
                    this.m_line = new Excel.ChartLineFormat(this.context, _createPropertyObjectPath(this.context, this, "Line", false, false));
                }
                return this.m_line;
            },
            enumerable: true,
            configurable: true
        });
        ChartAxisFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["font", "Font", "line", "Line"]);
        };
        ChartAxisFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartAxisFormat;
    })(OfficeExtension.ClientObject);
    Excel.ChartAxisFormat = ChartAxisFormat;
    var ChartAxisTitle = (function (_super) {
        __extends(ChartAxisTitle, _super);
        function ChartAxisTitle() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartAxisTitle.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ChartAxisTitleFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxisTitle.prototype, "text", {
            get: function () {
                _throwIfNotLoaded("text", this.m_text, "ChartAxisTitle", this._isNull);
                return this.m_text;
            },
            set: function (value) {
                this.m_text = value;
                _createSetPropertyAction(this.context, this, "Text", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartAxisTitle.prototype, "visible", {
            get: function () {
                _throwIfNotLoaded("visible", this.m_visible, "ChartAxisTitle", this._isNull);
                return this.m_visible;
            },
            set: function (value) {
                this.m_visible = value;
                _createSetPropertyAction(this.context, this, "Visible", value);
            },
            enumerable: true,
            configurable: true
        });
        ChartAxisTitle.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Text"])) {
                this.m_text = obj["Text"];
            }
            if (!_isUndefined(obj["Visible"])) {
                this.m_visible = obj["Visible"];
            }
            _handleNavigationPropertyResults(this, obj, ["format", "Format"]);
        };
        ChartAxisTitle.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartAxisTitle;
    })(OfficeExtension.ClientObject);
    Excel.ChartAxisTitle = ChartAxisTitle;
    var ChartAxisTitleFormat = (function (_super) {
        __extends(ChartAxisTitleFormat, _super);
        function ChartAxisTitleFormat() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartAxisTitleFormat.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Excel.ChartFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });
        ChartAxisTitleFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["font", "Font"]);
        };
        ChartAxisTitleFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartAxisTitleFormat;
    })(OfficeExtension.ClientObject);
    Excel.ChartAxisTitleFormat = ChartAxisTitleFormat;
    var ChartDataLabels = (function (_super) {
        __extends(ChartDataLabels, _super);
        function ChartDataLabels() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartDataLabels.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ChartDataLabelFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartDataLabels.prototype, "position", {
            get: function () {
                _throwIfNotLoaded("position", this.m_position, "ChartDataLabels", this._isNull);
                return this.m_position;
            },
            set: function (value) {
                this.m_position = value;
                _createSetPropertyAction(this.context, this, "Position", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartDataLabels.prototype, "separator", {
            get: function () {
                _throwIfNotLoaded("separator", this.m_separator, "ChartDataLabels", this._isNull);
                return this.m_separator;
            },
            set: function (value) {
                this.m_separator = value;
                _createSetPropertyAction(this.context, this, "Separator", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartDataLabels.prototype, "showBubbleSize", {
            get: function () {
                _throwIfNotLoaded("showBubbleSize", this.m_showBubbleSize, "ChartDataLabels", this._isNull);
                return this.m_showBubbleSize;
            },
            set: function (value) {
                this.m_showBubbleSize = value;
                _createSetPropertyAction(this.context, this, "ShowBubbleSize", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartDataLabels.prototype, "showCategoryName", {
            get: function () {
                _throwIfNotLoaded("showCategoryName", this.m_showCategoryName, "ChartDataLabels", this._isNull);
                return this.m_showCategoryName;
            },
            set: function (value) {
                this.m_showCategoryName = value;
                _createSetPropertyAction(this.context, this, "ShowCategoryName", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartDataLabels.prototype, "showLegendKey", {
            get: function () {
                _throwIfNotLoaded("showLegendKey", this.m_showLegendKey, "ChartDataLabels", this._isNull);
                return this.m_showLegendKey;
            },
            set: function (value) {
                this.m_showLegendKey = value;
                _createSetPropertyAction(this.context, this, "ShowLegendKey", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartDataLabels.prototype, "showPercentage", {
            get: function () {
                _throwIfNotLoaded("showPercentage", this.m_showPercentage, "ChartDataLabels", this._isNull);
                return this.m_showPercentage;
            },
            set: function (value) {
                this.m_showPercentage = value;
                _createSetPropertyAction(this.context, this, "ShowPercentage", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartDataLabels.prototype, "showSeriesName", {
            get: function () {
                _throwIfNotLoaded("showSeriesName", this.m_showSeriesName, "ChartDataLabels", this._isNull);
                return this.m_showSeriesName;
            },
            set: function (value) {
                this.m_showSeriesName = value;
                _createSetPropertyAction(this.context, this, "ShowSeriesName", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartDataLabels.prototype, "showValue", {
            get: function () {
                _throwIfNotLoaded("showValue", this.m_showValue, "ChartDataLabels", this._isNull);
                return this.m_showValue;
            },
            set: function (value) {
                this.m_showValue = value;
                _createSetPropertyAction(this.context, this, "ShowValue", value);
            },
            enumerable: true,
            configurable: true
        });
        ChartDataLabels.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Position"])) {
                this.m_position = obj["Position"];
            }
            if (!_isUndefined(obj["Separator"])) {
                this.m_separator = obj["Separator"];
            }
            if (!_isUndefined(obj["ShowBubbleSize"])) {
                this.m_showBubbleSize = obj["ShowBubbleSize"];
            }
            if (!_isUndefined(obj["ShowCategoryName"])) {
                this.m_showCategoryName = obj["ShowCategoryName"];
            }
            if (!_isUndefined(obj["ShowLegendKey"])) {
                this.m_showLegendKey = obj["ShowLegendKey"];
            }
            if (!_isUndefined(obj["ShowPercentage"])) {
                this.m_showPercentage = obj["ShowPercentage"];
            }
            if (!_isUndefined(obj["ShowSeriesName"])) {
                this.m_showSeriesName = obj["ShowSeriesName"];
            }
            if (!_isUndefined(obj["ShowValue"])) {
                this.m_showValue = obj["ShowValue"];
            }
            _handleNavigationPropertyResults(this, obj, ["format", "Format"]);
        };
        ChartDataLabels.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartDataLabels;
    })(OfficeExtension.ClientObject);
    Excel.ChartDataLabels = ChartDataLabels;
    var ChartDataLabelFormat = (function (_super) {
        __extends(ChartDataLabelFormat, _super);
        function ChartDataLabelFormat() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartDataLabelFormat.prototype, "fill", {
            get: function () {
                if (!this.m_fill) {
                    this.m_fill = new Excel.ChartFill(this.context, _createPropertyObjectPath(this.context, this, "Fill", false, false));
                }
                return this.m_fill;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartDataLabelFormat.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Excel.ChartFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });
        ChartDataLabelFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["fill", "Fill", "font", "Font"]);
        };
        ChartDataLabelFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartDataLabelFormat;
    })(OfficeExtension.ClientObject);
    Excel.ChartDataLabelFormat = ChartDataLabelFormat;
    var ChartGridlines = (function (_super) {
        __extends(ChartGridlines, _super);
        function ChartGridlines() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartGridlines.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ChartGridlinesFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartGridlines.prototype, "visible", {
            get: function () {
                _throwIfNotLoaded("visible", this.m_visible, "ChartGridlines", this._isNull);
                return this.m_visible;
            },
            set: function (value) {
                this.m_visible = value;
                _createSetPropertyAction(this.context, this, "Visible", value);
            },
            enumerable: true,
            configurable: true
        });
        ChartGridlines.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Visible"])) {
                this.m_visible = obj["Visible"];
            }
            _handleNavigationPropertyResults(this, obj, ["format", "Format"]);
        };
        ChartGridlines.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartGridlines;
    })(OfficeExtension.ClientObject);
    Excel.ChartGridlines = ChartGridlines;
    var ChartGridlinesFormat = (function (_super) {
        __extends(ChartGridlinesFormat, _super);
        function ChartGridlinesFormat() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartGridlinesFormat.prototype, "line", {
            get: function () {
                if (!this.m_line) {
                    this.m_line = new Excel.ChartLineFormat(this.context, _createPropertyObjectPath(this.context, this, "Line", false, false));
                }
                return this.m_line;
            },
            enumerable: true,
            configurable: true
        });
        ChartGridlinesFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["line", "Line"]);
        };
        ChartGridlinesFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartGridlinesFormat;
    })(OfficeExtension.ClientObject);
    Excel.ChartGridlinesFormat = ChartGridlinesFormat;
    var ChartLegend = (function (_super) {
        __extends(ChartLegend, _super);
        function ChartLegend() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartLegend.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ChartLegendFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartLegend.prototype, "overlay", {
            get: function () {
                _throwIfNotLoaded("overlay", this.m_overlay, "ChartLegend", this._isNull);
                return this.m_overlay;
            },
            set: function (value) {
                this.m_overlay = value;
                _createSetPropertyAction(this.context, this, "Overlay", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartLegend.prototype, "position", {
            get: function () {
                _throwIfNotLoaded("position", this.m_position, "ChartLegend", this._isNull);
                return this.m_position;
            },
            set: function (value) {
                this.m_position = value;
                _createSetPropertyAction(this.context, this, "Position", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartLegend.prototype, "visible", {
            get: function () {
                _throwIfNotLoaded("visible", this.m_visible, "ChartLegend", this._isNull);
                return this.m_visible;
            },
            set: function (value) {
                this.m_visible = value;
                _createSetPropertyAction(this.context, this, "Visible", value);
            },
            enumerable: true,
            configurable: true
        });
        ChartLegend.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Overlay"])) {
                this.m_overlay = obj["Overlay"];
            }
            if (!_isUndefined(obj["Position"])) {
                this.m_position = obj["Position"];
            }
            if (!_isUndefined(obj["Visible"])) {
                this.m_visible = obj["Visible"];
            }
            _handleNavigationPropertyResults(this, obj, ["format", "Format"]);
        };
        ChartLegend.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartLegend;
    })(OfficeExtension.ClientObject);
    Excel.ChartLegend = ChartLegend;
    var ChartLegendFormat = (function (_super) {
        __extends(ChartLegendFormat, _super);
        function ChartLegendFormat() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartLegendFormat.prototype, "fill", {
            get: function () {
                if (!this.m_fill) {
                    this.m_fill = new Excel.ChartFill(this.context, _createPropertyObjectPath(this.context, this, "Fill", false, false));
                }
                return this.m_fill;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartLegendFormat.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Excel.ChartFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });
        ChartLegendFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["fill", "Fill", "font", "Font"]);
        };
        ChartLegendFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartLegendFormat;
    })(OfficeExtension.ClientObject);
    Excel.ChartLegendFormat = ChartLegendFormat;
    var ChartTitle = (function (_super) {
        __extends(ChartTitle, _super);
        function ChartTitle() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartTitle.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.ChartTitleFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartTitle.prototype, "overlay", {
            get: function () {
                _throwIfNotLoaded("overlay", this.m_overlay, "ChartTitle", this._isNull);
                return this.m_overlay;
            },
            set: function (value) {
                this.m_overlay = value;
                _createSetPropertyAction(this.context, this, "Overlay", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartTitle.prototype, "text", {
            get: function () {
                _throwIfNotLoaded("text", this.m_text, "ChartTitle", this._isNull);
                return this.m_text;
            },
            set: function (value) {
                this.m_text = value;
                _createSetPropertyAction(this.context, this, "Text", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartTitle.prototype, "visible", {
            get: function () {
                _throwIfNotLoaded("visible", this.m_visible, "ChartTitle", this._isNull);
                return this.m_visible;
            },
            set: function (value) {
                this.m_visible = value;
                _createSetPropertyAction(this.context, this, "Visible", value);
            },
            enumerable: true,
            configurable: true
        });
        ChartTitle.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Overlay"])) {
                this.m_overlay = obj["Overlay"];
            }
            if (!_isUndefined(obj["Text"])) {
                this.m_text = obj["Text"];
            }
            if (!_isUndefined(obj["Visible"])) {
                this.m_visible = obj["Visible"];
            }
            _handleNavigationPropertyResults(this, obj, ["format", "Format"]);
        };
        ChartTitle.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartTitle;
    })(OfficeExtension.ClientObject);
    Excel.ChartTitle = ChartTitle;
    var ChartTitleFormat = (function (_super) {
        __extends(ChartTitleFormat, _super);
        function ChartTitleFormat() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartTitleFormat.prototype, "fill", {
            get: function () {
                if (!this.m_fill) {
                    this.m_fill = new Excel.ChartFill(this.context, _createPropertyObjectPath(this.context, this, "Fill", false, false));
                }
                return this.m_fill;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartTitleFormat.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Excel.ChartFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });
        ChartTitleFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["fill", "Fill", "font", "Font"]);
        };
        ChartTitleFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartTitleFormat;
    })(OfficeExtension.ClientObject);
    Excel.ChartTitleFormat = ChartTitleFormat;
    var ChartFill = (function (_super) {
        __extends(ChartFill, _super);
        function ChartFill() {
            _super.apply(this, arguments);
        }
        ChartFill.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        ChartFill.prototype.clear = function () {
            _createMethodAction(this.context, this, "Clear", 0 /* Default */, []);
        };
        ChartFill.prototype.setSolidColor = function (color) {
            _createMethodAction(this.context, this, "SetSolidColor", 0 /* Default */, [color]);
        };
        ChartFill.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        return ChartFill;
    })(OfficeExtension.ClientObject);
    Excel.ChartFill = ChartFill;
    var ChartLineFormat = (function (_super) {
        __extends(ChartLineFormat, _super);
        function ChartLineFormat() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartLineFormat.prototype, "color", {
            get: function () {
                _throwIfNotLoaded("color", this.m_color, "ChartLineFormat", this._isNull);
                return this.m_color;
            },
            set: function (value) {
                this.m_color = value;
                _createSetPropertyAction(this.context, this, "Color", value);
            },
            enumerable: true,
            configurable: true
        });
        ChartLineFormat.prototype.clear = function () {
            _createMethodAction(this.context, this, "Clear", 0 /* Default */, []);
        };
        ChartLineFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Color"])) {
                this.m_color = obj["Color"];
            }
        };
        ChartLineFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartLineFormat;
    })(OfficeExtension.ClientObject);
    Excel.ChartLineFormat = ChartLineFormat;
    var ChartFont = (function (_super) {
        __extends(ChartFont, _super);
        function ChartFont() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartFont.prototype, "bold", {
            get: function () {
                _throwIfNotLoaded("bold", this.m_bold, "ChartFont", this._isNull);
                return this.m_bold;
            },
            set: function (value) {
                this.m_bold = value;
                _createSetPropertyAction(this.context, this, "Bold", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartFont.prototype, "color", {
            get: function () {
                _throwIfNotLoaded("color", this.m_color, "ChartFont", this._isNull);
                return this.m_color;
            },
            set: function (value) {
                this.m_color = value;
                _createSetPropertyAction(this.context, this, "Color", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartFont.prototype, "italic", {
            get: function () {
                _throwIfNotLoaded("italic", this.m_italic, "ChartFont", this._isNull);
                return this.m_italic;
            },
            set: function (value) {
                this.m_italic = value;
                _createSetPropertyAction(this.context, this, "Italic", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartFont.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name, "ChartFont", this._isNull);
                return this.m_name;
            },
            set: function (value) {
                this.m_name = value;
                _createSetPropertyAction(this.context, this, "Name", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartFont.prototype, "size", {
            get: function () {
                _throwIfNotLoaded("size", this.m_size, "ChartFont", this._isNull);
                return this.m_size;
            },
            set: function (value) {
                this.m_size = value;
                _createSetPropertyAction(this.context, this, "Size", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartFont.prototype, "underline", {
            get: function () {
                _throwIfNotLoaded("underline", this.m_underline, "ChartFont", this._isNull);
                return this.m_underline;
            },
            set: function (value) {
                this.m_underline = value;
                _createSetPropertyAction(this.context, this, "Underline", value);
            },
            enumerable: true,
            configurable: true
        });
        ChartFont.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Bold"])) {
                this.m_bold = obj["Bold"];
            }
            if (!_isUndefined(obj["Color"])) {
                this.m_color = obj["Color"];
            }
            if (!_isUndefined(obj["Italic"])) {
                this.m_italic = obj["Italic"];
            }
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            if (!_isUndefined(obj["Size"])) {
                this.m_size = obj["Size"];
            }
            if (!_isUndefined(obj["Underline"])) {
                this.m_underline = obj["Underline"];
            }
        };
        ChartFont.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartFont;
    })(OfficeExtension.ClientObject);
    Excel.ChartFont = ChartFont;
    var RangeSort = (function (_super) {
        __extends(RangeSort, _super);
        function RangeSort() {
            _super.apply(this, arguments);
        }
        RangeSort.prototype.apply = function (fields, matchCase, hasHeaders, orientation, method) {
            _createMethodAction(this.context, this, "Apply", 0 /* Default */, [fields, matchCase, hasHeaders, orientation, method]);
        };
        RangeSort.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        return RangeSort;
    })(OfficeExtension.ClientObject);
    Excel.RangeSort = RangeSort;
    var TableSort = (function (_super) {
        __extends(TableSort, _super);
        function TableSort() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(TableSort.prototype, "fields", {
            get: function () {
                _throwIfNotLoaded("fields", this.m_fields, "TableSort", this._isNull);
                return this.m_fields;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableSort.prototype, "matchCase", {
            get: function () {
                _throwIfNotLoaded("matchCase", this.m_matchCase, "TableSort", this._isNull);
                return this.m_matchCase;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TableSort.prototype, "method", {
            get: function () {
                _throwIfNotLoaded("method", this.m_method, "TableSort", this._isNull);
                return this.m_method;
            },
            enumerable: true,
            configurable: true
        });
        TableSort.prototype.apply = function (fields, matchCase, method) {
            _createMethodAction(this.context, this, "Apply", 0 /* Default */, [fields, matchCase, method]);
        };
        TableSort.prototype.clear = function () {
            _createMethodAction(this.context, this, "Clear", 0 /* Default */, []);
        };
        TableSort.prototype.reapply = function () {
            _createMethodAction(this.context, this, "Reapply", 0 /* Default */, []);
        };
        TableSort.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Fields"])) {
                this.m_fields = obj["Fields"];
            }
            if (!_isUndefined(obj["MatchCase"])) {
                this.m_matchCase = obj["MatchCase"];
            }
            if (!_isUndefined(obj["Method"])) {
                this.m_method = obj["Method"];
            }
        };
        TableSort.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return TableSort;
    })(OfficeExtension.ClientObject);
    Excel.TableSort = TableSort;
    var Filter = (function (_super) {
        __extends(Filter, _super);
        function Filter() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Filter.prototype, "criteria", {
            get: function () {
                _throwIfNotLoaded("criteria", this.m_criteria, "Filter", this._isNull);
                return this.m_criteria;
            },
            enumerable: true,
            configurable: true
        });
        Filter.prototype.apply = function (criteria) {
            _createMethodAction(this.context, this, "Apply", 0 /* Default */, [criteria]);
        };
        Filter.prototype.applyBottomItemsFilter = function (count) {
            _createMethodAction(this.context, this, "ApplyBottomItemsFilter", 0 /* Default */, [count]);
        };
        Filter.prototype.applyBottomPercentFilter = function (percent) {
            _createMethodAction(this.context, this, "ApplyBottomPercentFilter", 0 /* Default */, [percent]);
        };
        Filter.prototype.applyCellColorFilter = function (color) {
            _createMethodAction(this.context, this, "ApplyCellColorFilter", 0 /* Default */, [color]);
        };
        Filter.prototype.applyCustomFilter = function (criteria1, criteria2, oper) {
            _createMethodAction(this.context, this, "ApplyCustomFilter", 0 /* Default */, [criteria1, criteria2, oper]);
        };
        Filter.prototype.applyDynamicFilter = function (criteria) {
            _createMethodAction(this.context, this, "ApplyDynamicFilter", 0 /* Default */, [criteria]);
        };
        Filter.prototype.applyFontColorFilter = function (color) {
            _createMethodAction(this.context, this, "ApplyFontColorFilter", 0 /* Default */, [color]);
        };
        Filter.prototype.applyIconFilter = function (icon) {
            _createMethodAction(this.context, this, "ApplyIconFilter", 0 /* Default */, [icon]);
        };
        Filter.prototype.applyTopItemsFilter = function (count) {
            _createMethodAction(this.context, this, "ApplyTopItemsFilter", 0 /* Default */, [count]);
        };
        Filter.prototype.applyTopPercentFilter = function (percent) {
            _createMethodAction(this.context, this, "ApplyTopPercentFilter", 0 /* Default */, [percent]);
        };
        Filter.prototype.applyValuesFilter = function (values) {
            _createMethodAction(this.context, this, "ApplyValuesFilter", 0 /* Default */, [values]);
        };
        Filter.prototype.clear = function () {
            _createMethodAction(this.context, this, "Clear", 0 /* Default */, []);
        };
        Filter.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Criteria"])) {
                this.m_criteria = obj["Criteria"];
            }
        };
        Filter.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return Filter;
    })(OfficeExtension.ClientObject);
    Excel.Filter = Filter;
    var _V1Api = (function (_super) {
        __extends(_V1Api, _super);
        function _V1Api() {
            _super.apply(this, arguments);
        }
        _V1Api.prototype.bindingAddColumns = function (input) {
            var action = _createMethodAction(this.context, this, "BindingAddColumns", 0 /* Default */, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingAddFromNamedItem = function (input) {
            var action = _createMethodAction(this.context, this, "BindingAddFromNamedItem", 1 /* Read */, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingAddFromPrompt = function (input) {
            var action = _createMethodAction(this.context, this, "BindingAddFromPrompt", 1 /* Read */, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingAddFromSelection = function (input) {
            var action = _createMethodAction(this.context, this, "BindingAddFromSelection", 1 /* Read */, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingAddRows = function (input) {
            var action = _createMethodAction(this.context, this, "BindingAddRows", 0 /* Default */, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingClearFormats = function (input) {
            var action = _createMethodAction(this.context, this, "BindingClearFormats", 0 /* Default */, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingDeleteAllDataValues = function (input) {
            var action = _createMethodAction(this.context, this, "BindingDeleteAllDataValues", 0 /* Default */, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingGetAll = function () {
            var action = _createMethodAction(this.context, this, "BindingGetAll", 1 /* Read */, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingGetById = function (input) {
            var action = _createMethodAction(this.context, this, "BindingGetById", 1 /* Read */, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingGetData = function (input) {
            var action = _createMethodAction(this.context, this, "BindingGetData", 1 /* Read */, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingReleaseById = function (input) {
            var action = _createMethodAction(this.context, this, "BindingReleaseById", 1 /* Read */, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingSetData = function (input) {
            var action = _createMethodAction(this.context, this, "BindingSetData", 0 /* Default */, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingSetFormats = function (input) {
            var action = _createMethodAction(this.context, this, "BindingSetFormats", 0 /* Default */, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.bindingSetTableOptions = function (input) {
            var action = _createMethodAction(this.context, this, "BindingSetTableOptions", 0 /* Default */, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.getSelectedData = function (input) {
            var action = _createMethodAction(this.context, this, "GetSelectedData", 1 /* Read */, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.gotoById = function (input) {
            var action = _createMethodAction(this.context, this, "GotoById", 1 /* Read */, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype.setSelectedData = function (input) {
            var action = _createMethodAction(this.context, this, "SetSelectedData", 0 /* Default */, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        _V1Api.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        return _V1Api;
    })(OfficeExtension.ClientObject);
    Excel._V1Api = _V1Api;
    var PivotTableCollection = (function (_super) {
        __extends(PivotTableCollection, _super);
        function PivotTableCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(PivotTableCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "PivotTableCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        PivotTableCollection.prototype.getItem = function (name) {
            return new Excel.PivotTable(this.context, _createIndexerObjectPath(this.context, this, [name]));
        };
        PivotTableCollection.prototype.getItemOrNull = function (name) {
            return new Excel.PivotTable(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNull", 1 /* Read */, [name], false, false, null));
        };
        PivotTableCollection.prototype.refreshAll = function () {
            _createMethodAction(this.context, this, "RefreshAll", 0 /* Default */, []);
        };
        PivotTableCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.PivotTable(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        PivotTableCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return PivotTableCollection;
    })(OfficeExtension.ClientObject);
    Excel.PivotTableCollection = PivotTableCollection;
    var PivotTable = (function (_super) {
        __extends(PivotTable, _super);
        function PivotTable() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(PivotTable.prototype, "worksheet", {
            get: function () {
                if (!this.m_worksheet) {
                    this.m_worksheet = new Excel.Worksheet(this.context, _createPropertyObjectPath(this.context, this, "Worksheet", false, false));
                }
                return this.m_worksheet;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PivotTable.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name, "PivotTable", this._isNull);
                return this.m_name;
            },
            set: function (value) {
                this.m_name = value;
                _createSetPropertyAction(this.context, this, "Name", value);
            },
            enumerable: true,
            configurable: true
        });
        PivotTable.prototype.refresh = function () {
            _createMethodAction(this.context, this, "Refresh", 0 /* Default */, []);
        };
        PivotTable.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            _handleNavigationPropertyResults(this, obj, ["worksheet", "Worksheet"]);
        };
        PivotTable.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return PivotTable;
    })(OfficeExtension.ClientObject);
    Excel.PivotTable = PivotTable;
    var ConditionalFormatCollection = (function (_super) {
        __extends(ConditionalFormatCollection, _super);
        function ConditionalFormatCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ConditionalFormatCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, "ConditionalFormatCollection", this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        ConditionalFormatCollection.prototype.clearAll = function () {
            _createMethodAction(this.context, this, "ClearAll", 1 /* Read */, []);
        };
        ConditionalFormatCollection.prototype.getCount = function () {
            var action = _createMethodAction(this.context, this, "GetCount", 1 /* Read */, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        ConditionalFormatCollection.prototype.getItemAt = function (index) {
            return new Excel.ConditionalFormat(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [index], false, false, null));
        };
        ConditionalFormatCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.ConditionalFormat(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(false, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        ConditionalFormatCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ConditionalFormatCollection;
    })(OfficeExtension.ClientObject);
    Excel.ConditionalFormatCollection = ConditionalFormatCollection;
    var ConditionalFormat = (function (_super) {
        __extends(ConditionalFormat, _super);
        function ConditionalFormat() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ConditionalFormat.prototype, "dataBarOrNull", {
            get: function () {
                if (!this.m_dataBarOrNull) {
                    this.m_dataBarOrNull = new Excel.ConditionalFormatDataBar(this.context, _createPropertyObjectPath(this.context, this, "DataBarOrNull", false, false));
                }
                return this.m_dataBarOrNull;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormat.prototype, "priority", {
            get: function () {
                _throwIfNotLoaded("priority", this.m_priority, "ConditionalFormat", this._isNull);
                return this.m_priority;
            },
            set: function (value) {
                this.m_priority = value;
                _createSetPropertyAction(this.context, this, "Priority", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormat.prototype, "stopIfTrue", {
            get: function () {
                _throwIfNotLoaded("stopIfTrue", this.m_stopIfTrue, "ConditionalFormat", this._isNull);
                return this.m_stopIfTrue;
            },
            set: function (value) {
                this.m_stopIfTrue = value;
                _createSetPropertyAction(this.context, this, "StopIfTrue", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormat.prototype, "type", {
            get: function () {
                _throwIfNotLoaded("type", this.m_type, "ConditionalFormat", this._isNull);
                return this.m_type;
            },
            enumerable: true,
            configurable: true
        });
        ConditionalFormat.prototype.delete = function () {
            _createMethodAction(this.context, this, "Delete", 0 /* Default */, []);
        };
        ConditionalFormat.prototype.getRangeOrNull = function () {
            return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRangeOrNull", 1 /* Read */, [], false, true, null));
        };
        ConditionalFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Priority"])) {
                this.m_priority = obj["Priority"];
            }
            if (!_isUndefined(obj["StopIfTrue"])) {
                this.m_stopIfTrue = obj["StopIfTrue"];
            }
            if (!_isUndefined(obj["Type"])) {
                this.m_type = obj["Type"];
            }
            _handleNavigationPropertyResults(this, obj, ["dataBarOrNull", "DataBarOrNull"]);
        };
        ConditionalFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ConditionalFormat;
    })(OfficeExtension.ClientObject);
    Excel.ConditionalFormat = ConditionalFormat;
    var ConditionalFormatDataBar = (function (_super) {
        __extends(ConditionalFormatDataBar, _super);
        function ConditionalFormatDataBar() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ConditionalFormatDataBar.prototype, "lowerBoundRule", {
            get: function () {
                if (!this.m_lowerBoundRule) {
                    this.m_lowerBoundRule = new Excel.ConditionalFormatDataBarRule(this.context, _createPropertyObjectPath(this.context, this, "LowerBoundRule", false, false));
                }
                return this.m_lowerBoundRule;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormatDataBar.prototype, "negativeFormat", {
            get: function () {
                if (!this.m_negativeFormat) {
                    this.m_negativeFormat = new Excel.ConditionalFormatDataBarNegativeFormat(this.context, _createPropertyObjectPath(this.context, this, "NegativeFormat", false, false));
                }
                return this.m_negativeFormat;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormatDataBar.prototype, "positiveFormat", {
            get: function () {
                if (!this.m_positiveFormat) {
                    this.m_positiveFormat = new Excel.ConditionalFormatDataBarPositiveFormat(this.context, _createPropertyObjectPath(this.context, this, "PositiveFormat", false, false));
                }
                return this.m_positiveFormat;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormatDataBar.prototype, "upperBoundRule", {
            get: function () {
                if (!this.m_upperBoundRule) {
                    this.m_upperBoundRule = new Excel.ConditionalFormatDataBarRule(this.context, _createPropertyObjectPath(this.context, this, "UpperBoundRule", false, false));
                }
                return this.m_upperBoundRule;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormatDataBar.prototype, "axisColor", {
            get: function () {
                _throwIfNotLoaded("axisColor", this.m_axisColor, "ConditionalFormatDataBar", this._isNull);
                return this.m_axisColor;
            },
            set: function (value) {
                this.m_axisColor = value;
                _createSetPropertyAction(this.context, this, "AxisColor", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormatDataBar.prototype, "axisFormat", {
            get: function () {
                _throwIfNotLoaded("axisFormat", this.m_axisFormat, "ConditionalFormatDataBar", this._isNull);
                return this.m_axisFormat;
            },
            set: function (value) {
                this.m_axisFormat = value;
                _createSetPropertyAction(this.context, this, "AxisFormat", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormatDataBar.prototype, "barDirection", {
            get: function () {
                _throwIfNotLoaded("barDirection", this.m_barDirection, "ConditionalFormatDataBar", this._isNull);
                return this.m_barDirection;
            },
            set: function (value) {
                this.m_barDirection = value;
                _createSetPropertyAction(this.context, this, "BarDirection", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormatDataBar.prototype, "showDataBarOnly", {
            get: function () {
                _throwIfNotLoaded("showDataBarOnly", this.m_showDataBarOnly, "ConditionalFormatDataBar", this._isNull);
                return this.m_showDataBarOnly;
            },
            set: function (value) {
                this.m_showDataBarOnly = value;
                _createSetPropertyAction(this.context, this, "ShowDataBarOnly", value);
            },
            enumerable: true,
            configurable: true
        });
        ConditionalFormatDataBar.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["AxisColor"])) {
                this.m_axisColor = obj["AxisColor"];
            }
            if (!_isUndefined(obj["AxisFormat"])) {
                this.m_axisFormat = obj["AxisFormat"];
            }
            if (!_isUndefined(obj["BarDirection"])) {
                this.m_barDirection = obj["BarDirection"];
            }
            if (!_isUndefined(obj["ShowDataBarOnly"])) {
                this.m_showDataBarOnly = obj["ShowDataBarOnly"];
            }
            _handleNavigationPropertyResults(this, obj, ["lowerBoundRule", "LowerBoundRule", "negativeFormat", "NegativeFormat", "positiveFormat", "PositiveFormat", "upperBoundRule", "UpperBoundRule"]);
        };
        ConditionalFormatDataBar.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ConditionalFormatDataBar;
    })(OfficeExtension.ClientObject);
    Excel.ConditionalFormatDataBar = ConditionalFormatDataBar;
    var ConditionalFormatDataBarPositiveFormat = (function (_super) {
        __extends(ConditionalFormatDataBarPositiveFormat, _super);
        function ConditionalFormatDataBarPositiveFormat() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ConditionalFormatDataBarPositiveFormat.prototype, "borderColor", {
            get: function () {
                _throwIfNotLoaded("borderColor", this.m_borderColor, "ConditionalFormatDataBarPositiveFormat", this._isNull);
                return this.m_borderColor;
            },
            set: function (value) {
                this.m_borderColor = value;
                _createSetPropertyAction(this.context, this, "BorderColor", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormatDataBarPositiveFormat.prototype, "fillColor", {
            get: function () {
                _throwIfNotLoaded("fillColor", this.m_fillColor, "ConditionalFormatDataBarPositiveFormat", this._isNull);
                return this.m_fillColor;
            },
            set: function (value) {
                this.m_fillColor = value;
                _createSetPropertyAction(this.context, this, "FillColor", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormatDataBarPositiveFormat.prototype, "gradientFill", {
            get: function () {
                _throwIfNotLoaded("gradientFill", this.m_gradientFill, "ConditionalFormatDataBarPositiveFormat", this._isNull);
                return this.m_gradientFill;
            },
            set: function (value) {
                this.m_gradientFill = value;
                _createSetPropertyAction(this.context, this, "GradientFill", value);
            },
            enumerable: true,
            configurable: true
        });
        ConditionalFormatDataBarPositiveFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["BorderColor"])) {
                this.m_borderColor = obj["BorderColor"];
            }
            if (!_isUndefined(obj["FillColor"])) {
                this.m_fillColor = obj["FillColor"];
            }
            if (!_isUndefined(obj["GradientFill"])) {
                this.m_gradientFill = obj["GradientFill"];
            }
        };
        ConditionalFormatDataBarPositiveFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ConditionalFormatDataBarPositiveFormat;
    })(OfficeExtension.ClientObject);
    Excel.ConditionalFormatDataBarPositiveFormat = ConditionalFormatDataBarPositiveFormat;
    var ConditionalFormatDataBarNegativeFormat = (function (_super) {
        __extends(ConditionalFormatDataBarNegativeFormat, _super);
        function ConditionalFormatDataBarNegativeFormat() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ConditionalFormatDataBarNegativeFormat.prototype, "borderColor", {
            get: function () {
                _throwIfNotLoaded("borderColor", this.m_borderColor, "ConditionalFormatDataBarNegativeFormat", this._isNull);
                return this.m_borderColor;
            },
            set: function (value) {
                this.m_borderColor = value;
                _createSetPropertyAction(this.context, this, "BorderColor", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormatDataBarNegativeFormat.prototype, "fillColor", {
            get: function () {
                _throwIfNotLoaded("fillColor", this.m_fillColor, "ConditionalFormatDataBarNegativeFormat", this._isNull);
                return this.m_fillColor;
            },
            set: function (value) {
                this.m_fillColor = value;
                _createSetPropertyAction(this.context, this, "FillColor", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormatDataBarNegativeFormat.prototype, "matchPositiveBorderColor", {
            get: function () {
                _throwIfNotLoaded("matchPositiveBorderColor", this.m_matchPositiveBorderColor, "ConditionalFormatDataBarNegativeFormat", this._isNull);
                return this.m_matchPositiveBorderColor;
            },
            set: function (value) {
                this.m_matchPositiveBorderColor = value;
                _createSetPropertyAction(this.context, this, "MatchPositiveBorderColor", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormatDataBarNegativeFormat.prototype, "matchPositiveFillColor", {
            get: function () {
                _throwIfNotLoaded("matchPositiveFillColor", this.m_matchPositiveFillColor, "ConditionalFormatDataBarNegativeFormat", this._isNull);
                return this.m_matchPositiveFillColor;
            },
            set: function (value) {
                this.m_matchPositiveFillColor = value;
                _createSetPropertyAction(this.context, this, "MatchPositiveFillColor", value);
            },
            enumerable: true,
            configurable: true
        });
        ConditionalFormatDataBarNegativeFormat.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["BorderColor"])) {
                this.m_borderColor = obj["BorderColor"];
            }
            if (!_isUndefined(obj["FillColor"])) {
                this.m_fillColor = obj["FillColor"];
            }
            if (!_isUndefined(obj["MatchPositiveBorderColor"])) {
                this.m_matchPositiveBorderColor = obj["MatchPositiveBorderColor"];
            }
            if (!_isUndefined(obj["MatchPositiveFillColor"])) {
                this.m_matchPositiveFillColor = obj["MatchPositiveFillColor"];
            }
        };
        ConditionalFormatDataBarNegativeFormat.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ConditionalFormatDataBarNegativeFormat;
    })(OfficeExtension.ClientObject);
    Excel.ConditionalFormatDataBarNegativeFormat = ConditionalFormatDataBarNegativeFormat;
    var ConditionalFormatDataBarRule = (function (_super) {
        __extends(ConditionalFormatDataBarRule, _super);
        function ConditionalFormatDataBarRule() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ConditionalFormatDataBarRule.prototype, "formula", {
            get: function () {
                _throwIfNotLoaded("formula", this.m_formula, "ConditionalFormatDataBarRule", this._isNull);
                return this.m_formula;
            },
            set: function (value) {
                this.m_formula = value;
                _createSetPropertyAction(this.context, this, "Formula", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormatDataBarRule.prototype, "formulaLocal", {
            get: function () {
                _throwIfNotLoaded("formulaLocal", this.m_formulaLocal, "ConditionalFormatDataBarRule", this._isNull);
                return this.m_formulaLocal;
            },
            set: function (value) {
                this.m_formulaLocal = value;
                _createSetPropertyAction(this.context, this, "FormulaLocal", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormatDataBarRule.prototype, "formulaR1C1", {
            get: function () {
                _throwIfNotLoaded("formulaR1C1", this.m_formulaR1C1, "ConditionalFormatDataBarRule", this._isNull);
                return this.m_formulaR1C1;
            },
            set: function (value) {
                this.m_formulaR1C1 = value;
                _createSetPropertyAction(this.context, this, "FormulaR1C1", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ConditionalFormatDataBarRule.prototype, "type", {
            get: function () {
                _throwIfNotLoaded("type", this.m_type, "ConditionalFormatDataBarRule", this._isNull);
                return this.m_type;
            },
            set: function (value) {
                this.m_type = value;
                _createSetPropertyAction(this.context, this, "Type", value);
            },
            enumerable: true,
            configurable: true
        });
        ConditionalFormatDataBarRule.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Formula"])) {
                this.m_formula = obj["Formula"];
            }
            if (!_isUndefined(obj["FormulaLocal"])) {
                this.m_formulaLocal = obj["FormulaLocal"];
            }
            if (!_isUndefined(obj["FormulaR1C1"])) {
                this.m_formulaR1C1 = obj["FormulaR1C1"];
            }
            if (!_isUndefined(obj["Type"])) {
                this.m_type = obj["Type"];
            }
        };
        ConditionalFormatDataBarRule.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ConditionalFormatDataBarRule;
    })(OfficeExtension.ClientObject);
    Excel.ConditionalFormatDataBarRule = ConditionalFormatDataBarRule;
    var BindingType;
    (function (BindingType) {
        BindingType.range = "Range";
        BindingType.table = "Table";
        BindingType.text = "Text";
    })(BindingType = Excel.BindingType || (Excel.BindingType = {}));
    var BorderIndex;
    (function (BorderIndex) {
        BorderIndex.edgeTop = "EdgeTop";
        BorderIndex.edgeBottom = "EdgeBottom";
        BorderIndex.edgeLeft = "EdgeLeft";
        BorderIndex.edgeRight = "EdgeRight";
        BorderIndex.insideVertical = "InsideVertical";
        BorderIndex.insideHorizontal = "InsideHorizontal";
        BorderIndex.diagonalDown = "DiagonalDown";
        BorderIndex.diagonalUp = "DiagonalUp";
    })(BorderIndex = Excel.BorderIndex || (Excel.BorderIndex = {}));
    var BorderLineStyle;
    (function (BorderLineStyle) {
        BorderLineStyle.none = "None";
        BorderLineStyle.continuous = "Continuous";
        BorderLineStyle.dash = "Dash";
        BorderLineStyle.dashDot = "DashDot";
        BorderLineStyle.dashDotDot = "DashDotDot";
        BorderLineStyle.dot = "Dot";
        BorderLineStyle.double = "Double";
        BorderLineStyle.slantDashDot = "SlantDashDot";
    })(BorderLineStyle = Excel.BorderLineStyle || (Excel.BorderLineStyle = {}));
    var BorderWeight;
    (function (BorderWeight) {
        BorderWeight.hairline = "Hairline";
        BorderWeight.thin = "Thin";
        BorderWeight.medium = "Medium";
        BorderWeight.thick = "Thick";
    })(BorderWeight = Excel.BorderWeight || (Excel.BorderWeight = {}));
    var CalculationMode;
    (function (CalculationMode) {
        CalculationMode.automatic = "Automatic";
        CalculationMode.automaticExceptTables = "AutomaticExceptTables";
        CalculationMode.manual = "Manual";
    })(CalculationMode = Excel.CalculationMode || (Excel.CalculationMode = {}));
    var CalculationType;
    (function (CalculationType) {
        CalculationType.recalculate = "Recalculate";
        CalculationType.full = "Full";
        CalculationType.fullRebuild = "FullRebuild";
    })(CalculationType = Excel.CalculationType || (Excel.CalculationType = {}));
    var ClearApplyTo;
    (function (ClearApplyTo) {
        ClearApplyTo.all = "All";
        ClearApplyTo.formats = "Formats";
        ClearApplyTo.contents = "Contents";
    })(ClearApplyTo = Excel.ClearApplyTo || (Excel.ClearApplyTo = {}));
    var ChartDataLabelPosition;
    (function (ChartDataLabelPosition) {
        ChartDataLabelPosition.invalid = "Invalid";
        ChartDataLabelPosition.none = "None";
        ChartDataLabelPosition.center = "Center";
        ChartDataLabelPosition.insideEnd = "InsideEnd";
        ChartDataLabelPosition.insideBase = "InsideBase";
        ChartDataLabelPosition.outsideEnd = "OutsideEnd";
        ChartDataLabelPosition.left = "Left";
        ChartDataLabelPosition.right = "Right";
        ChartDataLabelPosition.top = "Top";
        ChartDataLabelPosition.bottom = "Bottom";
        ChartDataLabelPosition.bestFit = "BestFit";
        ChartDataLabelPosition.callout = "Callout";
    })(ChartDataLabelPosition = Excel.ChartDataLabelPosition || (Excel.ChartDataLabelPosition = {}));
    var ChartLegendPosition;
    (function (ChartLegendPosition) {
        ChartLegendPosition.invalid = "Invalid";
        ChartLegendPosition.top = "Top";
        ChartLegendPosition.bottom = "Bottom";
        ChartLegendPosition.left = "Left";
        ChartLegendPosition.right = "Right";
        ChartLegendPosition.corner = "Corner";
        ChartLegendPosition.custom = "Custom";
    })(ChartLegendPosition = Excel.ChartLegendPosition || (Excel.ChartLegendPosition = {}));
    var ChartSeriesBy;
    (function (ChartSeriesBy) {
        ChartSeriesBy.auto = "Auto";
        ChartSeriesBy.columns = "Columns";
        ChartSeriesBy.rows = "Rows";
    })(ChartSeriesBy = Excel.ChartSeriesBy || (Excel.ChartSeriesBy = {}));
    var ChartType;
    (function (ChartType) {
        ChartType.invalid = "Invalid";
        ChartType.columnClustered = "ColumnClustered";
        ChartType.columnStacked = "ColumnStacked";
        ChartType.columnStacked100 = "ColumnStacked100";
        ChartType._3DColumnClustered = "3DColumnClustered";
        ChartType._3DColumnStacked = "3DColumnStacked";
        ChartType._3DColumnStacked100 = "3DColumnStacked100";
        ChartType.barClustered = "BarClustered";
        ChartType.barStacked = "BarStacked";
        ChartType.barStacked100 = "BarStacked100";
        ChartType._3DBarClustered = "3DBarClustered";
        ChartType._3DBarStacked = "3DBarStacked";
        ChartType._3DBarStacked100 = "3DBarStacked100";
        ChartType.lineStacked = "LineStacked";
        ChartType.lineStacked100 = "LineStacked100";
        ChartType.lineMarkers = "LineMarkers";
        ChartType.lineMarkersStacked = "LineMarkersStacked";
        ChartType.lineMarkersStacked100 = "LineMarkersStacked100";
        ChartType.pieOfPie = "PieOfPie";
        ChartType.pieExploded = "PieExploded";
        ChartType._3DPieExploded = "3DPieExploded";
        ChartType.barOfPie = "BarOfPie";
        ChartType.xyscatterSmooth = "XYScatterSmooth";
        ChartType.xyscatterSmoothNoMarkers = "XYScatterSmoothNoMarkers";
        ChartType.xyscatterLines = "XYScatterLines";
        ChartType.xyscatterLinesNoMarkers = "XYScatterLinesNoMarkers";
        ChartType.areaStacked = "AreaStacked";
        ChartType.areaStacked100 = "AreaStacked100";
        ChartType._3DAreaStacked = "3DAreaStacked";
        ChartType._3DAreaStacked100 = "3DAreaStacked100";
        ChartType.doughnutExploded = "DoughnutExploded";
        ChartType.radarMarkers = "RadarMarkers";
        ChartType.radarFilled = "RadarFilled";
        ChartType.surface = "Surface";
        ChartType.surfaceWireframe = "SurfaceWireframe";
        ChartType.surfaceTopView = "SurfaceTopView";
        ChartType.surfaceTopViewWireframe = "SurfaceTopViewWireframe";
        ChartType.bubble = "Bubble";
        ChartType.bubble3DEffect = "Bubble3DEffect";
        ChartType.stockHLC = "StockHLC";
        ChartType.stockOHLC = "StockOHLC";
        ChartType.stockVHLC = "StockVHLC";
        ChartType.stockVOHLC = "StockVOHLC";
        ChartType.cylinderColClustered = "CylinderColClustered";
        ChartType.cylinderColStacked = "CylinderColStacked";
        ChartType.cylinderColStacked100 = "CylinderColStacked100";
        ChartType.cylinderBarClustered = "CylinderBarClustered";
        ChartType.cylinderBarStacked = "CylinderBarStacked";
        ChartType.cylinderBarStacked100 = "CylinderBarStacked100";
        ChartType.cylinderCol = "CylinderCol";
        ChartType.coneColClustered = "ConeColClustered";
        ChartType.coneColStacked = "ConeColStacked";
        ChartType.coneColStacked100 = "ConeColStacked100";
        ChartType.coneBarClustered = "ConeBarClustered";
        ChartType.coneBarStacked = "ConeBarStacked";
        ChartType.coneBarStacked100 = "ConeBarStacked100";
        ChartType.coneCol = "ConeCol";
        ChartType.pyramidColClustered = "PyramidColClustered";
        ChartType.pyramidColStacked = "PyramidColStacked";
        ChartType.pyramidColStacked100 = "PyramidColStacked100";
        ChartType.pyramidBarClustered = "PyramidBarClustered";
        ChartType.pyramidBarStacked = "PyramidBarStacked";
        ChartType.pyramidBarStacked100 = "PyramidBarStacked100";
        ChartType.pyramidCol = "PyramidCol";
        ChartType._3DColumn = "3DColumn";
        ChartType.line = "Line";
        ChartType._3DLine = "3DLine";
        ChartType._3DPie = "3DPie";
        ChartType.pie = "Pie";
        ChartType.xyscatter = "XYScatter";
        ChartType._3DArea = "3DArea";
        ChartType.area = "Area";
        ChartType.doughnut = "Doughnut";
        ChartType.radar = "Radar";
    })(ChartType = Excel.ChartType || (Excel.ChartType = {}));
    var ChartUnderlineStyle;
    (function (ChartUnderlineStyle) {
        ChartUnderlineStyle.none = "None";
        ChartUnderlineStyle.single = "Single";
    })(ChartUnderlineStyle = Excel.ChartUnderlineStyle || (Excel.ChartUnderlineStyle = {}));
    var ConditionalFormatDataBarAxisFormat;
    (function (ConditionalFormatDataBarAxisFormat) {
        ConditionalFormatDataBarAxisFormat.automatic = "Automatic";
        ConditionalFormatDataBarAxisFormat.none = "None";
        ConditionalFormatDataBarAxisFormat.cellMidPoint = "CellMidPoint";
    })(ConditionalFormatDataBarAxisFormat = Excel.ConditionalFormatDataBarAxisFormat || (Excel.ConditionalFormatDataBarAxisFormat = {}));
    var ConditionalFormatDataBarDirection;
    (function (ConditionalFormatDataBarDirection) {
        ConditionalFormatDataBarDirection.context = "Context";
        ConditionalFormatDataBarDirection.leftToRight = "LeftToRight";
        ConditionalFormatDataBarDirection.rightToLeft = "RightToLeft";
    })(ConditionalFormatDataBarDirection = Excel.ConditionalFormatDataBarDirection || (Excel.ConditionalFormatDataBarDirection = {}));
    var ConditionalFormatDirection;
    (function (ConditionalFormatDirection) {
        ConditionalFormatDirection.top = "Top";
        ConditionalFormatDirection.bottom = "Bottom";
    })(ConditionalFormatDirection = Excel.ConditionalFormatDirection || (Excel.ConditionalFormatDirection = {}));
    var ConditionalFormatType;
    (function (ConditionalFormatType) {
        ConditionalFormatType.custom = "Custom";
        ConditionalFormatType.dataBar = "DataBar";
        ConditionalFormatType.colorScale = "ColorScale";
        ConditionalFormatType.iconSet = "IconSet";
    })(ConditionalFormatType = Excel.ConditionalFormatType || (Excel.ConditionalFormatType = {}));
    var ConditionalFormatRuleType;
    (function (ConditionalFormatRuleType) {
        ConditionalFormatRuleType.automatic = "Automatic";
        ConditionalFormatRuleType.lowestValue = "LowestValue";
        ConditionalFormatRuleType.highestValue = "HighestValue";
        ConditionalFormatRuleType.number = "Number";
        ConditionalFormatRuleType.percent = "Percent";
        ConditionalFormatRuleType.formula = "Formula";
        ConditionalFormatRuleType.percentile = "Percentile";
    })(ConditionalFormatRuleType = Excel.ConditionalFormatRuleType || (Excel.ConditionalFormatRuleType = {}));
    var ConditionalRangeFormatRuleType;
    (function (ConditionalRangeFormatRuleType) {
        ConditionalRangeFormatRuleType.blank = "Blank";
        ConditionalRangeFormatRuleType.expression = "Expression";
        ConditionalRangeFormatRuleType.between = "Between";
        ConditionalRangeFormatRuleType.notBetween = "NotBetween";
        ConditionalRangeFormatRuleType.count = "Count";
        ConditionalRangeFormatRuleType.percent = "Percent";
        ConditionalRangeFormatRuleType.average = "Average";
        ConditionalRangeFormatRuleType.unique = "Unique";
        ConditionalRangeFormatRuleType.error = "Error";
        ConditionalRangeFormatRuleType.textContains = "TextContains";
        ConditionalRangeFormatRuleType.dateOccurring = "DateOccurring";
    })(ConditionalRangeFormatRuleType = Excel.ConditionalRangeFormatRuleType || (Excel.ConditionalRangeFormatRuleType = {}));
    var DeleteShiftDirection;
    (function (DeleteShiftDirection) {
        DeleteShiftDirection.up = "Up";
        DeleteShiftDirection.left = "Left";
    })(DeleteShiftDirection = Excel.DeleteShiftDirection || (Excel.DeleteShiftDirection = {}));
    var DynamicFilterCriteria;
    (function (DynamicFilterCriteria) {
        DynamicFilterCriteria.unknown = "Unknown";
        DynamicFilterCriteria.aboveAverage = "AboveAverage";
        DynamicFilterCriteria.allDatesInPeriodApril = "AllDatesInPeriodApril";
        DynamicFilterCriteria.allDatesInPeriodAugust = "AllDatesInPeriodAugust";
        DynamicFilterCriteria.allDatesInPeriodDecember = "AllDatesInPeriodDecember";
        DynamicFilterCriteria.allDatesInPeriodFebruray = "AllDatesInPeriodFebruray";
        DynamicFilterCriteria.allDatesInPeriodJanuary = "AllDatesInPeriodJanuary";
        DynamicFilterCriteria.allDatesInPeriodJuly = "AllDatesInPeriodJuly";
        DynamicFilterCriteria.allDatesInPeriodJune = "AllDatesInPeriodJune";
        DynamicFilterCriteria.allDatesInPeriodMarch = "AllDatesInPeriodMarch";
        DynamicFilterCriteria.allDatesInPeriodMay = "AllDatesInPeriodMay";
        DynamicFilterCriteria.allDatesInPeriodNovember = "AllDatesInPeriodNovember";
        DynamicFilterCriteria.allDatesInPeriodOctober = "AllDatesInPeriodOctober";
        DynamicFilterCriteria.allDatesInPeriodQuarter1 = "AllDatesInPeriodQuarter1";
        DynamicFilterCriteria.allDatesInPeriodQuarter2 = "AllDatesInPeriodQuarter2";
        DynamicFilterCriteria.allDatesInPeriodQuarter3 = "AllDatesInPeriodQuarter3";
        DynamicFilterCriteria.allDatesInPeriodQuarter4 = "AllDatesInPeriodQuarter4";
        DynamicFilterCriteria.allDatesInPeriodSeptember = "AllDatesInPeriodSeptember";
        DynamicFilterCriteria.belowAverage = "BelowAverage";
        DynamicFilterCriteria.lastMonth = "LastMonth";
        DynamicFilterCriteria.lastQuarter = "LastQuarter";
        DynamicFilterCriteria.lastWeek = "LastWeek";
        DynamicFilterCriteria.lastYear = "LastYear";
        DynamicFilterCriteria.nextMonth = "NextMonth";
        DynamicFilterCriteria.nextQuarter = "NextQuarter";
        DynamicFilterCriteria.nextWeek = "NextWeek";
        DynamicFilterCriteria.nextYear = "NextYear";
        DynamicFilterCriteria.thisMonth = "ThisMonth";
        DynamicFilterCriteria.thisQuarter = "ThisQuarter";
        DynamicFilterCriteria.thisWeek = "ThisWeek";
        DynamicFilterCriteria.thisYear = "ThisYear";
        DynamicFilterCriteria.today = "Today";
        DynamicFilterCriteria.tomorrow = "Tomorrow";
        DynamicFilterCriteria.yearToDate = "YearToDate";
        DynamicFilterCriteria.yesterday = "Yesterday";
    })(DynamicFilterCriteria = Excel.DynamicFilterCriteria || (Excel.DynamicFilterCriteria = {}));
    var FilterDatetimeSpecificity;
    (function (FilterDatetimeSpecificity) {
        FilterDatetimeSpecificity.year = "Year";
        FilterDatetimeSpecificity.month = "Month";
        FilterDatetimeSpecificity.day = "Day";
        FilterDatetimeSpecificity.hour = "Hour";
        FilterDatetimeSpecificity.minute = "Minute";
        FilterDatetimeSpecificity.second = "Second";
    })(FilterDatetimeSpecificity = Excel.FilterDatetimeSpecificity || (Excel.FilterDatetimeSpecificity = {}));
    var FilterOn;
    (function (FilterOn) {
        FilterOn.bottomItems = "BottomItems";
        FilterOn.bottomPercent = "BottomPercent";
        FilterOn.cellColor = "CellColor";
        FilterOn.dynamic = "Dynamic";
        FilterOn.fontColor = "FontColor";
        FilterOn.values = "Values";
        FilterOn.topItems = "TopItems";
        FilterOn.topPercent = "TopPercent";
        FilterOn.icon = "Icon";
        FilterOn.custom = "Custom";
    })(FilterOn = Excel.FilterOn || (Excel.FilterOn = {}));
    var FilterOperator;
    (function (FilterOperator) {
        FilterOperator.and = "And";
        FilterOperator.or = "Or";
    })(FilterOperator = Excel.FilterOperator || (Excel.FilterOperator = {}));
    var HorizontalAlignment;
    (function (HorizontalAlignment) {
        HorizontalAlignment.general = "General";
        HorizontalAlignment.left = "Left";
        HorizontalAlignment.center = "Center";
        HorizontalAlignment.right = "Right";
        HorizontalAlignment.fill = "Fill";
        HorizontalAlignment.justify = "Justify";
        HorizontalAlignment.centerAcrossSelection = "CenterAcrossSelection";
        HorizontalAlignment.distributed = "Distributed";
    })(HorizontalAlignment = Excel.HorizontalAlignment || (Excel.HorizontalAlignment = {}));
    var IconSet;
    (function (IconSet) {
        IconSet.invalid = "Invalid";
        IconSet.threeArrows = "ThreeArrows";
        IconSet.threeArrowsGray = "ThreeArrowsGray";
        IconSet.threeFlags = "ThreeFlags";
        IconSet.threeTrafficLights1 = "ThreeTrafficLights1";
        IconSet.threeTrafficLights2 = "ThreeTrafficLights2";
        IconSet.threeSigns = "ThreeSigns";
        IconSet.threeSymbols = "ThreeSymbols";
        IconSet.threeSymbols2 = "ThreeSymbols2";
        IconSet.fourArrows = "FourArrows";
        IconSet.fourArrowsGray = "FourArrowsGray";
        IconSet.fourRedToBlack = "FourRedToBlack";
        IconSet.fourRating = "FourRating";
        IconSet.fourTrafficLights = "FourTrafficLights";
        IconSet.fiveArrows = "FiveArrows";
        IconSet.fiveArrowsGray = "FiveArrowsGray";
        IconSet.fiveRating = "FiveRating";
        IconSet.fiveQuarters = "FiveQuarters";
        IconSet.threeStars = "ThreeStars";
        IconSet.threeTriangles = "ThreeTriangles";
        IconSet.fiveBoxes = "FiveBoxes";
    })(IconSet = Excel.IconSet || (Excel.IconSet = {}));
    var ImageFittingMode;
    (function (ImageFittingMode) {
        ImageFittingMode.fit = "Fit";
        ImageFittingMode.fitAndCenter = "FitAndCenter";
        ImageFittingMode.fill = "Fill";
    })(ImageFittingMode = Excel.ImageFittingMode || (Excel.ImageFittingMode = {}));
    var InsertShiftDirection;
    (function (InsertShiftDirection) {
        InsertShiftDirection.down = "Down";
        InsertShiftDirection.right = "Right";
    })(InsertShiftDirection = Excel.InsertShiftDirection || (Excel.InsertShiftDirection = {}));
    var NamedItemType;
    (function (NamedItemType) {
        NamedItemType.string = "String";
        NamedItemType.integer = "Integer";
        NamedItemType.double = "Double";
        NamedItemType.boolean = "Boolean";
        NamedItemType.range = "Range";
    })(NamedItemType = Excel.NamedItemType || (Excel.NamedItemType = {}));
    var RangeUnderlineStyle;
    (function (RangeUnderlineStyle) {
        RangeUnderlineStyle.none = "None";
        RangeUnderlineStyle.single = "Single";
        RangeUnderlineStyle.double = "Double";
        RangeUnderlineStyle.singleAccountant = "SingleAccountant";
        RangeUnderlineStyle.doubleAccountant = "DoubleAccountant";
    })(RangeUnderlineStyle = Excel.RangeUnderlineStyle || (Excel.RangeUnderlineStyle = {}));
    var SheetVisibility;
    (function (SheetVisibility) {
        SheetVisibility.visible = "Visible";
        SheetVisibility.hidden = "Hidden";
        SheetVisibility.veryHidden = "VeryHidden";
    })(SheetVisibility = Excel.SheetVisibility || (Excel.SheetVisibility = {}));
    var RangeValueType;
    (function (RangeValueType) {
        RangeValueType.unknown = "Unknown";
        RangeValueType.empty = "Empty";
        RangeValueType.string = "String";
        RangeValueType.integer = "Integer";
        RangeValueType.double = "Double";
        RangeValueType.boolean = "Boolean";
        RangeValueType.error = "Error";
    })(RangeValueType = Excel.RangeValueType || (Excel.RangeValueType = {}));
    var SortOrientation;
    (function (SortOrientation) {
        SortOrientation.rows = "Rows";
        SortOrientation.columns = "Columns";
    })(SortOrientation = Excel.SortOrientation || (Excel.SortOrientation = {}));
    var SortOn;
    (function (SortOn) {
        SortOn.value = "Value";
        SortOn.cellColor = "CellColor";
        SortOn.fontColor = "FontColor";
        SortOn.icon = "Icon";
    })(SortOn = Excel.SortOn || (Excel.SortOn = {}));
    var SortDataOption;
    (function (SortDataOption) {
        SortDataOption.normal = "Normal";
        SortDataOption.textAsNumber = "TextAsNumber";
    })(SortDataOption = Excel.SortDataOption || (Excel.SortDataOption = {}));
    var SortMethod;
    (function (SortMethod) {
        SortMethod.pinYin = "PinYin";
        SortMethod.strokeCount = "StrokeCount";
    })(SortMethod = Excel.SortMethod || (Excel.SortMethod = {}));
    var VerticalAlignment;
    (function (VerticalAlignment) {
        VerticalAlignment.top = "Top";
        VerticalAlignment.center = "Center";
        VerticalAlignment.bottom = "Bottom";
        VerticalAlignment.justify = "Justify";
        VerticalAlignment.distributed = "Distributed";
    })(VerticalAlignment = Excel.VerticalAlignment || (Excel.VerticalAlignment = {}));
    var FunctionResult = (function (_super) {
        __extends(FunctionResult, _super);
        function FunctionResult() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(FunctionResult.prototype, "error", {
            get: function () {
                _throwIfNotLoaded("error", this.m_error, "FunctionResult", this._isNull);
                return this.m_error;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(FunctionResult.prototype, "value", {
            get: function () {
                _throwIfNotLoaded("value", this.m_value, "FunctionResult", this._isNull);
                return this.m_value;
            },
            enumerable: true,
            configurable: true
        });
        FunctionResult.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Error"])) {
                this.m_error = obj["Error"];
            }
            if (!_isUndefined(obj["Value"])) {
                this.m_value = obj["Value"];
            }
        };
        FunctionResult.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return FunctionResult;
    })(OfficeExtension.ClientObject);
    Excel.FunctionResult = FunctionResult;
    var Functions = (function (_super) {
        __extends(Functions, _super);
        function Functions() {
            _super.apply(this, arguments);
        }
        Functions.prototype.abs = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Abs", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.accrInt = function (issue, firstInterest, settlement, rate, par, frequency, basis, calcMethod) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AccrInt", 0 /* Default */, [issue, firstInterest, settlement, rate, par, frequency, basis, calcMethod], false, true, null));
        };
        Functions.prototype.accrIntM = function (issue, settlement, rate, par, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AccrIntM", 0 /* Default */, [issue, settlement, rate, par, basis], false, true, null));
        };
        Functions.prototype.acos = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Acos", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.acosh = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Acosh", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.acot = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Acot", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.acoth = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Acoth", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.amorDegrc = function (cost, datePurchased, firstPeriod, salvage, period, rate, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AmorDegrc", 0 /* Default */, [cost, datePurchased, firstPeriod, salvage, period, rate, basis], false, true, null));
        };
        Functions.prototype.amorLinc = function (cost, datePurchased, firstPeriod, salvage, period, rate, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AmorLinc", 0 /* Default */, [cost, datePurchased, firstPeriod, salvage, period, rate, basis], false, true, null));
        };
        Functions.prototype.and = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "And", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.arabic = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Arabic", 0 /* Default */, [text], false, true, null));
        };
        Functions.prototype.areas = function (reference) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Areas", 0 /* Default */, [reference], false, true, null));
        };
        Functions.prototype.asc = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Asc", 0 /* Default */, [text], false, true, null));
        };
        Functions.prototype.asin = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Asin", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.asinh = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Asinh", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.atan = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Atan", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.atan2 = function (xNum, yNum) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Atan2", 0 /* Default */, [xNum, yNum], false, true, null));
        };
        Functions.prototype.atanh = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Atanh", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.aveDev = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AveDev", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.average = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Average", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.averageA = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AverageA", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.averageIf = function (range, criteria, averageRange) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AverageIf", 0 /* Default */, [range, criteria, averageRange], false, true, null));
        };
        Functions.prototype.averageIfs = function (averageRange) {
            var values = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                values[_i - 1] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AverageIfs", 0 /* Default */, [averageRange, values], false, true, null));
        };
        Functions.prototype.bahtText = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "BahtText", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.base = function (number, radix, minLength) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Base", 0 /* Default */, [number, radix, minLength], false, true, null));
        };
        Functions.prototype.besselI = function (x, n) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "BesselI", 0 /* Default */, [x, n], false, true, null));
        };
        Functions.prototype.besselJ = function (x, n) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "BesselJ", 0 /* Default */, [x, n], false, true, null));
        };
        Functions.prototype.besselK = function (x, n) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "BesselK", 0 /* Default */, [x, n], false, true, null));
        };
        Functions.prototype.besselY = function (x, n) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "BesselY", 0 /* Default */, [x, n], false, true, null));
        };
        Functions.prototype.beta_Dist = function (x, alpha, beta, cumulative, A, B) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Beta_Dist", 0 /* Default */, [x, alpha, beta, cumulative, A, B], false, true, null));
        };
        Functions.prototype.beta_Inv = function (probability, alpha, beta, A, B) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Beta_Inv", 0 /* Default */, [probability, alpha, beta, A, B], false, true, null));
        };
        Functions.prototype.bin2Dec = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bin2Dec", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.bin2Hex = function (number, places) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bin2Hex", 0 /* Default */, [number, places], false, true, null));
        };
        Functions.prototype.bin2Oct = function (number, places) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bin2Oct", 0 /* Default */, [number, places], false, true, null));
        };
        Functions.prototype.binom_Dist = function (numberS, trials, probabilityS, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Binom_Dist", 0 /* Default */, [numberS, trials, probabilityS, cumulative], false, true, null));
        };
        Functions.prototype.binom_Dist_Range = function (trials, probabilityS, numberS, numberS2) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Binom_Dist_Range", 0 /* Default */, [trials, probabilityS, numberS, numberS2], false, true, null));
        };
        Functions.prototype.binom_Inv = function (trials, probabilityS, alpha) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Binom_Inv", 0 /* Default */, [trials, probabilityS, alpha], false, true, null));
        };
        Functions.prototype.bitand = function (number1, number2) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bitand", 0 /* Default */, [number1, number2], false, true, null));
        };
        Functions.prototype.bitlshift = function (number, shiftAmount) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bitlshift", 0 /* Default */, [number, shiftAmount], false, true, null));
        };
        Functions.prototype.bitor = function (number1, number2) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bitor", 0 /* Default */, [number1, number2], false, true, null));
        };
        Functions.prototype.bitrshift = function (number, shiftAmount) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bitrshift", 0 /* Default */, [number, shiftAmount], false, true, null));
        };
        Functions.prototype.bitxor = function (number1, number2) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bitxor", 0 /* Default */, [number1, number2], false, true, null));
        };
        Functions.prototype.ceiling_Math = function (number, significance, mode) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Ceiling_Math", 0 /* Default */, [number, significance, mode], false, true, null));
        };
        Functions.prototype.ceiling_Precise = function (number, significance) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Ceiling_Precise", 0 /* Default */, [number, significance], false, true, null));
        };
        Functions.prototype.char = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Char", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.chiSq_Dist = function (x, degFreedom, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ChiSq_Dist", 0 /* Default */, [x, degFreedom, cumulative], false, true, null));
        };
        Functions.prototype.chiSq_Dist_RT = function (x, degFreedom) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ChiSq_Dist_RT", 0 /* Default */, [x, degFreedom], false, true, null));
        };
        Functions.prototype.chiSq_Inv = function (probability, degFreedom) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ChiSq_Inv", 0 /* Default */, [probability, degFreedom], false, true, null));
        };
        Functions.prototype.chiSq_Inv_RT = function (probability, degFreedom) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ChiSq_Inv_RT", 0 /* Default */, [probability, degFreedom], false, true, null));
        };
        Functions.prototype.choose = function (indexNum) {
            var values = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                values[_i - 1] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Choose", 0 /* Default */, [indexNum, values], false, true, null));
        };
        Functions.prototype.clean = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Clean", 0 /* Default */, [text], false, true, null));
        };
        Functions.prototype.code = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Code", 0 /* Default */, [text], false, true, null));
        };
        Functions.prototype.columns = function (array) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Columns", 0 /* Default */, [array], false, true, null));
        };
        Functions.prototype.combin = function (number, numberChosen) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Combin", 0 /* Default */, [number, numberChosen], false, true, null));
        };
        Functions.prototype.combina = function (number, numberChosen) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Combina", 0 /* Default */, [number, numberChosen], false, true, null));
        };
        Functions.prototype.complex = function (realNum, iNum, suffix) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Complex", 0 /* Default */, [realNum, iNum, suffix], false, true, null));
        };
        Functions.prototype.concatenate = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Concatenate", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.confidence_Norm = function (alpha, standardDev, size) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Confidence_Norm", 0 /* Default */, [alpha, standardDev, size], false, true, null));
        };
        Functions.prototype.confidence_T = function (alpha, standardDev, size) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Confidence_T", 0 /* Default */, [alpha, standardDev, size], false, true, null));
        };
        Functions.prototype.convert = function (number, fromUnit, toUnit) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Convert", 0 /* Default */, [number, fromUnit, toUnit], false, true, null));
        };
        Functions.prototype.cos = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Cos", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.cosh = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Cosh", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.cot = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Cot", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.coth = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Coth", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.count = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Count", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.countA = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CountA", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.countBlank = function (range) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CountBlank", 0 /* Default */, [range], false, true, null));
        };
        Functions.prototype.countIf = function (range, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CountIf", 0 /* Default */, [range, criteria], false, true, null));
        };
        Functions.prototype.countIfs = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CountIfs", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.coupDayBs = function (settlement, maturity, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CoupDayBs", 0 /* Default */, [settlement, maturity, frequency, basis], false, true, null));
        };
        Functions.prototype.coupDays = function (settlement, maturity, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CoupDays", 0 /* Default */, [settlement, maturity, frequency, basis], false, true, null));
        };
        Functions.prototype.coupDaysNc = function (settlement, maturity, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CoupDaysNc", 0 /* Default */, [settlement, maturity, frequency, basis], false, true, null));
        };
        Functions.prototype.coupNcd = function (settlement, maturity, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CoupNcd", 0 /* Default */, [settlement, maturity, frequency, basis], false, true, null));
        };
        Functions.prototype.coupNum = function (settlement, maturity, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CoupNum", 0 /* Default */, [settlement, maturity, frequency, basis], false, true, null));
        };
        Functions.prototype.coupPcd = function (settlement, maturity, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CoupPcd", 0 /* Default */, [settlement, maturity, frequency, basis], false, true, null));
        };
        Functions.prototype.csc = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Csc", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.csch = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Csch", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.cumIPmt = function (rate, nper, pv, startPeriod, endPeriod, type) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CumIPmt", 0 /* Default */, [rate, nper, pv, startPeriod, endPeriod, type], false, true, null));
        };
        Functions.prototype.cumPrinc = function (rate, nper, pv, startPeriod, endPeriod, type) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CumPrinc", 0 /* Default */, [rate, nper, pv, startPeriod, endPeriod, type], false, true, null));
        };
        Functions.prototype.daverage = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DAverage", 0 /* Default */, [database, field, criteria], false, true, null));
        };
        Functions.prototype.dcount = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DCount", 0 /* Default */, [database, field, criteria], false, true, null));
        };
        Functions.prototype.dcountA = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DCountA", 0 /* Default */, [database, field, criteria], false, true, null));
        };
        Functions.prototype.dget = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DGet", 0 /* Default */, [database, field, criteria], false, true, null));
        };
        Functions.prototype.dmax = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DMax", 0 /* Default */, [database, field, criteria], false, true, null));
        };
        Functions.prototype.dmin = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DMin", 0 /* Default */, [database, field, criteria], false, true, null));
        };
        Functions.prototype.dproduct = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DProduct", 0 /* Default */, [database, field, criteria], false, true, null));
        };
        Functions.prototype.dstDev = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DStDev", 0 /* Default */, [database, field, criteria], false, true, null));
        };
        Functions.prototype.dstDevP = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DStDevP", 0 /* Default */, [database, field, criteria], false, true, null));
        };
        Functions.prototype.dsum = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DSum", 0 /* Default */, [database, field, criteria], false, true, null));
        };
        Functions.prototype.dvar = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DVar", 0 /* Default */, [database, field, criteria], false, true, null));
        };
        Functions.prototype.dvarP = function (database, field, criteria) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DVarP", 0 /* Default */, [database, field, criteria], false, true, null));
        };
        Functions.prototype.date = function (year, month, day) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Date", 0 /* Default */, [year, month, day], false, true, null));
        };
        Functions.prototype.datevalue = function (dateText) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Datevalue", 0 /* Default */, [dateText], false, true, null));
        };
        Functions.prototype.day = function (serialNumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Day", 0 /* Default */, [serialNumber], false, true, null));
        };
        Functions.prototype.days = function (endDate, startDate) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Days", 0 /* Default */, [endDate, startDate], false, true, null));
        };
        Functions.prototype.days360 = function (startDate, endDate, method) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Days360", 0 /* Default */, [startDate, endDate, method], false, true, null));
        };
        Functions.prototype.db = function (cost, salvage, life, period, month) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Db", 0 /* Default */, [cost, salvage, life, period, month], false, true, null));
        };
        Functions.prototype.dbcs = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Dbcs", 0 /* Default */, [text], false, true, null));
        };
        Functions.prototype.ddb = function (cost, salvage, life, period, factor) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Ddb", 0 /* Default */, [cost, salvage, life, period, factor], false, true, null));
        };
        Functions.prototype.dec2Bin = function (number, places) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Dec2Bin", 0 /* Default */, [number, places], false, true, null));
        };
        Functions.prototype.dec2Hex = function (number, places) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Dec2Hex", 0 /* Default */, [number, places], false, true, null));
        };
        Functions.prototype.dec2Oct = function (number, places) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Dec2Oct", 0 /* Default */, [number, places], false, true, null));
        };
        Functions.prototype.decimal = function (number, radix) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Decimal", 0 /* Default */, [number, radix], false, true, null));
        };
        Functions.prototype.degrees = function (angle) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Degrees", 0 /* Default */, [angle], false, true, null));
        };
        Functions.prototype.delta = function (number1, number2) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Delta", 0 /* Default */, [number1, number2], false, true, null));
        };
        Functions.prototype.devSq = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DevSq", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.disc = function (settlement, maturity, pr, redemption, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Disc", 0 /* Default */, [settlement, maturity, pr, redemption, basis], false, true, null));
        };
        Functions.prototype.dollar = function (number, decimals) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Dollar", 0 /* Default */, [number, decimals], false, true, null));
        };
        Functions.prototype.dollarDe = function (fractionalDollar, fraction) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DollarDe", 0 /* Default */, [fractionalDollar, fraction], false, true, null));
        };
        Functions.prototype.dollarFr = function (decimalDollar, fraction) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DollarFr", 0 /* Default */, [decimalDollar, fraction], false, true, null));
        };
        Functions.prototype.duration = function (settlement, maturity, coupon, yld, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Duration", 0 /* Default */, [settlement, maturity, coupon, yld, frequency, basis], false, true, null));
        };
        Functions.prototype.ecma_Ceiling = function (number, significance) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ECMA_Ceiling", 0 /* Default */, [number, significance], false, true, null));
        };
        Functions.prototype.edate = function (startDate, months) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "EDate", 0 /* Default */, [startDate, months], false, true, null));
        };
        Functions.prototype.effect = function (nominalRate, npery) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Effect", 0 /* Default */, [nominalRate, npery], false, true, null));
        };
        Functions.prototype.eoMonth = function (startDate, months) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "EoMonth", 0 /* Default */, [startDate, months], false, true, null));
        };
        Functions.prototype.erf = function (lowerLimit, upperLimit) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Erf", 0 /* Default */, [lowerLimit, upperLimit], false, true, null));
        };
        Functions.prototype.erfC = function (x) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ErfC", 0 /* Default */, [x], false, true, null));
        };
        Functions.prototype.erfC_Precise = function (X) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ErfC_Precise", 0 /* Default */, [X], false, true, null));
        };
        Functions.prototype.erf_Precise = function (X) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Erf_Precise", 0 /* Default */, [X], false, true, null));
        };
        Functions.prototype.error_Type = function (errorVal) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Error_Type", 0 /* Default */, [errorVal], false, true, null));
        };
        Functions.prototype.even = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Even", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.exact = function (text1, text2) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Exact", 0 /* Default */, [text1, text2], false, true, null));
        };
        Functions.prototype.exp = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Exp", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.expon_Dist = function (x, lambda, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Expon_Dist", 0 /* Default */, [x, lambda, cumulative], false, true, null));
        };
        Functions.prototype.fvschedule = function (principal, schedule) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "FVSchedule", 0 /* Default */, [principal, schedule], false, true, null));
        };
        Functions.prototype.f_Dist = function (x, degFreedom1, degFreedom2, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "F_Dist", 0 /* Default */, [x, degFreedom1, degFreedom2, cumulative], false, true, null));
        };
        Functions.prototype.f_Dist_RT = function (x, degFreedom1, degFreedom2) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "F_Dist_RT", 0 /* Default */, [x, degFreedom1, degFreedom2], false, true, null));
        };
        Functions.prototype.f_Inv = function (probability, degFreedom1, degFreedom2) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "F_Inv", 0 /* Default */, [probability, degFreedom1, degFreedom2], false, true, null));
        };
        Functions.prototype.f_Inv_RT = function (probability, degFreedom1, degFreedom2) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "F_Inv_RT", 0 /* Default */, [probability, degFreedom1, degFreedom2], false, true, null));
        };
        Functions.prototype.fact = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Fact", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.factDouble = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "FactDouble", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.false = function () {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "False", 0 /* Default */, [], false, true, null));
        };
        Functions.prototype.find = function (findText, withinText, startNum) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Find", 0 /* Default */, [findText, withinText, startNum], false, true, null));
        };
        Functions.prototype.findB = function (findText, withinText, startNum) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "FindB", 0 /* Default */, [findText, withinText, startNum], false, true, null));
        };
        Functions.prototype.fisher = function (x) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Fisher", 0 /* Default */, [x], false, true, null));
        };
        Functions.prototype.fisherInv = function (y) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "FisherInv", 0 /* Default */, [y], false, true, null));
        };
        Functions.prototype.fixed = function (number, decimals, noCommas) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Fixed", 0 /* Default */, [number, decimals, noCommas], false, true, null));
        };
        Functions.prototype.floor_Math = function (number, significance, mode) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Floor_Math", 0 /* Default */, [number, significance, mode], false, true, null));
        };
        Functions.prototype.floor_Precise = function (number, significance) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Floor_Precise", 0 /* Default */, [number, significance], false, true, null));
        };
        Functions.prototype.fv = function (rate, nper, pmt, pv, type) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Fv", 0 /* Default */, [rate, nper, pmt, pv, type], false, true, null));
        };
        Functions.prototype.gamma = function (x) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Gamma", 0 /* Default */, [x], false, true, null));
        };
        Functions.prototype.gammaLn = function (x) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "GammaLn", 0 /* Default */, [x], false, true, null));
        };
        Functions.prototype.gammaLn_Precise = function (x) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "GammaLn_Precise", 0 /* Default */, [x], false, true, null));
        };
        Functions.prototype.gamma_Dist = function (x, alpha, beta, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Gamma_Dist", 0 /* Default */, [x, alpha, beta, cumulative], false, true, null));
        };
        Functions.prototype.gamma_Inv = function (probability, alpha, beta) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Gamma_Inv", 0 /* Default */, [probability, alpha, beta], false, true, null));
        };
        Functions.prototype.gauss = function (x) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Gauss", 0 /* Default */, [x], false, true, null));
        };
        Functions.prototype.gcd = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Gcd", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.geStep = function (number, step) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "GeStep", 0 /* Default */, [number, step], false, true, null));
        };
        Functions.prototype.geoMean = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "GeoMean", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.hlookup = function (lookupValue, tableArray, rowIndexNum, rangeLookup) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "HLookup", 0 /* Default */, [lookupValue, tableArray, rowIndexNum, rangeLookup], false, true, null));
        };
        Functions.prototype.harMean = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "HarMean", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.hex2Bin = function (number, places) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Hex2Bin", 0 /* Default */, [number, places], false, true, null));
        };
        Functions.prototype.hex2Dec = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Hex2Dec", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.hex2Oct = function (number, places) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Hex2Oct", 0 /* Default */, [number, places], false, true, null));
        };
        Functions.prototype.hour = function (serialNumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Hour", 0 /* Default */, [serialNumber], false, true, null));
        };
        Functions.prototype.hypGeom_Dist = function (sampleS, numberSample, populationS, numberPop, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "HypGeom_Dist", 0 /* Default */, [sampleS, numberSample, populationS, numberPop, cumulative], false, true, null));
        };
        Functions.prototype.hyperlink = function (linkLocation, friendlyName) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Hyperlink", 0 /* Default */, [linkLocation, friendlyName], false, true, null));
        };
        Functions.prototype.iso_Ceiling = function (number, significance) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ISO_Ceiling", 0 /* Default */, [number, significance], false, true, null));
        };
        Functions.prototype.if = function (logicalTest, valueIfTrue, valueIfFalse) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "If", 0 /* Default */, [logicalTest, valueIfTrue, valueIfFalse], false, true, null));
        };
        Functions.prototype.imAbs = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImAbs", 0 /* Default */, [inumber], false, true, null));
        };
        Functions.prototype.imArgument = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImArgument", 0 /* Default */, [inumber], false, true, null));
        };
        Functions.prototype.imConjugate = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImConjugate", 0 /* Default */, [inumber], false, true, null));
        };
        Functions.prototype.imCos = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImCos", 0 /* Default */, [inumber], false, true, null));
        };
        Functions.prototype.imCosh = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImCosh", 0 /* Default */, [inumber], false, true, null));
        };
        Functions.prototype.imCot = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImCot", 0 /* Default */, [inumber], false, true, null));
        };
        Functions.prototype.imCsc = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImCsc", 0 /* Default */, [inumber], false, true, null));
        };
        Functions.prototype.imCsch = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImCsch", 0 /* Default */, [inumber], false, true, null));
        };
        Functions.prototype.imDiv = function (inumber1, inumber2) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImDiv", 0 /* Default */, [inumber1, inumber2], false, true, null));
        };
        Functions.prototype.imExp = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImExp", 0 /* Default */, [inumber], false, true, null));
        };
        Functions.prototype.imLn = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImLn", 0 /* Default */, [inumber], false, true, null));
        };
        Functions.prototype.imLog10 = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImLog10", 0 /* Default */, [inumber], false, true, null));
        };
        Functions.prototype.imLog2 = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImLog2", 0 /* Default */, [inumber], false, true, null));
        };
        Functions.prototype.imPower = function (inumber, number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImPower", 0 /* Default */, [inumber, number], false, true, null));
        };
        Functions.prototype.imProduct = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImProduct", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.imReal = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImReal", 0 /* Default */, [inumber], false, true, null));
        };
        Functions.prototype.imSec = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImSec", 0 /* Default */, [inumber], false, true, null));
        };
        Functions.prototype.imSech = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImSech", 0 /* Default */, [inumber], false, true, null));
        };
        Functions.prototype.imSin = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImSin", 0 /* Default */, [inumber], false, true, null));
        };
        Functions.prototype.imSinh = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImSinh", 0 /* Default */, [inumber], false, true, null));
        };
        Functions.prototype.imSqrt = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImSqrt", 0 /* Default */, [inumber], false, true, null));
        };
        Functions.prototype.imSub = function (inumber1, inumber2) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImSub", 0 /* Default */, [inumber1, inumber2], false, true, null));
        };
        Functions.prototype.imSum = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImSum", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.imTan = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImTan", 0 /* Default */, [inumber], false, true, null));
        };
        Functions.prototype.imaginary = function (inumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Imaginary", 0 /* Default */, [inumber], false, true, null));
        };
        Functions.prototype.int = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Int", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.intRate = function (settlement, maturity, investment, redemption, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IntRate", 0 /* Default */, [settlement, maturity, investment, redemption, basis], false, true, null));
        };
        Functions.prototype.ipmt = function (rate, per, nper, pv, fv, type) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Ipmt", 0 /* Default */, [rate, per, nper, pv, fv, type], false, true, null));
        };
        Functions.prototype.irr = function (values, guess) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Irr", 0 /* Default */, [values, guess], false, true, null));
        };
        Functions.prototype.isErr = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsErr", 0 /* Default */, [value], false, true, null));
        };
        Functions.prototype.isError = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsError", 0 /* Default */, [value], false, true, null));
        };
        Functions.prototype.isEven = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsEven", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.isFormula = function (reference) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsFormula", 0 /* Default */, [reference], false, true, null));
        };
        Functions.prototype.isLogical = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsLogical", 0 /* Default */, [value], false, true, null));
        };
        Functions.prototype.isNA = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsNA", 0 /* Default */, [value], false, true, null));
        };
        Functions.prototype.isNonText = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsNonText", 0 /* Default */, [value], false, true, null));
        };
        Functions.prototype.isNumber = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsNumber", 0 /* Default */, [value], false, true, null));
        };
        Functions.prototype.isOdd = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsOdd", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.isText = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsText", 0 /* Default */, [value], false, true, null));
        };
        Functions.prototype.isoWeekNum = function (date) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsoWeekNum", 0 /* Default */, [date], false, true, null));
        };
        Functions.prototype.ispmt = function (rate, per, nper, pv) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Ispmt", 0 /* Default */, [rate, per, nper, pv], false, true, null));
        };
        Functions.prototype.isref = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Isref", 0 /* Default */, [value], false, true, null));
        };
        Functions.prototype.kurt = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Kurt", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.large = function (array, k) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Large", 0 /* Default */, [array, k], false, true, null));
        };
        Functions.prototype.lcm = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Lcm", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.left = function (text, numChars) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Left", 0 /* Default */, [text, numChars], false, true, null));
        };
        Functions.prototype.leftb = function (text, numBytes) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Leftb", 0 /* Default */, [text, numBytes], false, true, null));
        };
        Functions.prototype.len = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Len", 0 /* Default */, [text], false, true, null));
        };
        Functions.prototype.lenb = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Lenb", 0 /* Default */, [text], false, true, null));
        };
        Functions.prototype.ln = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Ln", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.log = function (number, base) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Log", 0 /* Default */, [number, base], false, true, null));
        };
        Functions.prototype.log10 = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Log10", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.logNorm_Dist = function (x, mean, standardDev, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "LogNorm_Dist", 0 /* Default */, [x, mean, standardDev, cumulative], false, true, null));
        };
        Functions.prototype.logNorm_Inv = function (probability, mean, standardDev) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "LogNorm_Inv", 0 /* Default */, [probability, mean, standardDev], false, true, null));
        };
        Functions.prototype.lookup = function (lookupValue, lookupVector, resultVector) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Lookup", 0 /* Default */, [lookupValue, lookupVector, resultVector], false, true, null));
        };
        Functions.prototype.lower = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Lower", 0 /* Default */, [text], false, true, null));
        };
        Functions.prototype.mduration = function (settlement, maturity, coupon, yld, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "MDuration", 0 /* Default */, [settlement, maturity, coupon, yld, frequency, basis], false, true, null));
        };
        Functions.prototype.mirr = function (values, financeRate, reinvestRate) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "MIrr", 0 /* Default */, [values, financeRate, reinvestRate], false, true, null));
        };
        Functions.prototype.mround = function (number, multiple) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "MRound", 0 /* Default */, [number, multiple], false, true, null));
        };
        Functions.prototype.match = function (lookupValue, lookupArray, matchType) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Match", 0 /* Default */, [lookupValue, lookupArray, matchType], false, true, null));
        };
        Functions.prototype.max = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Max", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.maxA = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "MaxA", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.median = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Median", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.mid = function (text, startNum, numChars) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Mid", 0 /* Default */, [text, startNum, numChars], false, true, null));
        };
        Functions.prototype.midb = function (text, startNum, numBytes) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Midb", 0 /* Default */, [text, startNum, numBytes], false, true, null));
        };
        Functions.prototype.min = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Min", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.minA = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "MinA", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.minute = function (serialNumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Minute", 0 /* Default */, [serialNumber], false, true, null));
        };
        Functions.prototype.mod = function (number, divisor) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Mod", 0 /* Default */, [number, divisor], false, true, null));
        };
        Functions.prototype.month = function (serialNumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Month", 0 /* Default */, [serialNumber], false, true, null));
        };
        Functions.prototype.multiNomial = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "MultiNomial", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.n = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "N", 0 /* Default */, [value], false, true, null));
        };
        Functions.prototype.nper = function (rate, pmt, pv, fv, type) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "NPer", 0 /* Default */, [rate, pmt, pv, fv, type], false, true, null));
        };
        Functions.prototype.na = function () {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Na", 0 /* Default */, [], false, true, null));
        };
        Functions.prototype.negBinom_Dist = function (numberF, numberS, probabilityS, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "NegBinom_Dist", 0 /* Default */, [numberF, numberS, probabilityS, cumulative], false, true, null));
        };
        Functions.prototype.networkDays = function (startDate, endDate, holidays) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "NetworkDays", 0 /* Default */, [startDate, endDate, holidays], false, true, null));
        };
        Functions.prototype.networkDays_Intl = function (startDate, endDate, weekend, holidays) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "NetworkDays_Intl", 0 /* Default */, [startDate, endDate, weekend, holidays], false, true, null));
        };
        Functions.prototype.nominal = function (effectRate, npery) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Nominal", 0 /* Default */, [effectRate, npery], false, true, null));
        };
        Functions.prototype.norm_Dist = function (x, mean, standardDev, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Norm_Dist", 0 /* Default */, [x, mean, standardDev, cumulative], false, true, null));
        };
        Functions.prototype.norm_Inv = function (probability, mean, standardDev) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Norm_Inv", 0 /* Default */, [probability, mean, standardDev], false, true, null));
        };
        Functions.prototype.norm_S_Dist = function (z, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Norm_S_Dist", 0 /* Default */, [z, cumulative], false, true, null));
        };
        Functions.prototype.norm_S_Inv = function (probability) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Norm_S_Inv", 0 /* Default */, [probability], false, true, null));
        };
        Functions.prototype.not = function (logical) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Not", 0 /* Default */, [logical], false, true, null));
        };
        Functions.prototype.now = function () {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Now", 0 /* Default */, [], false, true, null));
        };
        Functions.prototype.npv = function (rate) {
            var values = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                values[_i - 1] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Npv", 0 /* Default */, [rate, values], false, true, null));
        };
        Functions.prototype.numberValue = function (text, decimalSeparator, groupSeparator) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "NumberValue", 0 /* Default */, [text, decimalSeparator, groupSeparator], false, true, null));
        };
        Functions.prototype.oct2Bin = function (number, places) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Oct2Bin", 0 /* Default */, [number, places], false, true, null));
        };
        Functions.prototype.oct2Dec = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Oct2Dec", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.oct2Hex = function (number, places) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Oct2Hex", 0 /* Default */, [number, places], false, true, null));
        };
        Functions.prototype.odd = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Odd", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.oddFPrice = function (settlement, maturity, issue, firstCoupon, rate, yld, redemption, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "OddFPrice", 0 /* Default */, [settlement, maturity, issue, firstCoupon, rate, yld, redemption, frequency, basis], false, true, null));
        };
        Functions.prototype.oddFYield = function (settlement, maturity, issue, firstCoupon, rate, pr, redemption, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "OddFYield", 0 /* Default */, [settlement, maturity, issue, firstCoupon, rate, pr, redemption, frequency, basis], false, true, null));
        };
        Functions.prototype.oddLPrice = function (settlement, maturity, lastInterest, rate, yld, redemption, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "OddLPrice", 0 /* Default */, [settlement, maturity, lastInterest, rate, yld, redemption, frequency, basis], false, true, null));
        };
        Functions.prototype.oddLYield = function (settlement, maturity, lastInterest, rate, pr, redemption, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "OddLYield", 0 /* Default */, [settlement, maturity, lastInterest, rate, pr, redemption, frequency, basis], false, true, null));
        };
        Functions.prototype.or = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Or", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.pduration = function (rate, pv, fv) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "PDuration", 0 /* Default */, [rate, pv, fv], false, true, null));
        };
        Functions.prototype.percentRank_Exc = function (array, x, significance) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "PercentRank_Exc", 0 /* Default */, [array, x, significance], false, true, null));
        };
        Functions.prototype.percentRank_Inc = function (array, x, significance) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "PercentRank_Inc", 0 /* Default */, [array, x, significance], false, true, null));
        };
        Functions.prototype.percentile_Exc = function (array, k) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Percentile_Exc", 0 /* Default */, [array, k], false, true, null));
        };
        Functions.prototype.percentile_Inc = function (array, k) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Percentile_Inc", 0 /* Default */, [array, k], false, true, null));
        };
        Functions.prototype.permut = function (number, numberChosen) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Permut", 0 /* Default */, [number, numberChosen], false, true, null));
        };
        Functions.prototype.permutationa = function (number, numberChosen) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Permutationa", 0 /* Default */, [number, numberChosen], false, true, null));
        };
        Functions.prototype.phi = function (x) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Phi", 0 /* Default */, [x], false, true, null));
        };
        Functions.prototype.pi = function () {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Pi", 0 /* Default */, [], false, true, null));
        };
        Functions.prototype.pmt = function (rate, nper, pv, fv, type) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Pmt", 0 /* Default */, [rate, nper, pv, fv, type], false, true, null));
        };
        Functions.prototype.poisson_Dist = function (x, mean, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Poisson_Dist", 0 /* Default */, [x, mean, cumulative], false, true, null));
        };
        Functions.prototype.power = function (number, power) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Power", 0 /* Default */, [number, power], false, true, null));
        };
        Functions.prototype.ppmt = function (rate, per, nper, pv, fv, type) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Ppmt", 0 /* Default */, [rate, per, nper, pv, fv, type], false, true, null));
        };
        Functions.prototype.price = function (settlement, maturity, rate, yld, redemption, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Price", 0 /* Default */, [settlement, maturity, rate, yld, redemption, frequency, basis], false, true, null));
        };
        Functions.prototype.priceDisc = function (settlement, maturity, discount, redemption, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "PriceDisc", 0 /* Default */, [settlement, maturity, discount, redemption, basis], false, true, null));
        };
        Functions.prototype.priceMat = function (settlement, maturity, issue, rate, yld, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "PriceMat", 0 /* Default */, [settlement, maturity, issue, rate, yld, basis], false, true, null));
        };
        Functions.prototype.product = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Product", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.proper = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Proper", 0 /* Default */, [text], false, true, null));
        };
        Functions.prototype.pv = function (rate, nper, pmt, fv, type) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Pv", 0 /* Default */, [rate, nper, pmt, fv, type], false, true, null));
        };
        Functions.prototype.quartile_Exc = function (array, quart) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Quartile_Exc", 0 /* Default */, [array, quart], false, true, null));
        };
        Functions.prototype.quartile_Inc = function (array, quart) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Quartile_Inc", 0 /* Default */, [array, quart], false, true, null));
        };
        Functions.prototype.quotient = function (numerator, denominator) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Quotient", 0 /* Default */, [numerator, denominator], false, true, null));
        };
        Functions.prototype.radians = function (angle) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Radians", 0 /* Default */, [angle], false, true, null));
        };
        Functions.prototype.rand = function () {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rand", 0 /* Default */, [], false, true, null));
        };
        Functions.prototype.randBetween = function (bottom, top) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "RandBetween", 0 /* Default */, [bottom, top], false, true, null));
        };
        Functions.prototype.rank_Avg = function (number, ref, order) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rank_Avg", 0 /* Default */, [number, ref, order], false, true, null));
        };
        Functions.prototype.rank_Eq = function (number, ref, order) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rank_Eq", 0 /* Default */, [number, ref, order], false, true, null));
        };
        Functions.prototype.rate = function (nper, pmt, pv, fv, type, guess) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rate", 0 /* Default */, [nper, pmt, pv, fv, type, guess], false, true, null));
        };
        Functions.prototype.received = function (settlement, maturity, investment, discount, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Received", 0 /* Default */, [settlement, maturity, investment, discount, basis], false, true, null));
        };
        Functions.prototype.replace = function (oldText, startNum, numChars, newText) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Replace", 0 /* Default */, [oldText, startNum, numChars, newText], false, true, null));
        };
        Functions.prototype.replaceB = function (oldText, startNum, numBytes, newText) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ReplaceB", 0 /* Default */, [oldText, startNum, numBytes, newText], false, true, null));
        };
        Functions.prototype.rept = function (text, numberTimes) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rept", 0 /* Default */, [text, numberTimes], false, true, null));
        };
        Functions.prototype.right = function (text, numChars) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Right", 0 /* Default */, [text, numChars], false, true, null));
        };
        Functions.prototype.rightb = function (text, numBytes) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rightb", 0 /* Default */, [text, numBytes], false, true, null));
        };
        Functions.prototype.roman = function (number, form) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Roman", 0 /* Default */, [number, form], false, true, null));
        };
        Functions.prototype.round = function (number, numDigits) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Round", 0 /* Default */, [number, numDigits], false, true, null));
        };
        Functions.prototype.roundDown = function (number, numDigits) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "RoundDown", 0 /* Default */, [number, numDigits], false, true, null));
        };
        Functions.prototype.roundUp = function (number, numDigits) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "RoundUp", 0 /* Default */, [number, numDigits], false, true, null));
        };
        Functions.prototype.rows = function (array) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rows", 0 /* Default */, [array], false, true, null));
        };
        Functions.prototype.rri = function (nper, pv, fv) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rri", 0 /* Default */, [nper, pv, fv], false, true, null));
        };
        Functions.prototype.sec = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sec", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.sech = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sech", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.second = function (serialNumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Second", 0 /* Default */, [serialNumber], false, true, null));
        };
        Functions.prototype.seriesSum = function (x, n, m, coefficients) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "SeriesSum", 0 /* Default */, [x, n, m, coefficients], false, true, null));
        };
        Functions.prototype.sheet = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sheet", 0 /* Default */, [value], false, true, null));
        };
        Functions.prototype.sheets = function (reference) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sheets", 0 /* Default */, [reference], false, true, null));
        };
        Functions.prototype.sign = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sign", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.sin = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sin", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.sinh = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sinh", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.skew = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Skew", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.skew_p = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Skew_p", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.sln = function (cost, salvage, life) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sln", 0 /* Default */, [cost, salvage, life], false, true, null));
        };
        Functions.prototype.small = function (array, k) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Small", 0 /* Default */, [array, k], false, true, null));
        };
        Functions.prototype.sqrt = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sqrt", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.sqrtPi = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "SqrtPi", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.stDevA = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "StDevA", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.stDevPA = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "StDevPA", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.stDev_P = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "StDev_P", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.stDev_S = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "StDev_S", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.standardize = function (x, mean, standardDev) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Standardize", 0 /* Default */, [x, mean, standardDev], false, true, null));
        };
        Functions.prototype.substitute = function (text, oldText, newText, instanceNum) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Substitute", 0 /* Default */, [text, oldText, newText, instanceNum], false, true, null));
        };
        Functions.prototype.subtotal = function (functionNum) {
            var values = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                values[_i - 1] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Subtotal", 0 /* Default */, [functionNum, values], false, true, null));
        };
        Functions.prototype.sum = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sum", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.sumIf = function (range, criteria, sumRange) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "SumIf", 0 /* Default */, [range, criteria, sumRange], false, true, null));
        };
        Functions.prototype.sumIfs = function (sumRange) {
            var values = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                values[_i - 1] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "SumIfs", 0 /* Default */, [sumRange, values], false, true, null));
        };
        Functions.prototype.sumSq = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "SumSq", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.syd = function (cost, salvage, life, per) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Syd", 0 /* Default */, [cost, salvage, life, per], false, true, null));
        };
        Functions.prototype.t = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "T", 0 /* Default */, [value], false, true, null));
        };
        Functions.prototype.tbillEq = function (settlement, maturity, discount) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "TBillEq", 0 /* Default */, [settlement, maturity, discount], false, true, null));
        };
        Functions.prototype.tbillPrice = function (settlement, maturity, discount) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "TBillPrice", 0 /* Default */, [settlement, maturity, discount], false, true, null));
        };
        Functions.prototype.tbillYield = function (settlement, maturity, pr) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "TBillYield", 0 /* Default */, [settlement, maturity, pr], false, true, null));
        };
        Functions.prototype.t_Dist = function (x, degFreedom, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "T_Dist", 0 /* Default */, [x, degFreedom, cumulative], false, true, null));
        };
        Functions.prototype.t_Dist_2T = function (x, degFreedom) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "T_Dist_2T", 0 /* Default */, [x, degFreedom], false, true, null));
        };
        Functions.prototype.t_Dist_RT = function (x, degFreedom) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "T_Dist_RT", 0 /* Default */, [x, degFreedom], false, true, null));
        };
        Functions.prototype.t_Inv = function (probability, degFreedom) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "T_Inv", 0 /* Default */, [probability, degFreedom], false, true, null));
        };
        Functions.prototype.t_Inv_2T = function (probability, degFreedom) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "T_Inv_2T", 0 /* Default */, [probability, degFreedom], false, true, null));
        };
        Functions.prototype.tan = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Tan", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.tanh = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Tanh", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.text = function (value, formatText) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Text", 0 /* Default */, [value, formatText], false, true, null));
        };
        Functions.prototype.time = function (hour, minute, second) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Time", 0 /* Default */, [hour, minute, second], false, true, null));
        };
        Functions.prototype.timevalue = function (timeText) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Timevalue", 0 /* Default */, [timeText], false, true, null));
        };
        Functions.prototype.today = function () {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Today", 0 /* Default */, [], false, true, null));
        };
        Functions.prototype.trim = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Trim", 0 /* Default */, [text], false, true, null));
        };
        Functions.prototype.trimMean = function (array, percent) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "TrimMean", 0 /* Default */, [array, percent], false, true, null));
        };
        Functions.prototype.true = function () {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "True", 0 /* Default */, [], false, true, null));
        };
        Functions.prototype.trunc = function (number, numDigits) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Trunc", 0 /* Default */, [number, numDigits], false, true, null));
        };
        Functions.prototype.type = function (value) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Type", 0 /* Default */, [value], false, true, null));
        };
        Functions.prototype.usdollar = function (number, decimals) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "USDollar", 0 /* Default */, [number, decimals], false, true, null));
        };
        Functions.prototype.unichar = function (number) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Unichar", 0 /* Default */, [number], false, true, null));
        };
        Functions.prototype.unicode = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Unicode", 0 /* Default */, [text], false, true, null));
        };
        Functions.prototype.upper = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Upper", 0 /* Default */, [text], false, true, null));
        };
        Functions.prototype.vlookup = function (lookupValue, tableArray, colIndexNum, rangeLookup) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "VLookup", 0 /* Default */, [lookupValue, tableArray, colIndexNum, rangeLookup], false, true, null));
        };
        Functions.prototype.value = function (text) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Value", 0 /* Default */, [text], false, true, null));
        };
        Functions.prototype.varA = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "VarA", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.varPA = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "VarPA", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.var_P = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Var_P", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.var_S = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Var_S", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.vdb = function (cost, salvage, life, startPeriod, endPeriod, factor, noSwitch) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Vdb", 0 /* Default */, [cost, salvage, life, startPeriod, endPeriod, factor, noSwitch], false, true, null));
        };
        Functions.prototype.weekNum = function (serialNumber, returnType) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "WeekNum", 0 /* Default */, [serialNumber, returnType], false, true, null));
        };
        Functions.prototype.weekday = function (serialNumber, returnType) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Weekday", 0 /* Default */, [serialNumber, returnType], false, true, null));
        };
        Functions.prototype.weibull_Dist = function (x, alpha, beta, cumulative) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Weibull_Dist", 0 /* Default */, [x, alpha, beta, cumulative], false, true, null));
        };
        Functions.prototype.workDay = function (startDate, days, holidays) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "WorkDay", 0 /* Default */, [startDate, days, holidays], false, true, null));
        };
        Functions.prototype.workDay_Intl = function (startDate, days, weekend, holidays) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "WorkDay_Intl", 0 /* Default */, [startDate, days, weekend, holidays], false, true, null));
        };
        Functions.prototype.xirr = function (values, dates, guess) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Xirr", 0 /* Default */, [values, dates, guess], false, true, null));
        };
        Functions.prototype.xnpv = function (rate, values, dates) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Xnpv", 0 /* Default */, [rate, values, dates], false, true, null));
        };
        Functions.prototype.xor = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Xor", 0 /* Default */, [values], false, true, null));
        };
        Functions.prototype.year = function (serialNumber) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Year", 0 /* Default */, [serialNumber], false, true, null));
        };
        Functions.prototype.yearFrac = function (startDate, endDate, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "YearFrac", 0 /* Default */, [startDate, endDate, basis], false, true, null));
        };
        Functions.prototype.yield = function (settlement, maturity, rate, pr, redemption, frequency, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Yield", 0 /* Default */, [settlement, maturity, rate, pr, redemption, frequency, basis], false, true, null));
        };
        Functions.prototype.yieldDisc = function (settlement, maturity, pr, redemption, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "YieldDisc", 0 /* Default */, [settlement, maturity, pr, redemption, basis], false, true, null));
        };
        Functions.prototype.yieldMat = function (settlement, maturity, issue, rate, pr, basis) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "YieldMat", 0 /* Default */, [settlement, maturity, issue, rate, pr, basis], false, true, null));
        };
        Functions.prototype.z_Test = function (array, x, sigma) {
            return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Z_Test", 0 /* Default */, [array, x, sigma], false, true, null));
        };
        Functions.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        return Functions;
    })(OfficeExtension.ClientObject);
    Excel.Functions = Functions;
    var ErrorCodes;
    (function (ErrorCodes) {
        ErrorCodes.accessDenied = "AccessDenied";
        ErrorCodes.apiNotFound = "ApiNotFound";
        ErrorCodes.generalException = "GeneralException";
        ErrorCodes.insertDeleteConflict = "InsertDeleteConflict";
        ErrorCodes.invalidArgument = "InvalidArgument";
        ErrorCodes.invalidBinding = "InvalidBinding";
        ErrorCodes.invalidOperation = "InvalidOperation";
        ErrorCodes.invalidReference = "InvalidReference";
        ErrorCodes.invalidSelection = "InvalidSelection";
        ErrorCodes.itemAlreadyExists = "ItemAlreadyExists";
        ErrorCodes.itemNotFound = "ItemNotFound";
        ErrorCodes.notImplemented = "NotImplemented";
        ErrorCodes.unsupportedOperation = "UnsupportedOperation";
    })(ErrorCodes = Excel.ErrorCodes || (Excel.ErrorCodes = {}));
})(Excel || (Excel = {}));
var Excel;
(function (Excel) {
    var RequestContext = (function (_super) {
        __extends(RequestContext, _super);
        function RequestContext(url) {
            _super.call(this, url);
            this.m_workbook = new Excel.Workbook(this, OfficeExtension.ObjectPathFactory.createGlobalObjectObjectPath(this));
            this._rootObject = this.m_workbook;
        }
        Object.defineProperty(RequestContext.prototype, "workbook", {
            get: function () {
                return this.m_workbook;
            },
            enumerable: true,
            configurable: true
        });
        return RequestContext;
    })(OfficeExtension.ClientRequestContext);
    Excel.RequestContext = RequestContext;
    function run(arg1, arg2) {
        return OfficeExtension.ClientRequestContext._runBatch("Excel.run", arguments, function () { return new Excel.RequestContext(); });
    }
    Excel.run = run;
})(Excel || (Excel = {}));
var Excel;
(function (Excel) {
    Excel._RedirectV1APIs = false;
    Excel._V1APIMap = {
        "GetDataAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingGetData(callArgs); },
            postprocess: getDataCommonPostprocess
        },
        "GetSelectedDataAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.getSelectedData(callArgs); },
            postprocess: getDataCommonPostprocess
        },
        "GoToByIdAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.gotoById(callArgs); }
        },
        "AddColumnsAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingAddColumns(callArgs); }
        },
        "AddFromSelectionAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingAddFromSelection(callArgs); },
            postprocess: postprocessBindingDescriptor
        },
        "AddFromNamedItemAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingAddFromNamedItem(callArgs); },
            postprocess: postprocessBindingDescriptor
        },
        "AddFromPromptAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingAddFromPrompt(callArgs); },
            postprocess: postprocessBindingDescriptor
        },
        "AddRowsAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingAddRows(callArgs); }
        },
        "GetByIdAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingGetById(callArgs); },
            postprocess: postprocessBindingDescriptor
        },
        "ReleaseByIdAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingReleaseById(callArgs); }
        },
        "GetAllAsync": {
            call: function (ctx) { return ctx.workbook._V1Api.bindingGetAll(); },
            postprocess: function (response) {
                return response.bindings.map(function (descriptor) { return postprocessBindingDescriptor(descriptor); });
            }
        },
        "DeleteAllDataValuesAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingDeleteAllDataValues(callArgs); }
        },
        "SetSelectedDataAsync": {
            preprocess: function (callArgs) {
                var preimage = callArgs["cellFormat"];
                if (window.OSF.DDA.SafeArray) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
                    }
                }
                else if (window.OSF.DDA.WAC) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
                    }
                }
                return callArgs;
            },
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.setSelectedData(callArgs); }
        },
        "SetDataAsync": {
            preprocess: function (callArgs) {
                var preimage = callArgs["cellFormat"];
                if (window.OSF.DDA.SafeArray) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
                    }
                }
                else if (window.OSF.DDA.WAC) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
                    }
                }
                return callArgs;
            },
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingSetData(callArgs); }
        },
        "SetFormatsAsync": {
            preprocess: function (callArgs) {
                var preimage = callArgs["cellFormat"];
                if (window.OSF.DDA.SafeArray) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
                    }
                }
                else if (window.OSF.DDA.WAC) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
                    }
                }
                return callArgs;
            },
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingSetFormats(callArgs); }
        },
        "SetTableOptionsAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingSetTableOptions(callArgs); }
        },
        "ClearFormatsAsync": {
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingClearFormats(callArgs); }
        },
    };
    function postprocessBindingDescriptor(response) {
        var bindingDescriptor = {
            BindingColumnCount: response.bindingColumnCount,
            BindingId: response.bindingId,
            BindingRowCount: response.bindingRowCount,
            bindingType: response.bindingType,
            HasHeaders: response.hasHeaders
        };
        return window.OSF.DDA.OMFactory.manufactureBinding(bindingDescriptor, window.Microsoft.Office.WebExtension.context.document);
    }
    function getDataCommonPostprocess(response, callArgs) {
        var isPlainData = response.headers == null;
        var data;
        if (isPlainData) {
            data = response.rows;
        }
        else {
            data = response;
        }
        data = window.OSF.DDA.DataCoercion.coerceData(data, callArgs[window.Microsoft.Office.WebExtension.Parameters.CoercionType]);
        return data == undefined ? null : data;
    }
})(Excel || (Excel = {}));
