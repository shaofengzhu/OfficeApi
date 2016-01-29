// --------------------------------------------------------------------------------------------------
// 
// <copyright file="ExcelApi.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
// <summary>
// Contains the metadata of Excel API that is currently implemented.
// The following is the workflow to add a new API
// 1) DEV add the API to xlshared\src\api\metadata\current\ExcelApi.cs
// 2) DEV runs xlshared\util\XlsApiGen.bat to re-generate the following files
//      xlshared\src\api\Xlapi.h                COM CoClass header file
//      xlshared\src\api\Xlapi_i.h              COM interface header file
//      xlshared\src\api\Xlapi_i.cpp            COM GUIDs
//      xlshared\src\api\TypeRegistration.cpp   Type registration file
//      xlshared\src\api\*.disp.cpp             COM IDispatch interface related implementation
//      xlshared\src\api\script\Xlapi.ts        TypeScript file
// 3) DEV implement the new API, update xlshared\src\api\sources.inc if necessary.
// </summary>
// --------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using Microsoft.OfficeExtension.CodeGen;

[assembly: ClientCallableNamespaceMap("Microsoft.ExcelServices", ComCoClassNamespaceName = "ExcelApiImpl", ComInterfaceNamespaceName = "ExcelApi", TypeScriptNamespaceName = "Excel")]

// Default error (fallback if not uniquely mapped below)
[assembly: HResultDefaultError(HttpStatusCode.InternalServerError, Microsoft.ExcelServices.ErrorCodes.GeneralException, "stridsApiGeneralException")]

// Errors we specifically want to hide into general exception (500)
[assembly: HResultError("hrFail", HttpStatusCode.InternalServerError, Microsoft.ExcelServices.ErrorCodes.GeneralException, "stridsApiGeneralException")]
[assembly: HResultError("hrOutOfMemory", HttpStatusCode.InternalServerError, Microsoft.ExcelServices.ErrorCodes.GeneralException, "stridsApiGeneralException")]
[assembly: HResultError("SharedInterimIfs::hrFormulaParseError", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsApiInvalidArgument")]

// Errors 400s
[assembly: HResultError("E_POINTER", HttpStatusCode.NotFound, Microsoft.ExcelServices.ErrorCodes.ItemNotFound, "stridsApiItemNotFound")]
[assembly: HResultError("hrBadIndex", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsWCOUTOFBOUNDS")]
[assembly: HResultError("hrInvalidArg", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsApiInvalidArgument")]
[assembly: HResultError("hrInvalidAPIOperation", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidOperation, "stridsApiInvalidAPIOperation")]
[assembly: HResultError("hrInvalidBinding", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidBinding, "stridsApiInvalidBinding")]
[assembly: HResultError("hrInvalidAPISelection", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidSelection, "stridsApiInvalidSelection")]
[assembly: HResultError("hrInvalidAPIReference", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidReference, "stridsApiInvalidReference")]
[assembly: HResultError("hrNotFound", HttpStatusCode.NotFound, Microsoft.ExcelServices.ErrorCodes.ItemNotFound, "stridsApiItemNotFound")]
[assembly: HResultError("SharedInterimIfs::hrInsDelDisallowedByFeature", HttpStatusCode.Conflict, Microsoft.ExcelServices.ErrorCodes.InsertDeleteConflict, "stridsBadListInsDel")]
[assembly: HResultError("hrListCannotGrow", HttpStatusCode.Conflict, Microsoft.ExcelServices.ErrorCodes.InsertDeleteConflict, "stridsBadListInsDel")]
[assembly: HResultError("hrNotYetSupportedApiOperation", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.UnsupportedOperation, "stridsApiNotImplemented")]
[assembly: HResultError("SharedInterimIfs::hrRangeSheetsMismatch", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsApiInvalidArgument")]
[assembly: HResultError("SharedInterimIfs::hrRangeParseError", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsApiInvalidArgument")]
[assembly: HResultError("SharedInterimIfs::hrRangeWrong", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsApiInvalidArgument")]
[assembly: HResultError("hrNoPermission", HttpStatusCode.Forbidden, Microsoft.ExcelServices.ErrorCodes.AccessDenied, "stridsApiAccessDenied")]
[assembly: HResultError("E_ACCESSDENIED", HttpStatusCode.Forbidden, Microsoft.ExcelServices.ErrorCodes.AccessDenied, "stridsApiAccessDenied")]
[assembly: HResultError("SharedInterimIfs::hrCreateTableBadListSrcRange", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsBadListPasteSrcRange")]
[assembly: HResultError("SharedInterimIfs::hrGetTableBadListSrcRange", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsBadListSrcRange")]
[assembly: HResultError("SharedInterimIfs::hrCreateTableFormulaInListHdr", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsFormulaInListHdr")]
[assembly: HResultError("SharedInterimIfs::hrCreateTableColHdrTruncate", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsTableColHdrTruncate")]
[assembly: HResultError("SharedInterimIfs::hrGetTableListsOverlap", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsListsOverlap")]
[assembly: HResultError("hrItemAlreadyExists", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.ItemAlreadyExists, "stridsApiItemAlreadyExists")]
[assembly: HResultError("hrNoInterface", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsApiInvalidArgument")]
[assembly: HResultError("hrApiBadPassword", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.BadPassword, "stridsApiBadPassword")]

// Errors 500s
[assembly: HResultError("hrNotImplemented", HttpStatusCode.NotImplemented, Microsoft.ExcelServices.ErrorCodes.NotImplemented, "stridsApiNotImplemented")]
//[assembly: HResultError("hrAborted", HttpStatusCode.InternalServerError, Microsoft.ExcelServices.ErrorCodes.RequestAborted, "stridsApiAborted")] (hrAborted is not yet implemented)

namespace Microsoft.ExcelServices
{

	internal static class ErrorCodes
	{
		internal const string GeneralException = "GeneralException";
		internal const string InvalidArgument = "InvalidArgument";
		internal const string InvalidOperation = "InvalidOperation";
		internal const string InvalidSelection = "InvalidSelection";
		internal const string InvalidBinding = "InvalidBinding";
		internal const string InsertDeleteConflict = "InsertDeleteConflict";
		internal const string ItemNotFound = "ItemNotFound";
		internal const string NotImplemented = "NotImplemented";
		internal const string InvalidReference = "InvalidReference";
		internal const string InvalidRequest = "InvalidRequest";
		internal const string ApiNotAvailable = "ApiNotAvailable";
		internal const string Unauthenticated = "Unauthenticated";
		internal const string AccessDenied = "AccessDenied";
		internal const string Conflict = "Conflict";
		internal const string ItemAlreadyExists = "ItemAlreadyExists";
		internal const string ContentLengthRequired = "ContentLengthRequired";
		internal const string ActivityLimitReached = "ActivityLimitReached";
		internal const string RequestAborted = "RequestAborted";
		internal const string ServiceNotAvailable = "ServiceNotAvailable";
		internal const string UnsupportedOperation = "UnsupportedOperation";
		internal const string BadPassword = "BadPassword";
	}

	/// <summary>
	/// Dispatch Ids
	/// </summary>
	/// <remarks>
	/// Please keep them ordered and grouped by type name alphabetically, then ordered by the value of dispatch id.
	/// </remarks>
	internal static class DispatchIds
	{
		internal const int Application_CalculationMode = 1;
		internal const int Application_Calculate = 2;

		internal const int Binding_Id = 1;
		internal const int Binding_Type = 2;
		internal const int Binding_Table = 3;
		internal const int Binding_Range = 4;
		internal const int Binding_Text = 5;
		internal const int Binding_OnAccess = 6;

		internal const int BindingCollection_Indexer = 1;
		internal const int BindingCollection_Count = 2;
		internal const int BindingCollection_ItemAt = 3;

		internal const int ChartAxes_Category = 1;
		internal const int ChartAxes_Series = 2;
		internal const int ChartAxes_Value = 3;
		internal const int ChartAxes_OnAccess = 4;

		internal const int ChartAxis_MajorGridlines = 1;
		internal const int ChartAxis_MajorUnit = 2;
		internal const int ChartAxis_Maximum = 3;
		internal const int ChartAxis_Minimum = 4;
		internal const int ChartAxis_MinorGridlines = 5;
		internal const int ChartAxis_MinorUnit = 6;
		internal const int ChartAxis_Title = 7;
		internal const int ChartAxis_Format = 8;
		internal const int ChartAxis_OnAccess = 9;

		internal const int ChartAxisFormat_Font = 1;
		internal const int ChartAxisFormat_Line = 2;
		internal const int ChartAxisFormat_OnAccess = 3;

		internal const int ChartAxisTitle_Text = 1;
		internal const int ChartAxisTitle_Visible = 2;
		internal const int ChartAxisTitle_Format = 3;
		internal const int ChartAxisTitle_OnAccess = 4;

		internal const int ChartAxisTitleFormat_Font = 1;
		internal const int ChartAxisTitleFormat_OnAccess = 2;

		internal const int Chart_Title = 1;
		internal const int Chart_SetData = 2;
		internal const int Chart_DataLabels = 3;
		internal const int Chart_Legend = 4;
		internal const int Chart_Name = 5;
		internal const int Chart_Top = 6;
		internal const int Chart_Left = 7;
		internal const int Chart_Width = 8;
		internal const int Chart_Height = 9;
		internal const int Chart_Delete = 10;
		internal const int Chart_Series = 11;
		internal const int Chart_Id = 12;
		internal const int Chart_Axes = 13;
		internal const int Chart_Format = 14;
		internal const int Chart_OnAccess = 15;
		internal const int Chart_SetPosition = 16;
		internal const int Chart_GetImage = 17;
		internal const int Chart_Worksheet = 18;

		internal const int ChartAreaFormat_Fill = 1;
		internal const int ChartAreaFormat_Font = 2;
		internal const int ChartAreaFormat_OnAccess = 3;

		internal const int ChartCollection_Add = 1;
		internal const int ChartCollection_Count = 2;
		internal const int ChartCollection_ItemAt = 3;
		internal const int ChartCollection_Indexer = 4;
		internal const int ChartCollection_GetByName = 5;
		internal const int ChartCollection_GetItem = 6;
		internal const int ChartCollection_OnAccess = 7;

		internal const int ChartDataLabels_Position = 1;
		internal const int ChartDataLabels_ShowValue = 2;
		internal const int ChartDataLabels_ShowSeriesName = 3;
		internal const int ChartDataLabels_ShowCategoryName = 4;
		internal const int ChartDataLabels_ShowLegendKey = 5;
		internal const int ChartDataLabels_ShowPercentage = 6;
		internal const int ChartDataLabels_ShowBubbleSize = 7;
		internal const int ChartDataLabels_Separator = 8;
		internal const int ChartDataLabels_Format = 9;
		internal const int ChartDataLabels_OnAccess = 10;

		internal const int ChartDataLabelFormat_Font = 1;
		internal const int ChartDataLabelFormat_Fill = 2;
		internal const int ChartDataLabelFormat_OnAccess = 3;

		internal const int ChartFill_SolidColor = 1;
		internal const int ChartFill_Clear = 2;
		internal const int ChartFill_OnAccess = 3;

		internal const int ChartFont_Bold = 1;
		internal const int ChartFont_Color = 2;
		internal const int ChartFont_Italic = 3;
		internal const int ChartFont_Name = 4;
		internal const int ChartFont_Size = 5;
		internal const int ChartFont_Underline = 6;
		internal const int ChartFont_OnAccess = 7;

		internal const int ChartGridlines_Visible = 1;
		internal const int ChartGridlines_Format = 2;
		internal const int ChartGridlines_OnAccess = 3;

		internal const int ChartGridlinesFormat_Line = 1;
		internal const int ChartGridlinesFormat_OnAccess = 2;

		internal const int ChartLegend_Visible = 1;
		internal const int ChartLegend_Position = 2;
		internal const int ChartLegend_Overlay = 3;
		internal const int ChartLegend_Format = 4;
		internal const int ChartLegend_OnAccess = 5;

		internal const int ChartLegendFormat_Font = 1;
		internal const int ChartLegendFormat_Fill = 2;
		internal const int ChartLegendFormat_OnAccess = 3;

		internal const int ChartLineFormat_Clear = 1;
		internal const int ChartLineFormat_Color = 2;
		internal const int ChartLineFormat_OnAccess = 3;

		internal const int ChartTitle_Visible = 1;
		internal const int ChartTitle_Text = 2;
		internal const int ChartTitle_Overlay = 3;
		internal const int ChartTitle_Format = 4;
		internal const int ChartTitle_OnAccess = 5;

		internal const int ChartTitleFormat_Font = 1;
		internal const int ChartTitleFormat_Fill = 2;
		internal const int ChartTitleFormat_OnAccess = 3;

		internal const int ChartPoint_Format = 1;
		internal const int ChartPoint_Value = 2;
		internal const int ChartPoint_OnAccess = 3;

		internal const int ChartPointFormat_Fill = 1;
		internal const int ChartPointFormat_OnAccess = 2;

		internal const int ChartPointsCollection_Count = 1;
		internal const int ChartPointsCollection_ItemAt = 2;
		internal const int ChartPointsCollection_OnAccess = 3;

		internal const int ChartSeries_Name = 1;
		internal const int ChartSeries_Points = 2;
		internal const int ChartSeries_Format = 3;
		internal const int ChartSeries_OnAccess = 4;

		internal const int ChartSeriesFormat_Fill = 1;
		internal const int ChartSeriesFormat_Line = 2;
		internal const int ChartSeriesFormat_OnAccess = 3;

		internal const int ChartSeriesCollection_Count = 1;
		internal const int ChartSeriesCollection_ItemAt = 2;
		internal const int ChartSeriesCollection_OnAccess = 3;

		internal const int FormatProtection_OnAccess = 1;
		internal const int FormatProtection_Locked = 2;
		internal const int FormatProtection_FormulaHidden = 3;

		internal const int Filter_Apply = 1;
		internal const int Filter_OnAccess = 2;
		internal const int Filter_Clear = 3;
		internal const int Filter_Criteria = 4;
		internal const int Filter_BottomItems = 5;
		internal const int Filter_BottomPercent = 6;
		internal const int Filter_CellColor = 7;
		internal const int Filter_Dynamic = 8;
		internal const int Filter_FontColor = 9;
		internal const int Filter_Values = 10;
		internal const int Filter_TopItems = 11;
		internal const int Filter_TopPercent = 12;
		internal const int Filter_Icon = 13;
		internal const int Filter_Custom = 14;

		internal const int FilterCriteria_Criterion1 = 1;
		internal const int FilterCriteria_Criterion2 = 2;
		internal const int FilterCriteria_Color = 3;
		internal const int FilterCriteria_Operator = 4;
		internal const int FilterCriteria_Icon = 5;
		internal const int FilterCriteria_DynamicCriteria = 6;
		internal const int FilterCriteria_Values = 7;
		internal const int FilterCriteria_FilterOn = 8;

		internal const int FilterDatetime_Date = 1;
		internal const int FilterDatetime_Specificity = 2;
		internal const int FunctionResult_Error = 1;
		internal const int FunctionResult_Value = 2;

		internal const int Icon_Set = 1;
		internal const int Icon_Index = 2;

		internal const int NamedItem_Name = 1;
		internal const int NamedItem_Type = 2;
		internal const int NamedItem_Value = 3;
		internal const int NamedItem_Range = 4;
		internal const int NamedItem_Visible = 5;
		internal const int NamedItem_Id = 6;
		internal const int NamedItem_OnAccess = 7;

		internal const int NamedItemCollection_Indexer = 1;

		internal const int Range_NumberFormat = 1;	// DO NOT CHANGE Order of NumberFormat and Values
		internal const int Range_Values = 2;		// DO NOT CHANGE Order of NumberFormat and Values
		internal const int Range_Text = 3;
		internal const int Range_Formulas = 4;
		internal const int Range_FormulasLocal = 5;
		internal const int Range_RowIndex = 6;
		internal const int Range_ColumnIndex = 7;
		internal const int Range_RowCount = 8;
		internal const int Range_ColumnCount = 9;
		internal const int Range_Format = 10;
		internal const int Range_Address = 11;
		internal const int Range_AddressLocal = 12;
		internal const int Range_Cell = 13;
		internal const int Range_CellCount = 14;
		internal const int Range_UsedRange = 15;
		internal const int Range_Clear = 16;
		internal const int Range_Insert = 17;
		internal const int Range_Delete = 18;
		internal const int Range_EntireColumn = 19;
		internal const int Range_EntireRow = 20;
		internal const int Range_Worksheet = 21;
		internal const int Range_Select = 22;
		internal const int Range_ReferenceId = 23;
		internal const int Range_KeepReference = 24;
		internal const int Range_GetOffsetRange = 25;
		internal const int Range_GetRow = 26;
		internal const int Range_GetColumn = 27;
		internal const int Range_OnAccess = 28;
		internal const int Range_GetIntersection = 29;
		internal const int Range_GetBoundingRect = 30;
		internal const int Range_ValueTypes = 31;
		internal const int Range_GetLastCell = 32;
		internal const int Range_GetLastColumn = 33;
		internal const int Range_GetLastRow = 34;
		internal const int Range_FormulasR1C1 = 35;
		internal const int Range_Sort = 36;
		internal const int Range_Merge = 37;
		internal const int Range_Unmerge = 38;
		internal const int Range_Hidden = 39;
		internal const int Range_RowHidden = 40;
		internal const int Range_ColumnHidden = 41;

		internal const int RangeBorder_SideIndex = 1;
		internal const int RangeBorder_LineStyle = 2;
		internal const int RangeBorder_Weight = 3;
		internal const int RangeBorder_Color = 4;
		internal const int RangeBorder_OnAccess = 5;
		internal const int RangeBorder_Id = 6;

		internal const int RangeBorderCollection_Indexer = 1;
		internal const int RangeBorderCollection_Count = 2;
		internal const int RangeBorderCollection_ItemAt = 3;
		internal const int RangeBorderCollection_OnAccess = 4;

		internal const int RangeFill_Color = 1;
		internal const int RangeFill_Clear = 2;
		internal const int RangeFill_OnAccess = 3;

		internal const int RangeFont_Name = 1;
		internal const int RangeFont_Size = 2;
		internal const int RangeFont_Color = 3;
		internal const int RangeFont_Italic = 4;
		internal const int RangeFont_Bold = 5;
		internal const int RangeFont_Underline = 6;
		internal const int RangeFont_OnAccess = 7;

		internal const int RangeFormat_Fill = 1;
		internal const int RangeFormat_Font = 2;
		internal const int RangeFormat_WrapText = 3;
		internal const int RangeFormat_HorizontalAlignment = 4;
		internal const int RangeFormat_VerticalAlignment = 5;
		internal const int RangeFormat_Borders = 6;
		internal const int RangeFormat_OnAccess = 7;
		internal const int RangeFormat_ColumnWidth = 8;
		internal const int RangeFormat_RowHeight = 9;
		internal const int RangeFormat_AutofitColumns = 10;
		internal const int RangeFormat_AutofitRows = 11;
		internal const int RangeFormat_Protection = 12;

		internal const int RangeReference_Address = 1;

		internal const int RangeSort_Apply = 1;
		internal const int RangeSort_OnAccess = 2;

		internal const int SortField_Key = 1;
		internal const int SortField_SortOn = 2;
		internal const int SortField_Ascending = 3;
		internal const int SortField_Color = 4;
		internal const int SortField_DataOption = 5;
		internal const int SortField_Icon = 6;

		internal const int Table_Id = 1;
		internal const int Table_Name = 2;
		internal const int Table_Range = 3;
		internal const int Table_HeaderRowRange = 4;
		internal const int Table_DataBodyRange = 5;
		internal const int Table_TotalRowRange = 6;
		internal const int Table_ShowHeaders = 7;
		internal const int Table_ShowTotals = 8;
		internal const int Table_TableStyle = 9;
		internal const int Table_TableColumns = 10;
		internal const int Table_TableRows = 11;
		internal const int Table_Delete = 12;
		internal const int Table_OnAccess = 13;
		internal const int Table_Sort = 14;
		internal const int Table_ConvertToRange = 15;
		internal const int Table_Worksheet = 16;
		internal const int Table_ClearFilters = 17;
		internal const int Table_ReapplyFilters = 18;

		internal const int TableCollection_Count = 1;
		internal const int TableCollection_Indexer = 2;
		internal const int TableCollection_ItemAt = 3;
		internal const int TableCollection_Add = 4;
		internal const int TableCollection_OnAccess = 5;

		internal const int TableColumn_Id = 1;
		internal const int TableColumn_Name = 2;
		internal const int TableColumn_Index = 3;
		internal const int TableColumn_Range = 4;
		internal const int TableColumn_HeaderRowRange = 5;
		internal const int TableColumn_DataBodyRange = 6;
		internal const int TableColumn_TotalRowRange = 7;
		internal const int TableColumn_Values = 8;
		internal const int TableColumn_Delete = 9;
		internal const int TableColumn_OnAccess = 10;
		internal const int TableColumn_Filter = 11;

		internal const int TableColumnCollection_Count = 1;
		internal const int TableColumnCollection_Indexer = 2;
		internal const int TableColumnCollection_ItemAt = 3;
		internal const int TableColumnCollection_Insert = 4;
		internal const int TableColumnCollection_OnAccess = 5;

		internal const int TableRow_Index = 1;
		internal const int TableRow_Range = 2;
		internal const int TableRow_Values = 3;
		internal const int TableRow_Delete = 4;
		internal const int TableRow_OnAccess = 5;

		internal const int TableRowCollection_Count = 1;
		internal const int TableRowCollection_ItemAt = 2;
		internal const int TableRowCollection_Insert = 3;
		internal const int TableRowCollection_OnAccess = 4;

		internal const int TableSort_Apply = 1;
		internal const int TableSort_MatchCase = 2;
		internal const int TableSort_Method = 3;
		internal const int TableSort_OnAccess = 4;
		internal const int TableSort_Clear = 5;
		internal const int TableSort_Reapply = 6;
		internal const int TableSort_Fields = 7;

		internal const int Workbook_Worksheets = 1;
		internal const int Workbook_Names = 2;
		internal const int Workbook_Tables = 3;
		internal const int Workbook_Application = 4;
		internal const int Workbook_SelectedRange = 5;
		internal const int Workbook_Bindings = 6;
		internal const int Workbook_RemoveReference = 7;
		internal const int Workbook_GetObjectByReferenceId = 8;
		internal const int Workbook_GetObjectTypeNameByReferenceId = 9;
		internal const int Workbook_RemoveAllReferences = 10;
		internal const int Workbook_GetReferenceCount = 11;
		internal const int Workbook_Functions = 12;

		internal const int Worksheet_Range = 1;
		internal const int Worksheet_UsedRange = 2;
		internal const int Worksheet_Charts = 3;
		internal const int Worksheet_Cell = 4;
		internal const int Worksheet_Name = 5;
		internal const int Worksheet_Delete = 6;
		internal const int Worksheet_Id = 7;
		internal const int Worksheet_Tables = 8;
		internal const int Worksheet_Activate = 9;
		internal const int Worksheet_Position = 10;
		internal const int Worksheet_OnAccess = 11;
		internal const int Worksheet_Visible = 12;
		internal const int Worksheet_Protection = 13;

		internal const int WorksheetCollection_Indexer = 1;
		internal const int WorksheetCollection_Add = 2;
		internal const int WorksheetCollection_ActiveWorksheet = 3;

		internal const int WorksheetProtection_OnAccess = 1;
		internal const int WorksheetProtection_Protected = 2;
		internal const int WorksheetProtection_Options = 3;
		internal const int WorksheetProtection_Protect = 4;
		internal const int WorksheetProtection_Unprotect = 5;

		internal const int WorksheetProtectionOptions_AllowFormatCells = 1;
		internal const int WorksheetProtectionOptions_AllowFormatColumns = 2;
		internal const int WorksheetProtectionOptions_AllowFormatRows = 3;
		internal const int WorksheetProtectionOptions_AllowInsertColumns = 4;
		internal const int WorksheetProtectionOptions_AllowInsertRows = 5;
		internal const int WorksheetProtectionOptions_AllowInsertHyperlinks = 6;
		internal const int WorksheetProtectionOptions_AllowDeleteColumns = 7;
		internal const int WorksheetProtectionOptions_AllowDeleteRows = 8;
		internal const int WorksheetProtectionOptions_AllowSort = 9;
		internal const int WorksheetProtectionOptions_AllowAutoFilter = 10;
		internal const int WorksheetProtectionOptions_AllowPivotTables = 11;
	}

#region Application
	/// <summary>
	/// Represents the Excel application that manages the workbook.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IApplication", InterfaceId = "053AAB3F-C5B6-4A91-93A5-A2C4DA223516", CoClassName = "Application")]
	public interface Application
	{
		/// <summary>
		/// Returns the calculation mode used in the workbook. See Excel.CalculationMode for details. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Application_CalculationMode)]
		CalculationMode CalculationMode { get; }

		/// <summary>
		/// Recalculate all currently opened workbooks in Excel.
		/// </summary>
		/// <param name="calculationType">Specifies the calculation type to use. See Excel.CalculationType for details.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Application_Calculate)]
		void Calculate(CalculationType calculationType);
	}
#endregion Application

#region Workbook
	/// <summary>
	/// Workbook is the top level object which contains related workbook objects such as worksheets, tables, ranges, etc.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IWorkbook", InterfaceId = "bb02266c-6204-4e0d-baa3-cc1a928f573e", CoClassName = "Workbook")]
	[ClientCallableServiceRoot]
	public interface Workbook
	{
		/// <summary>
		/// Gets the currently selected range from the workbook.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Workbook_SelectedRange)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetSelectedRange();

		[ClientCallableComMember(DispatchId = DispatchIds.Workbook_GetObjectByReferenceId)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		object _GetObjectByReferenceId(string bstrReferenceId);

		[ClientCallableComMember(DispatchId = DispatchIds.Workbook_GetObjectTypeNameByReferenceId)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		string _GetObjectTypeNameByReferenceId(string bstrReferenceId);

		[ClientCallableComMember(DispatchId = DispatchIds.Workbook_GetReferenceCount)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		int _GetReferenceCount();

		[ClientCallableComMember(DispatchId = DispatchIds.Workbook_RemoveAllReferences)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _RemoveAllReferences();

		[ClientCallableComMember(DispatchId = DispatchIds.Workbook_RemoveReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _RemoveReference(string bstrReferenceId);

		/// <summary>
		/// Represents Excel application instance that contains this workbook. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Workbook_Application)]
		Application Application { get; }

		/// <summary>
		/// Represents Excel application instance that contains this workbook. Read-only.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Workbook_Functions)]
		Functions Functions { get; }

		/// <summary>
		/// Represents a collection of workbook scoped named items (named ranges and constants). Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Workbook_Names)]
		NamedItemCollection Names { get; }

		/// <summary>
		/// Represents a collection of worksheets associated with the workbook. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Workbook_Worksheets)]
		WorksheetCollection Worksheets { get; }

		/// <summary>
		/// Represents a collection of tables associated with the workbook. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Workbook_Tables)]
		TableCollection Tables { get; }

		/// <summary>
		/// Represents a collection of bindings that are part of the workbook. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Workbook_Bindings)]
		[ClientCallableProperty(ExcludedFromRest = true)]
		BindingCollection Bindings { get; }
	}
#endregion Workbook

#region Worksheet
	/// <summary>
	/// An Excel worksheet is a grid of cells. It can contain data, tables, charts, etc.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableType(DeleteOperationName = "Delete")]
	[ClientCallableComType(Name = "IWorksheet", InterfaceId = "b86e5ae1-476e-4e56-825d-885468e549f3", CoClassName = "Worksheet")]
	public interface Worksheet
	{
		[ClientCallableComMember(DispatchId = DispatchIds.Worksheet_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Activate the worksheet in the Excel UI.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Worksheet_Activate)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "")]
		void Activate();
		/// <summary>
		/// Returns collection of charts that are part of the worksheet. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Worksheet_Charts)]
		ChartCollection Charts { get; }
		/// <summary>
		/// Deletes the worksheet from the workbook.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Worksheet_Delete)]
		void Delete();
		/// <summary>
		/// Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it's stays within the worksheet grid.
		/// </summary>
		/// <param name="row">The row number of the cell to be retrieved. Zero-indexed.</param>
		/// <param name="column">the column number of the cell to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Worksheet_Cell)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "Cell", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetCell(int row, int column);
		/// <summary>
		/// Gets the range object specified by the address or name.
		/// </summary>
		/// <param name="address">The address or the name of the range. If not specified, the entire worksheet range is returned.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Worksheet_Range)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "Range", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetRange([Optional]string address);
		/// <summary>
		/// Returns a value that uniquely identifies the worksheet in a given workbook. The value of the identifier remains the same even when the worksheet is renamed or moved. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Worksheet_Id)]
		string Id { get; }
		/// <summary>
		/// The zero-based position of the worksheet within the workbook.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Worksheet_Position)]
		int Position { get; set; }
		/// <summary>
		/// The display name of the worksheet.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Worksheet_Name)]
		string Name { get; set; }
		/// <summary>
		/// The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the worksheet is blank, this function will return the top left cell.
		/// </summary>
		/// <param name="valuesOnly">Considers only cells with values as used cells (ignores formatting).</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Worksheet_UsedRange)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "UsedRange", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetUsedRange([ApiSet(Version = 1.2)][Optional]bool valuesOnly);
		/// <summary>
		/// Collection of tables that are part of the worksheet. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Worksheet_Tables)]
		TableCollection Tables { get; }
		/// <summary>
		/// The Visibility of the worksheet.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Worksheet_Visible)]
		[ApiSet(Version = 1.1, CustomText = "1.1 for reading visibility; 1.2 for setting it.")]
		SheetVisibility Visibility { get; set; }
		/// <summary>
		/// Returns sheet protection object for a worksheet.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Worksheet_Protection)]
		WorksheetProtection Protection { get; }
	}

	/// <summary>
	/// Represents a collection of worksheet objects that are part of the workbook.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableType(CreateItemOperationName = "Add")]
	[ClientCallableComType(Name = "IWorksheetCollection", InterfaceId = "55a36c77-3310-4afb-aa64-3c1a685f2f50", CoClassName = "WorksheetCollection", SupportEnumeration = true)]
	public interface WorksheetCollection : IEnumerable<Worksheet>
	{
		/// <summary>
		/// Gets the currently active worksheet in the workbook.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.WorksheetCollection_ActiveWorksheet)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "")]
		Worksheet GetActiveWorksheet();
		/// <summary>
		/// Gets a worksheet object using its Name or ID.
		/// </summary>
		/// <param name="key">The Name or ID of the worksheet.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.WorksheetCollection_Indexer)]
		Worksheet this[string key] { get; }
		/// <summary>
		/// Adds a new worksheet to the workbook. The worksheet will be added at the end of existing worksheets. If you wish to activate the newly added worksheet, call ".activate() on it.
		/// </summary>
		/// <param name="name">The name of the worksheet to be added. If specified, name should be unqiue. If not specified, Excel determines the name of the new worksheet.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.WorksheetCollection_Add)]
		Worksheet Add([Optional]string name);
	}

	/// <summary>
	/// Represents the protection of a sheet object.
	/// </summary>
	[ApiSet(Version = 1.2)]
	[ClientCallableComType(Name = "IWorksheetProtection", InterfaceId = "C84C0D35-DEDB-4865-B4A0-B027BAFEC20D", CoClassName = "WorksheetProtection")]
	public interface WorksheetProtection
	{
		[ClientCallableComMember(DispatchId = DispatchIds.WorksheetProtection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Indicates if the worksheet is protected. Read-Only.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.WorksheetProtection_Protected)]
		bool Protected { get; }
		/// <summary>
		/// Sheet protection options.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.WorksheetProtection_Options)]
		WorksheetProtectionOptions Options { get; }
		/// <summary>
		/// Protect a worksheet. It throws if the worksheet has been protected.
		/// </summary>
		/// <param name="options">sheet protection options.</param>
		/// <param name="password">sheet protection password.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.WorksheetProtection_Protect)]
		void Protect([Optional]WorksheetProtectionOptions options, [Optional]string password);
		/// <summary>
		/// Unprotect a worksheet
		/// </summary>
		/// <param name="password">sheet protection password.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.WorksheetProtection_Unprotect)]
		void Unprotect([Optional]string password);
	}

	/// <summary>
	/// Represents the options in sheet protection.
	/// </summary>
	[ApiSet(Version = 1.2)]
	[ClientCallableComType(Name = "IWorksheetProtectionOptions", InterfaceId = "201D75BE-81F5-4B2A-A3A8-AE4E72E47ECB", CoClassName = "WorksheetProtectionOptions", CoClassId = "56C94DB3-B781-44CF-9CA8-29FB47A6A267")]
	public struct WorksheetProtectionOptions
	{
		/// <summary>
		/// Represents the worksheet protection option of allowing formatting cells.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.WorksheetProtectionOptions_AllowFormatCells)]
		[Optional]
		bool AllowFormatCells { get; set; }
		/// <summary>
		/// Represents the worksheet protection option of allowing formatting columns.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.WorksheetProtectionOptions_AllowFormatColumns)]
		[Optional]
		bool AllowFormatColumns { get; set; }
		/// <summary>
		/// Represents the worksheet protection option of allowing formatting rows.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.WorksheetProtectionOptions_AllowFormatRows)]
		[Optional]
		bool AllowFormatRows { get; set; }
		/// <summary>
		/// Represents the worksheet protection option of allowing inserting columns.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.WorksheetProtectionOptions_AllowInsertColumns)]
		[Optional]
		bool AllowInsertColumns { get; set; }
		/// <summary>
		/// Represents the worksheet protection option of allowing inserting rows.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.WorksheetProtectionOptions_AllowInsertRows)]
		[Optional]
		bool AllowInsertRows { get; set; }
		/// <summary>
		/// Represents the worksheet protection option of allowing inserting hyperlinks.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.WorksheetProtectionOptions_AllowInsertHyperlinks)]
		[Optional]
		bool AllowInsertHyperlinks { get; set; }
		/// <summary>
		/// Represents the worksheet protection option of allowing deleting columns.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.WorksheetProtectionOptions_AllowDeleteColumns)]
		[Optional]
		bool AllowDeleteColumns { get; set; }
		/// <summary>
		/// Represents the worksheet protection option of allowing deleting rows.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.WorksheetProtectionOptions_AllowDeleteRows)]
		[Optional]
		bool AllowDeleteRows { get; set; }
		/// <summary>
		/// Represents the worksheet protection option of allowing using sort feature.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.WorksheetProtectionOptions_AllowSort)]
		[Optional]
		bool AllowSort { get; set; }
		/// <summary>
		/// Represents the worksheet protection option of allowing using auto filter feature.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.WorksheetProtectionOptions_AllowAutoFilter)]
		[Optional]
		bool AllowAutoFilter { get; set; }
		/// <summary>
		/// Represents the worksheet protection option of allowing using pivot table feature.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.WorksheetProtectionOptions_AllowPivotTables)]
		[Optional]
		bool AllowPivotTables { get; set; }
	}
#endregion Worksheet

#region Range
	/// <summary>
	/// Range represents a set of one or more contiguous cells such as a cell, a row, a column, block of cells, etc.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IRange", InterfaceId = "906962e8-a18a-4cc9-9342-279f056bc293", CoClassName = "Range")]
	public interface Range
	{
		[ClientCallableComMember(DispatchId = DispatchIds.Range_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		[ClientCallableComMember(DispatchId = DispatchIds.Range_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();
		[ClientCallableComMember(DispatchId = DispatchIds.Range_ReferenceId)]
		string _ReferenceId { get; }
		/// <summary>
		/// Represents the range reference in A1-style. Address value will contain the Sheet reference (e.g. Sheet1!A1:B4). Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Address)]
		string Address { get; }
		/// <summary>
		/// Represents range reference for the specified range in the language of the user. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_AddressLocal)]
		string AddressLocal { get; }
		/// <summary>
		/// Number of cells in the range. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_CellCount)]
		int CellCount { get; }
		/// <summary>
		/// Clear range values, format, fill, border, etc.
		/// </summary>
		/// <param name="applyTo">Determines the type of clear action. See Excel.ClearApplyTo for details.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Clear)]
		void Clear([Optional]ClearApplyTo applyTo);
		/// <summary>
		/// Represents the total number of columns in the range. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_ColumnCount)]
		int ColumnCount { get; }
		/// <summary>
		/// Represents the column number of the first cell in the range. Zero-indexed. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_ColumnIndex)]
		int ColumnIndex { get; }
		/// <summary>
		/// Deletes the cells associated with the range.
		/// </summary>
		/// <param name="shift">Specifies which way to shift the cells. See Excel.DeleteShiftDirection for details.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Delete)]
		void Delete(DeleteShiftDirection shift);
		/// <summary>
		/// Gets an object that represents the entire column of the range.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_EntireColumn)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "EntireColumn", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetEntireColumn();
		/// <summary>
		/// Gets an object that represents the entire row of the range.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_EntireRow)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "EntireRow", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetEntireRow();
		/// <summary>
		/// Represents the type of data of each cell. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_ValueTypes)]
		RangeValueType[][] ValueTypes { get; }
		/// <summary>
		/// Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Format)]
		RangeFormat Format { get; }
		/// <summary>
		/// Represents the formula in A1-style notation.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Formulas)]
		object[][] Formulas { get; set; }
		/// <summary>
		/// Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_FormulasLocal)]
		object[][] FormulasLocal { get; set; }
		/// <summary>
		/// Represents the formula in R1C1-style notation.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_FormulasR1C1)]
		object[][] FormulasR1C1 { get; set; }
		/// <summary>
		/// Gets the smallest range object that encompasses the given ranges. For example, the GetBoundingRect of "B2:C5" and "D10:E15" is "B2:E16".
		/// </summary>
		/// <param name="anotherRange">The range object or address or range name.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_GetBoundingRect)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true)]
		Range GetBoundingRect([TypeScriptType("Excel.Range|string")]object anotherRange);
		/// <summary>
		/// Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it's stays within the worksheet grid. The returned cell is located relative to the top left cell of the range.
		/// </summary>
		/// <param name="row">Row number of the cell to be retrieved. Zero-indexed.</param>
		/// <param name="column">Column number of the cell to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Cell)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true)]
		Range GetCell(int row, int column);
		/// <summary>
		/// Gets a column contained in the range.
		/// </summary>
		/// <param name="column">Column number of the range to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_GetColumn)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true)]
		Range GetColumn(int column);
		/// <summary>
		/// Gets the range object that represents the rectangular intersection of the given ranges.
		/// </summary>
		/// <param name="anotherRange">The range object or range address that will be used to determine the intersection of ranges.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_GetIntersection)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true)]
		Range GetIntersection([TypeScriptType("Excel.Range|string")]object anotherRange);
		/// <summary>
		/// Gets the last cell within the range. For example, the last cell of "B2:D5" is "D5".
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_GetLastCell)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "LastCell", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetLastCell();
		/// <summary>
		/// Gets the last column within the range. For example, the last column of "B2:D5" is "D2:D5".
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_GetLastColumn)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "LastColumn", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetLastColumn();
		/// <summary>
		/// Gets the last row within the range. For example, the last row of "B2:D5" is "B5:D5".
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_GetLastRow)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "LastRow", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetLastRow();
		/// <summary>
		/// Gets an object which represents a range that's offset from the specified range. The dimension of the returned range will match this range. If the resulting range is forced outside the bounds of the worksheet grid, an exception will be thrown.
		/// </summary>
		/// <param name="rowOffset">The number of rows (positive, negative, or 0) by which the range is to be offset. Positive values are offset downward, and negative values are offset upward.</param>
		/// <param name="columnOffset">The number of columns (positive, negative, or 0) by which the range is to be offset. Positive values are offset to the right, and negative values are offset to the left.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_GetOffsetRange)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true)]
		Range GetOffsetRange(int rowOffset, int columnOffset);
		/// <summary>
		/// Gets a row contained in the range.
		/// </summary>
		/// <param name="row">Row number of the range to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_GetRow)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true)]
		Range GetRow(int row);
		/// <summary>
		/// Represents if all cells of the current range are hidden.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Hidden)]
		bool? Hidden { get; }
		/// <summary>
		/// Represents if all rows of the current range are hidden.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_RowHidden)]
		bool? RowHidden { get; set; }
		/// <summary>
		/// Represents if all columns of the current range are hidden.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_ColumnHidden)]
		bool? ColumnHidden { get; set; }
		/// <summary>
		/// Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space. Returns a new Range object at the now blank space.
		/// </summary>
		/// <param name="shift">Specifies which way to shift the cells. See Excel.InsertShiftDirection for details.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Insert)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		Range Insert(InsertShiftDirection shift);
		/// <summary>
		/// Merge the range cells into one region in the worksheet.
		/// </summary>
		/// <param name="across">Set true to merge cells in each row of the specified range as separate merged cells. The default value is false.</param> 
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Merge)]
		void Merge([Optional]bool across);
		/// <summary>
		/// Unmerge the range cells into separate cells.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Unmerge)]
		void Unmerge();
		/// <summary>
		/// Represents Excel's number format code for the given cell.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_NumberFormat)]
		object[][] NumberFormat { get; set; }
		/// <summary>
		/// Returns the total number of rows in the range. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_RowCount)]
		int RowCount { get; }
		/// <summary>
		/// Returns the row number of the first cell in the range. Zero-indexed. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_RowIndex)]
		int RowIndex { get; }
		/// <summary>
		/// Selects the specified range in the Excel UI.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Select)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "")]
		void Select();
		/// <summary>
		/// Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Text)]
		object[][] Text { get; }
		/// <summary>
		/// Returns the used range of the given range object.
		/// </summary>
		/// <param name="valuesOnly">Considers only cells with values as used cells.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_UsedRange)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "UsedRange", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetUsedRange([ApiSet(Version = 1.2)][Optional]bool valuesOnly);
		/// <summary>
		/// Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Values)]
		object[][] Values { get; set; }
		/// <summary>
		/// The worksheet containing the current range. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Worksheet)]
		Worksheet Worksheet { get; }

		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Sort)]
		RangeSort Sort { get; }
	}

	/// <summary>
	/// Represents a string reference of the form SheetName!A1:B5, or a global or local named range
	/// </summary>
	[ClientCallableComType(Name = "IRangeReference", InterfaceId = "A253E7A6-82CA-4314-9FEA-411507C37024", CoClassName = "RangeReference", CoClassId = "3A7C6019-23C3-4A18-AEDE-21CD89AAA672")]
	[ApiSet(Version = 1.2)]
	public struct RangeReference
	{
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeReference_Address)]
		string Address { get; set; }
	}

#endregion Range

#region NamedItem
	/// <summary>
	/// A collection of all the nameditem objects that are part of the workbook.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "INamedItemCollection", InterfaceId = "BD4C9F4B-F762-4779-AF4E-9E9665797830", CoClassName = "NamedItemCollection", SupportEnumeration = true)]
	public interface NamedItemCollection : IEnumerable<NamedItem>
	{
		/// <summary>
		/// Gets a nameditem object using its name
		/// </summary>
		/// <param name="name">nameditem name.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.NamedItemCollection_Indexer)]
		NamedItem this[string name] { get; }
	}

	/// <summary>
	/// Represents a defined name for a range of cells or value. Names can be primitive named objects (as seen in the type below), range object, reference to a range. This object can be used to obtain range object associated with names.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "INamedItem", InterfaceId = "E76EE454-3E5E-4187-9389-3C65234609EF", CoClassName = "NamedItem")]
	public interface NamedItem
	{
		[ClientCallableComMember(DispatchId = DispatchIds.NamedItem_Id)]
		string _Id { get; }
		[ClientCallableComMember(DispatchId = DispatchIds.NamedItem_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// The name of the object. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.NamedItem_Name)]
		string Name { get; }
		/// <summary>
		/// Returns the range object that is associated with the name. Throws an exception if the named item's type is not a range.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.NamedItem_Range)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "Range", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetRange();
		/// <summary>
		/// Indicates what type of reference is associated with the name. See Excel.NamedItemType for details. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.NamedItem_Type)]
		NamedItemType? Type { get; }
		/// <summary>
		/// Represents the formula that the name is defined to refer to. E.g. =Sheet14!$B$2:$H$12, =4.75, etc. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.NamedItem_Value)]
		object Value { get; }
		/// <summary>
		/// Specifies whether the object is visible or not.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.NamedItem_Visible)]
		bool Visible { get; set; }
	}
#endregion NamedItem

#region Binding

	/// <summary>
	/// Represents an Office.js binding that is defined in the workbook.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IBinding", InterfaceId = "7957FCE9-D0AF-4302-9F89-6818D8DEC5D5", CoClassName = "Binding")]
	public interface Binding
	{
		[ClientCallableComMember(DispatchId = DispatchIds.Binding_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Represents binding identifier. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Binding_Id)]
		[ClientCallableProperty(ExcludedFromRest = true)]
		string Id { get; }
		/// <summary>
		/// Returns the range represented by the binding. Will throw an error if binding is not of the correct type.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Binding_Range)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "")]
		Range GetRange();
		/// <summary>
		/// Returns the table represented by the binding. Will throw an error if binding is not of the correct type.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Binding_Table)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "")]
		Table GetTable();
		/// <summary>
		/// Returns the text represented by the binding. Will throw an error if binding is not of the correct type.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Binding_Text)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "")]
		string GetText();
		/// <summary>
		/// Returns the type of the binding. See Excel.BindingType for details. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Binding_Type)]
		[ClientCallableProperty(ExcludedFromRest = true)]
		BindingType Type { get; }
	}

	/// <summary>
	/// Represents the collection of all the binding objects that are part of the workbook.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IBindingCollection", InterfaceId = "0D1B5A8F-B3C1-4386-A285-5533EA59846E", CoClassName = "BindingCollection")]
	public interface BindingCollection : IEnumerable<Binding>
	{
		/// <summary>
		/// Gets a binding object by ID.
		/// </summary>
		/// <param name="id">Id of the binding object to be retrieved.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.BindingCollection_Indexer)]
		Binding this[string id] { get; }
		/// <summary>
		/// Returns the number of bindings in the collection. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.BindingCollection_Count)]
		int Count { get; }
		/// <summary>
		/// Gets a binding object based on its position in the items array.
		/// </summary>
		/// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.BindingCollection_ItemAt)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		Binding GetItemAt(int index);
	}

#endregion Binding

#region Table
	/// <summary>
	/// Represents a collection of all the tables that are part of the workbook.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableType(CreateItemOperationName = "Add")]
	[ClientCallableComType(Name = "ITableCollection", InterfaceId = "D0BDE1B5-7F2E-480A-A803-98CE6BEBB873", CoClassName = "TableCollection")]
	public interface TableCollection : IEnumerable<Table>
	{
		[ClientCallableComMember(DispatchId = DispatchIds.TableCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Gets a table by Name or ID.
		/// </summary>
		/// <param name="key">Name or ID of the table to be retrieved.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableCollection_Indexer)]
		Table this[[TypeScriptType("number|string")]object key] { get; }
		/// <summary>
		/// Returns the number of tables in the workbook. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableCollection_Count)]
		int Count { get; }
		/// <summary>
		/// Gets a table based on its position in the collection.
		/// </summary>
		/// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableCollection_ItemAt)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		Table GetItemAt(int index);
		/// <summary>
		/// Create a new table. The range source address determines the worksheet under which the table will be added. If the table cannot be added (e.g., because the address is invalid, or the table would overlap with another table), an error will be thrown.
		/// </summary>
		/// <param name="address">Address or name of the range object representing the data source. If the address does not contain a sheet name, the currently-active sheet is used.</param>
		/// <param name="hasHeaders">Boolean value that indicates whether the data being imported has column labels. If the source does not contain headers (i.e,. when this property set to false), Excel will automatically generate header shifting the data down by one row.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableCollection_Add)]
		Table Add(string address, bool hasHeaders);
	}

	/// <summary>
	/// Represents an Excel table.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableType(DeleteOperationName = "Delete", ConvertIntegerKeyValueToString = true)]
	[ClientCallableComType(Name = "ITable", InterfaceId = "302DF59F-3294-46A2-8046-6A7647C75847", CoClassName = "Table")]
	public interface Table
	{
		[ClientCallableComMember(DispatchId = DispatchIds.Table_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Table_Id)]
		int Id { get; }
		/// <summary>
		/// Name of the table.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Table_Name)]
		string Name { get; set; }
		/// <summary>
		/// Gets the range object associated with the entire table.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Table_Range)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "Range", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetRange();
		/// <summary>
		/// Gets the range object associated with header row of the table.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Table_HeaderRowRange)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "HeaderRowRange", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetHeaderRowRange();
		/// <summary>
		/// Gets the range object associated with the data body of the table.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Table_DataBodyRange)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "DataBodyRange", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetDataBodyRange();
		/// <summary>
		/// Gets the range object associated with totals row of the table.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Table_TotalRowRange)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "TotalRowRange", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetTotalRowRange();
		/// <summary>
		/// Indicates whether the header row is visible or not. This value can be set to show or remove the header row.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Table_ShowHeaders)]
		bool ShowHeaders { get; set; }
		/// <summary>
		/// Indicates whether the total row is visible or not. This value can be set to show or remove the total row.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Table_ShowTotals)]
		bool ShowTotals { get; set; }
		/// <summary>
		/// Constant value that represents the Table style. Possible values are: TableStyleLight1 thru TableStyleLight21, TableStyleMedium1 thru TableStyleMedium28, TableStyleStyleDark1 thru TableStyleStyleDark11. A custom user-defined style present in the workbook can also be specified.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Table_TableStyle)]
		string Style { get; set; }
		/// <summary>
		/// Represents a collection of all the columns in the table. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Table_TableColumns)]
		TableColumnCollection Columns { get; }
		/// <summary>
		/// Represents a collection of all the rows in the table. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Table_TableRows)]
		TableRowCollection Rows { get; }
		/// <summary>
		/// The worksheet containing the current table. Read-only.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Table_Worksheet)]
		Worksheet Worksheet { get; }
		/// <summary>
		/// Deletes the table.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Table_Delete)]
		void Delete();
		/// <summary>
		/// Represents the sorting for the table.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Table_Sort)]
		TableSort Sort { get; }
		/// <summary>
		/// Converts the table into a normal range of cells. All data is preserved.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Table_ConvertToRange)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		Range ConvertToRange();

		/// <summary>
		/// Clears all the filters currently applied on the table.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Table_ClearFilters)]
		void ClearFilters();

		/// <summary>
		/// Reapplies all the filters currently on the table.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Table_ReapplyFilters)]
		void ReapplyFilters();
	}

	/// <summary>
	/// Represents a collection of all the columns that are part of the table.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableType(CreateItemOperationName = "Add")]
	[ClientCallableComType(Name = "ITableColumnCollection", InterfaceId = "97FD1554-DDA6-49CD-9D39-737AF8297E70", CoClassName = "TableColumnCollection")]
	public interface TableColumnCollection : IEnumerable<TableColumn>
	{
		[ClientCallableComMember(DispatchId = DispatchIds.TableColumnCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Gets a column object by Name or ID.
		/// </summary>
		/// <param name="key"> Column Name or ID.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableColumnCollection_Indexer)]
		TableColumn this[[TypeScriptType("number|string")]object key] { get; }
		/// <summary>
		/// Returns the number of columns in the table. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableColumnCollection_Count)]
		int Count { get; }
		/// <summary>
		/// Gets a column based on its position in the collection.
		/// </summary>
		/// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableColumnCollection_ItemAt)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		TableColumn GetItemAt(int index);
		/// <summary>
		/// Adds a new column to the table.
		/// </summary>
		/// <param name="index">Specifies the relative position of the new column. The previous column at this position is shifted to the right. The index value should be equal to or less than the last column's index value, so it cannot be used to append a column at the end of the table. Zero-indexed.</param>
		/// <param name="values">A 2-dimensional array of unformatted values of the table column.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableColumnCollection_Insert)]
		TableColumn Add(int index, [Optional][TypeScriptType("Array<Array<boolean|string|number>>|boolean|string|number")]object values);
	}

	/// <summary>
	/// Represents a column in a table.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableType(DeleteOperationName = "Delete", ConvertIntegerKeyValueToString = true)]
	[ClientCallableComType(Name = "ITableColumn", InterfaceId = "3291F5CF-437F-482D-BAA1-B0F4C2E430D0", CoClassName = "TableColumn")]
	public interface TableColumn
	{
		[ClientCallableComMember(DispatchId = DispatchIds.TableColumn_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Returns a unique key that identifies the column within the table. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableColumn_Id)]
		int Id { get; }
		/// <summary>
		/// Returns the name of the table column. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableColumn_Name)]
		string Name { get; }
		/// <summary>
		/// Returns the index number of the column within the columns collection of the table. Zero-indexed. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableColumn_Index)]
		int Index { get; }
		/// <summary>
		/// Gets the range object associated with the entire column.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableColumn_Range)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "Range", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetRange();
		/// <summary>
		/// Gets the range object associated with the header row of the column.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableColumn_HeaderRowRange)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "HeaderRowRange", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetHeaderRowRange();
		/// <summary>
		/// Gets the range object associated with the data body of the column.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableColumn_DataBodyRange)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "DataBodyRange", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetDataBodyRange();
		/// <summary>
		/// Gets the range object associated with the totals row of the column.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableColumn_TotalRowRange)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "TotalRowRange", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetTotalRowRange();
		/// <summary>
		/// Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableColumn_Values)]
		object[][] Values { get; set; }
		/// <summary>
		/// Deletes the column from the table.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableColumn_Delete)]
		void Delete();
		/// <summary>
		/// Retrieve the filter applied to the column.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableColumn_Filter)]
		Filter Filter { get; }
	}

	/// <summary>
	/// Represents a collection of all the rows that are part of the table.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableType(CreateItemOperationName = "Add")]
	[ClientCallableComType(Name = "ITableRowCollection", InterfaceId = "70544D5B-C1BD-4D4F-A410-87785C4BF2B4", CoClassName = "TableRowCollection")]
	public interface TableRowCollection : IEnumerable<TableRow>
	{
		[ClientCallableComMember(DispatchId = DispatchIds.TableRowCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Returns the number of rows in the table. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableRowCollection_Count)]
		int Count { get; }
		/// <summary>
		/// Gets a row based on its position in the collection.
		/// </summary>
		/// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableRowCollection_ItemAt)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		TableRow GetItemAt(int index);
		/// <summary>
		/// Adds a new row to the table.
		/// </summary>
		/// <param name="index">Specifies the relative position of the new row. If null, the addition happens at the end. Any rows below the inserted row are shifted downwards. Zero-indexed.</param>
		/// <param name="values">A 2-dimensional array of unformatted values of the table row.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableRowCollection_Insert)]
		TableRow Add([Optional]int? index, [Optional][TypeScriptType("Array<Array<boolean|string|number>>|boolean|string|number")]object values);
	}

	/// <summary>
	/// Represents a row in a table.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableType(DeleteOperationName = "Delete")]
	[ClientCallableComType(Name = "ITableRow", InterfaceId = "2604BD8F-678C-4688-9A24-A43F5B3BE4C2", CoClassName = "TableRow")]
	public interface TableRow
	{
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Returns the index number of the row within the rows collection of the table. Zero-indexed. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_Index)]
		int Index { get; }
		/// <summary>
		/// Returns the range object associated with the entire row.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_Range)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "Range", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetRange();
		/// <summary>
		/// Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_Values)]
		object[][] Values { get; set; }
		/// <summary>
		/// Deletes the row from the table.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_Delete)]
		void Delete();
	}
#endregion Table

#region Range Formats
	/// <summary>
	/// A format object encapsulating the range's font, fill, borders, alignment, and other properties.
	/// </summary>
	[ClientCallableComType(Name = "IRangeFormat", InterfaceId = "E97D0B6E-8FBA-4FD5-9922-495283F3C44C", CoClassName = "RangeFormat")]
	[ApiSet(Version = 1.1)]
	public interface RangeFormat
	{
		[ClientCallableComMember(DispatchId = DispatchIds.RangeFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Gets or sets the width of all colums within the range. If the column widths are not uniform, null will be returned.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeFormat_ColumnWidth)]
		double? ColumnWidth { get; set; }
		/// <summary>
		/// Changes the width of the columns of the current range to achieve the best fit, based on the current data in the columns.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeFormat_AutofitColumns)]
		void AutofitColumns();
		/// <summary>
		/// Returns the fill object defined on the overall range. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeFormat_Fill)]
		RangeFill Fill { get; }
		/// <summary>
		/// Collection of border objects that apply to the overall range selected Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeFormat_Borders)]
		RangeBorderCollection Borders { get; }
		/// <summary>
		/// Returns the font object defined on the overall range selected Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeFormat_Font)]
		RangeFont Font { get; }
		/// <summary>
		/// Represents the horizontal alignment for the specified object. See Excel.HorizontalAlignment for details.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeFormat_HorizontalAlignment)]
		HorizontalAlignment? HorizontalAlignment { get; set; }
		/// <summary>
		/// Gets or sets the height of all rows in the range. If the row heights are not uniform null will be returned.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeFormat_RowHeight)]
		double? RowHeight { get; set; }
		/// <summary>
		/// Changes the height of the rows of the current range to achieve the best fit, based on the current data in the columns.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeFormat_AutofitRows)]
		void AutofitRows();
		/// <summary>
		/// Represents the vertical alignment for the specified object. See Excel.VerticalAlignment for details.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeFormat_VerticalAlignment)]
		VerticalAlignment? VerticalAlignment { get; set; }
		/// <summary>
		/// Indicates if Excel wraps the text in the object. A null value indicates that the entire range doesn't have uniform wrap setting
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeFormat_WrapText)]
		bool? WrapText { get; set; }
		/// <summary>
		/// Returns the format protection object for a range.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeFormat_Protection)]
		FormatProtection Protection { get; }
	}

	/// <summary>
	/// Represents the format protection of a range object.
	/// </summary>
	[ApiSet(Version = 1.2)]
	[ClientCallableComType(Name = "IFormatProtection", InterfaceId = "52AB99FC-FBC1-4E4B-B08B-3AD22314A32E", CoClassName = "FormatProtection")]
	public interface FormatProtection
	{
		[ClientCallableComMember(DispatchId = DispatchIds.FormatProtection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Indicates if Excel locks the cells in the object. A null value indicates that the entire range doesn't have uniform lock setting.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.FormatProtection_Locked)]
		bool? Locked { get; set; }
		/// <summary>
		/// Indicates if Excel hides the formula for the cells in the range. A null value indicates that the entire range doesn't have uniform formula hidden setting.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.FormatProtection_FormulaHidden)]
		bool? FormulaHidden { get; set; }
	}

	/// <summary>
	/// Represents the background of a range object.
	/// </summary>
	[ClientCallableComType(Name = "IRangeFill", InterfaceId = "C4514652-D1DB-41D1-8B25-9A27F1B33413", CoClassName = "RangeFill")]
	[ApiSet(Version = 1.1)]
	public interface RangeFill
	{
		[ClientCallableComMember(DispatchId = DispatchIds.RangeFill_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange")
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeFill_Color)]
		string Color { get; set; }
		/// <summary>
		/// Resets the range background.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeFill_Clear)]
		void Clear();
	}

	/// <summary>
	/// Represents the border of an object.
	/// </summary>
	[ClientCallableComType(Name = "IRangeBorder", InterfaceId = "AACFA926-132B-4B49-9D78-1AD4E20B1382", CoClassName = "RangeBorder")]
	[ApiSet(Version = 1.1)]
	public interface RangeBorder
	{
		/// <summary>
		/// Represents border identifier. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.RangeBorder_Id)]
		[ClientCallableProperty(ExcludedFromClientLibrary = true)]
		[ApiSet(Version = 1.1)]
		BorderIndex Id { get; }

		[ClientCallableComMember(DispatchId = DispatchIds.RangeBorder_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeBorder_Color)]
		string Color { get; set; }
		/// <summary>
		/// One of the constants of line style specifying the line style for the border. See Excel.BorderLineStyle for details.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeBorder_LineStyle)]
		BorderLineStyle? Style { get; set; }
		/// <summary>
		/// Constant value that indicates the specific side of the border. See Excel.BorderIndex for details. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeBorder_SideIndex)]
		BorderIndex? SideIndex { get; }
		/// <summary>
		/// Specifies the weight of the border around a range. See Excel.BorderWeight for details.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeBorder_Weight)]
		BorderWeight? Weight { get; set; }
	}

	/// <summary>
	/// Represents the border objects that make up range border.
	/// </summary>
	[ClientCallableComType(Name = "IRangeBorderCollection", InterfaceId = "BD62C8A4-0125-4EB9-9FE5-91E58E718D06", CoClassName = "RangeBorderCollection")]
	[ApiSet(Version = 1.1)]
	public interface RangeBorderCollection : IEnumerable<RangeBorder>
	{
		[ClientCallableComMember(DispatchId = DispatchIds.RangeBorderCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Gets a border object using its name
		/// </summary>
		/// <param name="index">Index value of the border object to be retrieved. See Excel.BorderIndex for details.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeBorderCollection_Indexer)]
		RangeBorder this[BorderIndex index] { get; }
		/// <summary>
		/// Number of border objects in the collection. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeBorderCollection_Count)]
		int Count { get; }
		/// <summary>
		/// Gets a border object using its index
		/// </summary>
		/// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeBorderCollection_ItemAt)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		RangeBorder GetItemAt(int index);
	}

	/// <summary>
	/// This object represents the font attributes (font name, font size, color, etc.) for an object.
	/// </summary>
	[ClientCallableComType(Name = "IRangeFont", InterfaceId = "FAAF874F-30F4-4445-8D6A-F99A6EE81C72", CoClassName = "RangeFont")]
	[ApiSet(Version = 1.1)]
	public interface RangeFont
	{
		[ClientCallableComMember(DispatchId = DispatchIds.RangeFont_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Represents the bold status of font.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeFont_Bold)]
		bool? Bold { get; set; }
		/// <summary>
		/// HTML color code representation of the text color. E.g. #FF0000 represents Red.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeFont_Color)]
		string Color { get; set; }
		/// <summary>
		/// Represents the italic status of the font.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeFont_Italic)]
		bool? Italic { get; set; }
		/// <summary>
		/// Font name (e.g. "Calibri")
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeFont_Name)]
		string Name { get; set; }
		/// <summary>
		/// Font size.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeFont_Size)]
		double? Size { get; set; }
		/// <summary>
		/// Type of underline applied to the font. See Excel.RangeUnderlineStyle for details.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeFont_Underline)]
		RangeUnderlineStyle? Underline { get; set; }
	}
#endregion Formats

#region Charts
	/// <summary>
	/// A collection of all the chart objects on a worksheet.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableType(CreateItemOperationName = "Add", HiddenIndexerMethod = true)]
	[ClientCallableComType(Name = "IChartCollection", InterfaceId = "c70eaacf-0ea6-4a54-b148-c600f9a5f5e4", CoClassName = "ChartCollection")]
	public interface ChartCollection : IEnumerable<Chart>
	{
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartCollection_Indexer)]
		Chart this[string key] { get; }

		/// <summary>
		/// Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.
		/// </summary>
		/// <param name="name">Name of the chart to be retrieved.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartCollection_GetItem)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		Chart GetItem(string name);

		/// <summary>
		/// Creates a new chart.
		/// </summary>
		/// <param name="type">Represents the type of a chart. See Excel.ChartType for details.</param>
		/// <param name="sourceData">The Range object corresponding to the source data.</param>
		/// <param name="seriesBy">Specifies the way columns or rows are used as data series on the chart. See Excel.ChartSeriesBy for details.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartCollection_Add)]
		// Note: while sourceData can accept either a Range object or a string (necessary for REST), we will ONLY allow Range objects in JS.
		// Otherwise, desktop code and WAC behavior diverges, given their different treatement of multi-range areas (WAC disallows them), table expansion (desktop does, WAC doesn't), etc.
		Chart Add(ChartType type, [TypeScriptType("Excel.Range")]object sourceData, [Optional]ChartSeriesBy seriesBy);

		/// <summary>
		/// Returns the number of charts in the worksheet. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartCollection_Count)]
		int Count { get; }

		/// <summary>
		/// Gets a chart based on its position in the collection.
		/// </summary>
		/// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartCollection_ItemAt)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		Chart GetItemAt(int index);
	}

	/// <summary>
	/// Represents a chart object in a workbook.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableType(DeleteOperationName = "Delete")]
	[ClientCallableComType(Name = "IChart", InterfaceId = "b35ce724-5414-4380-8eac-582651db71e7", CoClassName = "Chart")]
	public interface Chart
	{
		[ClientCallableComMember(DispatchId = DispatchIds.Chart_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		[ClientCallableComMember(DispatchId = DispatchIds.Chart_Id)]
		[ClientCallableProperty(ExcludedFromClientLibrary = true)]
		[ApiSet(Version = 1.2)]
		string Id { get; }

		/// <summary>
		/// Represents chart axes. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Chart_Axes)]
		ChartAxes Axes { get; }

		/// <summary>
		/// Represents the datalabels on the chart. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Chart_DataLabels)]
		ChartDataLabels DataLabels { get; }

		/// <summary>
		/// Deletes the chart object.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Chart_Delete)]
		void Delete();

		/// <summary>
		/// Represents the height, in points, of the chart object.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Chart_Height)]
		double Height { get; set; }

		/// <summary>
		/// The distance, in points, from the left side of the chart to the worksheet origin.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Chart_Left)]
		double Left { get; set; }

		/// <summary>
		/// Represents the legend for the chart. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Chart_Legend)]
		ChartLegend Legend { get; }

		/// <summary>
		/// Represents the name of a chart object.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Chart_Name)]
		string Name { get; set; }

		/// <summary>
		/// Represents either a single series or collection of series in the chart. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Chart_Series)]
		ChartSeriesCollection Series { get; }

		/// <summary>
		/// Resets the source data for the chart.
		/// </summary>
		/// <param name="sourceData">The Range object corresponding to the source data.</param>
		/// <param name="seriesBy">Specifies the way columns or rows are used as data series on the chart. Can be one of the following: Auto (default), Rows, Columns. See Excel.ChartSeriesBy for details.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Chart_SetData)]
		// Note: while sourceData can accept either a Range object or a string (necessary for REST), we will ONLY allow Range objects in JS.
		// Otherwise, desktop code and WAC behavior diverges, given their different treatement of multi-range areas (WAC disallows them), table expansion (desktop does, WAC doesn't), etc.
		void SetData([TypeScriptType("Excel.Range")]object sourceData, [Optional]ChartSeriesBy seriesBy);

		/// <summary>
		/// Positions the chart relative to cells on the worksheet.
		/// </summary>
		/// <param name="startCell">The start cell. This is where the chart will be moved to. The start cell is the top-left or top-right cell, depending on the user's right-to-left display settings.</param>
		/// <param name="endCell">(Optional) The end cell. If specified, the chart's width and height will be set to fully cover up this cell/range.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Chart_SetPosition)]
		void SetPosition([TypeScriptType("Excel.Range|string")]object startCell, [Optional][TypeScriptType("Excel.Range|string")]object endCell);

		/// <summary>
		/// Represents the title of the specified chart, including the text, visibility, position and formating of the title. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Chart_Title)]
		ChartTitle Title { get; }

		/// <summary>
		/// Represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Chart_Top)]
		double Top { get; set; }

		/// <summary>
		/// Represents the width, in points, of the chart object.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Chart_Width)]
		double Width { get; set; }

		/// <summary>
		/// Encapsulates the format properties for the chart area. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.Chart_Format)]
		ChartAreaFormat Format { get; }

		/// <summary>
		/// Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.
		/// The aspect ratio is preserved as part of the resizing.
		/// </summary>
		/// <param name="height">(Optional) The desired height of the resulting image.</param>
		/// <param name="width">(Optional) The desired width of the resulting image.</param>
		/// <param name="fittingMode">(Optional) The method used to scale the chart to the specified to the specified dimensions (if both height and width are set)."</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Chart_GetImage)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "Image")]
		System.IO.Stream GetImage([Optional]int width, [Optional]int height, [Optional]ImageFittingMode fittingMode);

		/// <summary>
		/// The worksheet containing the current chart. Read-only.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Chart_Worksheet)]
		Worksheet Worksheet { get; }
	}

	/// <summary>
	/// Encapsulates the format properties for the overall chart area.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartAreaFormat", InterfaceId = "8D3ACDD2-720E-4F0D-B318-8EAA58356A9F", CoClassName = "ChartAreaFormat")]
	public interface ChartAreaFormat
	{
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAreaFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the fill format of an object, which includes background formatting information. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAreaFormat_Fill)]
		ChartFill Fill { get; }

		/// <summary>
		/// Represents the font attributes (font name, font size, color, etc.) for the current object. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAreaFormat_Font)]
		ChartFont Font { get; }
	}

	/// <summary>
	/// Represents a collection of chart series.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartSeriesCollection", InterfaceId = "6FC3E0B3-4A68-4EEE-A181-477EB069BAC1", CoClassName = "ChartSeriesCollection")]
	public interface ChartSeriesCollection : IEnumerable<ChartSeries>
	{
		[ClientCallableComMember(DispatchId = DispatchIds.ChartSeriesCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Returns the number of series in the collection. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartSeriesCollection_Count)]
		int Count { get; }

		/// <summary>
		/// Retrieves a series based on its position in the collection
		/// </summary>
		/// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartSeriesCollection_ItemAt)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		ChartSeries GetItemAt(int index);
	}

	/// <summary>
	/// Represents a series in a chart.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartSeries", InterfaceId = "54454749-3FDB-401D-B5E6-6667F7F80F11", CoClassName = "ChartSeries")]
	public interface ChartSeries
	{
		[ClientCallableComMember(DispatchId = DispatchIds.ChartSeries_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the formatting of a chart series, which includes fill and line formatting. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartSeries_Format)]
		ChartSeriesFormat Format { get; }

		/// <summary>
		/// Represents the name of a series in a chart.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartSeries_Name)]
		string Name { get; set; }

		/// <summary>
		/// Represents a collection of all points in the series. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartSeries_Points)]
		ChartPointsCollection Points { get; }
	}

	/// <summary>
	/// encapsulates the format properties for the chart series
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartSeriesFormat", InterfaceId = "1D3D150E-E2B2-498C-B53C-57F55E9C6CF6", CoClassName = "ChartSeriesFormat")]
	public interface ChartSeriesFormat
	{
		[ClientCallableComMember(DispatchId = DispatchIds.ChartSeriesFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the fill format of a chart series, which includes background formating information. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartSeriesFormat_Fill)]
		ChartFill Fill { get; }

		/// <summary>
		/// Represents line formatting. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartSeriesFormat_Line)]
		ChartLineFormat Line { get; }
	}

	/// <summary>
	/// A collection of all the chart points within a series inside a chart.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartPointsCollection", InterfaceId = "1BDB22BF-3690-4E75-9406-1BF54DB0A127", CoClassName = "ChartPointsCollection")]
	public interface ChartPointsCollection : IEnumerable<ChartPoint>
	{
		[ClientCallableComMember(DispatchId = DispatchIds.ChartPointsCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Returns the number of chart points in the collection. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartPointsCollection_Count)]
		int Count { get; }

		/// <summary>
		/// Retrieve a point based on its position within the series.
		/// </summary>
		/// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartPointsCollection_ItemAt)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		ChartPoint GetItemAt(int index);
	}

	/// <summary>
	/// Represents a point of a series in a chart.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartPoint", InterfaceId = "76E71D2A-FB56-4CC8-9375-AFA5C1052E9C", CoClassName = "ChartPoint")]
	public interface ChartPoint
	{
		[ClientCallableComMember(DispatchId = DispatchIds.ChartPoint_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Encapsulates the format properties chart point. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartPoint_Format)]
		ChartPointFormat Format { get; }

		/// <summary>
		/// Returns the value of a chart point. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartPoint_Value)]
		object Value { get; }
	}

	/// <summary>
	/// Represents formatting object for chart points.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartPointFormat", InterfaceId = "D907B031-B51A-4CE6-B903-004554FBD2D2", CoClassName = "ChartPointFormat")]
	public interface ChartPointFormat
	{
		[ClientCallableComMember(DispatchId = DispatchIds.ChartPointFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the fill format of a chart, which includes background formating information. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAreaFormat_Fill)]
		ChartFill Fill { get; }
	}

	/// <summary>
	/// Represents the chart axes.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartAxes", InterfaceId = "a1635994-4bf2-4358-9a13-924c8ebf53aa", CoClassName = "ChartAxes")]
	public interface ChartAxes
	{
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAxes_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the category axis in a chart. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAxes_Category)]
		ChartAxis CategoryAxis { get; }

		/// <summary>
		/// Represents the series axis of a 3-dimensional chart. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAxes_Series)]
		ChartAxis SeriesAxis { get; }

		/// <summary>
		/// Represents the value axis in an axis. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAxes_Value)]
		ChartAxis ValueAxis { get; }
	}

	/// <summary>
	/// Represents a single axis in a chart.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartAxis", InterfaceId = "f6beb340-c24b-4087-8127-521e79dc326a", CoClassName = "ChartAxis")]
	public interface ChartAxis
	{
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAxis_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the formatting of a chart object, which includes line and font formatting. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAxis_Format)]
		ChartAxisFormat Format { get; }

		/// <summary>
		/// Returns a gridlines object that represents the major gridlines for the specified axis. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAxis_MajorGridlines)]
		ChartGridlines MajorGridlines { get; }

		/// <summary>
		/// Represents the interval between two major tick marks. Can be set to a numeric value or an empty string.  The returned value is always a number.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAxis_MajorUnit)]
		object MajorUnit { get; set; }

		/// <summary>
		/// Represents the maximum value on the value axis.  Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAxis_Maximum)]
		object Maximum { get; set; }

		/// <summary>
		/// Represents the minimum value on the value axis. Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAxis_Minimum)]
		object Minimum { get; set; }

		/// <summary>
		/// Returns a Gridlines object that represents the minor gridlines for the specified axis. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAxis_MinorGridlines)]
		ChartGridlines MinorGridlines { get; }

		/// <summary>
		/// Represents the interval between two minor tick marks. "Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAxis_MinorUnit)]
		object MinorUnit { get; set; }

		/// <summary>
		/// Represents the axis title. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAxis_Title)]
		ChartAxisTitle Title { get; }
	}

	/// <summary>
	/// Encapsulates the format properties for the chart axis.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartAxisFormat", InterfaceId = "3ECEE01A-4340-4F99-82AB-EF9B65646F30", CoClassName = "ChartAxisFormat")]
	public interface ChartAxisFormat
	{
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAxisFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the font attributes (font name, font size, color, etc.) for a chart axis element. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAxisFormat_Font)]
		ChartFont Font { get; }

		/// <summary>
		/// Represents chart line formatting. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAxisFormat_Line)]
		ChartLineFormat Line { get; }
	}

	/// <summary>
	/// Represents the title of a chart axis.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartAxisTitle", InterfaceId = "ecedd0b6-a619-46f1-bf98-09c97aadd9df", CoClassName = "ChartAxisTitle")]
	public interface ChartAxisTitle
	{
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAxisTitle_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the formatting of chart axis title. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAxisTitle_Format)]
		ChartAxisTitleFormat Format { get; }

		/// <summary>
		/// Represents the axis title.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAxisTitle_Text)]
		string Text { get; set; }

		/// <summary>
		/// A boolean that specifies the visibility of an axis title.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAxisTitle_Visible)]
		bool Visible { get; set; }
	}

	/// <summary>
	/// Represents the chart axis title formatting.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartAxisTitleFormat", InterfaceId = "4CE21BA4-E4C0-4F10-A968-61AFFE7C372F", CoClassName = "ChartAxisTitleFormat")]
	public interface ChartAxisTitleFormat
	{
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAxisTitleFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the font attributes, such as font name, font size, color, etc. of chart axis title object. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartAxisTitleFormat_Font)]
		ChartFont Font { get; }
	}

	/// <summary>
	/// Represents a collection of all the data labels on a chart point.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartDataLabels", InterfaceId = "9fe05b7b-dd28-489d-aab5-7497e4d5c346", CoClassName = "ChartDataLabels")]
	public interface ChartDataLabels
	{
		[ClientCallableComMember(DispatchId = DispatchIds.ChartDataLabels_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the format of chart data labels, which includes fill and font formatting. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartDataLabels_Format)]
		ChartDataLabelFormat Format { get; }

		/// <summary>
		/// DataLabelPosition value that represents the position of the data label. See Excel.ChartDataLabelPosition for details.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartDataLabels_Position)]
		ChartDataLabelPosition? Position { get; set; }

		/// <summary>
		/// Boolean value representing if the data label value is visible or not.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartDataLabels_ShowValue)]
		bool? ShowValue { get; set; }

		/// <summary>
		/// Boolean value representing if the data label series name is visible or not.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartDataLabels_ShowSeriesName)]
		bool? ShowSeriesName { get; set; }

		/// <summary>
		/// Boolean value representing if the data label category name is visible or not.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartDataLabels_ShowCategoryName)]
		bool? ShowCategoryName { get; set; }

		/// <summary>
		/// Boolean value representing if the data label legend key is visible or not.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartDataLabels_ShowLegendKey)]
		bool? ShowLegendKey { get; set; }

		/// <summary>
		/// Boolean value representing if the data label percentage is visible or not.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartDataLabels_ShowPercentage)]
		bool? ShowPercentage { get; set; }

		/// <summary>
		/// Boolean value representing if the data label bubble size is visible or not.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartDataLabels_ShowBubbleSize)]
		bool? ShowBubbleSize { get; set; }

		/// <summary>
		/// String representing the separator used for the data labels on a chart.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartDataLabels_Separator)]
		string Separator { get; set; }
	}

	/// <summary>
	/// Encapsulates the format properties for the chart data labels.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartDataLabelFormat", InterfaceId = "B2BD0519-4F5B-43AC-9584-AD507172CC6F", CoClassName = "ChartDataLabelFormat")]
	public interface ChartDataLabelFormat
	{
		[ClientCallableComMember(DispatchId = DispatchIds.ChartDataLabelFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the font attributes (font name, font size, color, etc.) for a chart data label. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartDataLabelFormat_Font)]
		ChartFont Font { get; }

		/// <summary>
		/// Represents the fill format of the current chart data label. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartDataLabelFormat_Fill)]
		ChartFill Fill { get; }
	}

	/// <summary>
	/// Represents major or minor gridlines on a chart axis.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartGridlines", InterfaceId = "7af19b5b-5665-4759-a78e-397318ff75e2", CoClassName = "ChartGridlines")]
	public interface ChartGridlines
	{
		[ClientCallableComMember(DispatchId = DispatchIds.ChartGridlines_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Boolean value representing if the axis gridlines are visible or not.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartGridlines_Visible)]
		bool Visible { get; set; }

		/// <summary>
		/// Represents the formatting of chart gridlines. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartGridlines_Format)]
		ChartGridlinesFormat Format { get; }
	}


	/// <summary>
	/// Encapsulates the format properties for chart gridlines.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartGridlinesFormat", InterfaceId = "DD906913-5D3B-4AC1-88ED-3F2DBC98CB04", CoClassName = "ChartGridlinesFormat")]
	public interface ChartGridlinesFormat
	{
		[ClientCallableComMember(DispatchId = DispatchIds.ChartGridlinesFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents chart line formatting. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartGridlinesFormat_Line)]
		ChartLineFormat Line { get; }
	}

	/// <summary>
	/// Represents the legend in a chart.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartLegend", InterfaceId = "a5c915bf-d752-4b33-95e0-5f84c6e9a46a", CoClassName = "ChartLegend")]
	public interface ChartLegend
	{
		[ClientCallableComMember(DispatchId = DispatchIds.ChartLegend_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the formatting of a chart legend, which includes fill and font formatting. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartLegend_Format)]
		ChartLegendFormat Format { get; }

		/// <summary>
		/// A boolean value the represents the visibility of a ChartLegend object.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartLegend_Visible)]
		bool Visible { get; set; }

		/// <summary>
		/// Represents the position of the legend on the chart. See Excel.ChartLegendPosition for details.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartLegend_Position)]
		ChartLegendPosition? Position { get; set; }

		/// <summary>
		/// Boolean value for whether the chart legend should overlap with the main body of the chart.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartLegend_Overlay)]
		bool? Overlay { get; set; }
	}

	/// <summary>
	/// Encapsulates the format properties of a chart legend.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartLegendFormat", InterfaceId = "B2BD0519-4F5B-43AC-9584-AD507172CC6F", CoClassName = "ChartLegendFormat")]
	public interface ChartLegendFormat
	{
		[ClientCallableComMember(DispatchId = DispatchIds.ChartLegendFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the font attributes such as font name, font size, color, etc. of a chart legend. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartLegendFormat_Font)]
		ChartFont Font { get; }

		/// <summary>
		/// Represents the fill format of an object, which includes background formating information. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartLegendFormat_Fill)]
		ChartFill Fill { get; }
	}

	/// <summary>
	/// Represents a chart title object of a chart.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartTitle", InterfaceId = "953ac91f-9c3a-480c-bdec-15d446ad0b82", CoClassName = "ChartTitle")]
	public interface ChartTitle
	{
		[ClientCallableComMember(DispatchId = DispatchIds.ChartTitle_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the formatting of a chart title, which includes fill and font formatting. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartTitle_Format)]
		ChartTitleFormat Format { get; }

		/// <summary>
		/// Boolean value representing if the chart title will overlay the chart or not.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartTitle_Overlay)]
		bool? Overlay { get; set; }

		/// <summary>
		/// Represents the title text of a chart.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartTitle_Text)]
		string Text { get; set; }

		/// <summary>
		/// A boolean value the represents the visibility of a chart title object.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartTitle_Visible)]
		bool Visible { get; set; }
	}

	/// <summary>
	/// Provides access to the office art formatting for chart title.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartTitleFormat", InterfaceId = "ACA6BCFA-EFDD-4B81-9478-B1508EC42CB9", CoClassName = "ChartTitleFormat")]
	public interface ChartTitleFormat
	{
		[ClientCallableComMember(DispatchId = DispatchIds.ChartTitleFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the font attributes (font name, font size, color, etc.) for an object. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartTitleFormat_Font)]
		ChartFont Font { get; }

		/// <summary>
		/// Represents the fill format of an object, which includes background formating information. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartTitleFormat_Fill)]
		ChartFill Fill { get; }
	}

	/// <summary>
	/// Represents the fill formatting for a chart element.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartFill", InterfaceId = "3147230c-d46d-40ea-b3d8-11970eb8a0af", CoClassName = "ChartFill")]
	public interface ChartFill
	{
		[ClientCallableComMember(DispatchId = DispatchIds.ChartFill_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Clear the fill color of a chart element.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartFill_Clear)]
		void Clear();

		/// <summary>
		/// Sets the fill formatting of a chart element to a uniform color.
		/// </summary>
		/// <param name="color">HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartFill_SolidColor)]
		void SetSolidColor(string color);
	}

	/// <summary>
	/// Enapsulates the formatting options for line elements.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartLineFormat", InterfaceId = "0E0D5F3D-DB8D-46CC-B268-BA1D3D190A38", CoClassName = "ChartLineFormat")]
	public interface ChartLineFormat
	{
		[ClientCallableComMember(DispatchId = DispatchIds.ChartLineFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Clear the line format of a chart element.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartLineFormat_Clear)]
		void Clear();

		/// <summary>
		/// HTML color code representing the color of lines in the chart.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartLineFormat_Color)]
		string Color { get; set; }
	}

	/// <summary>
	/// This object represents the font attributes (font name, font size, color, etc.) for a chart object.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartFont", InterfaceId = "d62d7af0-54f2-4c16-9e0b-8d5a0ff611b2", CoClassName = "ChartFont")]
	public interface ChartFont
	{
		[ClientCallableComMember(DispatchId = DispatchIds.ChartFont_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the bold status of font.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartFont_Bold)]
		bool? Bold { get; set; }

		/// <summary>
		/// HTML color code representation of the text color. E.g. #FF0000 represents Red.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartFont_Color)]
		string Color { get; set; }

		/// <summary>
		/// Represents the italic status of the font.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartFont_Italic)]
		bool? Italic { get; set; }

		/// <summary>
		/// Font name (e.g. "Calibri")
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartFont_Name)]
		string Name { get; set; }

		/// <summary>
		/// Size of the font (e.g. 11)
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartFont_Size)]
		double? Size { get; set; }

		/// <summary>
		/// Type of underline applied to the font. See Excel.ChartUnderlineStyle for details.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = DispatchIds.ChartFont_Underline)]
		ChartUnderlineStyle? Underline { get; set; }
	}

	#endregion Charts

	#region Sort

	/// <summary>
	/// Manages sorting operations on Range objects.
	/// </summary>
	[ApiSet(Version = 1.2)]
	[ClientCallableComType(Name = "IRangeSort", InterfaceId = "8D69987D-B7AD-4AF7-B297-529C21A39ACC", CoClassName = "RangeSort")]
	public interface RangeSort
	{

		[ClientCallableComMember(DispatchId = DispatchIds.RangeSort_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Perform a sort operation.
		/// </summary>
		/// <param name="fields">The list of conditions to sort on.</param>
		/// <param name="matchCase">Whether to have the casing impact string ordering.</param>
		/// <param name="hasHeaders">Whether the range has a header.</param>
		/// <param name="orientation">Whether the operation is sorting rows or columns.</param>
		/// <param name="method">The ordering method used for Chinese characters.</param>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = DispatchIds.RangeSort_Apply)]
		void Apply(SortField[] fields, [Optional] bool matchCase, [Optional] bool hasHeaders, [Optional] SortOrientation orientation, [Optional] SortMethod method);
	}

	/// <summary>
	/// Manages sorting operations on Table objects.
	/// </summary>
	[ApiSet(Version = 1.2)]
	[ClientCallableComType(Name = "ITableSort", InterfaceId = "2FA61C80-F2B7-46A2-8713-AD13E8C3DC4E", CoClassName = "TableSort")]
	public interface TableSort
	{
		[ClientCallableComMember(DispatchId = DispatchIds.TableSort_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Perform a sort operation.
		/// </summary>
		/// <param name="fields">The list of conditions to sort on.</param>
		/// <param name="matchCase">Whether to have the casing impact string ordering.</param>
		/// <param name="method">The ordering method used for Chinese characters.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableSort_Apply)]
		void Apply(SortField[] fields, [Optional] bool matchCase, [Optional] SortMethod method);

		/// <summary>
		/// Represents whether the casing impacted the last sort of the table.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableSort_MatchCase)]
		bool MatchCase { get; }

		/// <summary>
		/// Represents Chinese character ordering method last used to sort the table.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableSort_Method)]
		SortMethod Method { get; }

		/// <summary>
		/// Clears the sorting that is currently on the table. While this doesn't modify the table's ordering, it clears the state of the header buttons.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableSort_Clear)]
		void Clear();

		/// <summary>
		/// Reapplies the current sorting parameters to the table.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableSort_Reapply)]
		void Reapply();

		/// <summary>
		/// Represents the current conditions used to last sort the table.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.TableSort_Fields)]
		SortField[] Fields { get; }
	}

	/// <summary>
	/// Represents a condition in a sorting operation.
	/// </summary>
	[ApiSet(Version = 1.2)]
	[ClientCallableComType(Name = "ISortField", InterfaceId = "DFE7801F-F972-476C-A4ED-1E6E11D59148", CoClassName = "SortField", CoClassId = "9EB4FF82-6464-49F8-908A-A744F934AB17")]
	public struct SortField
	{
		/// <summary>
		/// Represents the column (or row, depending on the sort orientation) that the condition is on. Represented as an offset from the first column (or row).
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.SortField_Key)]
		int Key { get; set; }

		/// <summary>
		/// Represents the type of sorting of this condition.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.SortField_SortOn)]
		[Optional]
		SortOn SortOn { get; set; }

		/// <summary>
		/// Represents whether the sorting is done in an ascending fashion.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.SortField_Ascending)]
		[Optional]
		bool Ascending { get; set; }

		/// <summary>
		/// Represents the color that is the target of the condition if the sorting is on font or cell color.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.SortField_Color)]
		[Optional]
		string Color { get; set; }

		/// <summary>
		/// Represents additional sorting options for this field.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.SortField_DataOption)]
		[Optional]
		SortDataOption DataOption { get; set; }

		/// <summary>
		/// Represents the icon that is the target of the condition if the sorting is on the cell's icon.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.SortField_Icon)]
		[Optional]
		Icon Icon { get; set; }
	}

	#endregion Sort

	#region Filter
	/// <summary>
	/// Manages the filtering of a table's column.
	/// </summary>
	[ApiSet(Version = 1.2)]
	[ClientCallableComType(Name = "IFilter", InterfaceId = "44E193B3-7AE0-4F97-9A63-D79033780ECF", CoClassName = "Filter")]
	public interface Filter
	{
		[ClientCallableComMember(DispatchId = DispatchIds.Filter_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Apply the given filter criteria on the given column.
		/// </summary>
		/// <param name="criteria">The criteria to apply.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Filter_Apply)]
		void Apply(FilterCriteria criteria);

		/// <summary>
		/// Apply a "Bottom Item" filter to the column for the given number of elements.
		/// </summary>
		/// <param name="count">The number of elements from the bottom to show.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Filter_BottomItems)]
		void ApplyBottomItemsFilter(int count);

		/// <summary>
		/// Apply a "Bottom Percent" filter to the column for the given percentage of elements.
		/// </summary>
		/// <param name="percent">The percentage of elements from the bottom to show.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Filter_BottomPercent)]
		void ApplyBottomPercentFilter(int percent);

		/// <summary>
		/// Apply a "Cell Color" filter to the column for the given color.
		/// </summary>
		/// <param name="color">The background color of the cells to show.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Filter_CellColor)]
		void ApplyCellColorFilter(string color);

		/// <summary>
		/// Apply a "Dynamic" filter to the column.
		/// </summary>
		/// <param name="criteria">The dynamic criteria to apply.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Filter_Dynamic)]
		void ApplyDynamicFilter(DynamicFilterCriteria criteria);

		/// <summary>
		/// Apply a "Font Color" filter to the column for the given color.
		/// </summary>
		/// <param name="color">The font color of the cells to show.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Filter_FontColor)]
		void ApplyFontColorFilter(string color);

		/// <summary>
		/// Apply a "Values" filter to the column for the given values.
		/// </summary>
		/// <param name="values">The list of values to show.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Filter_Values)]
		void ApplyValuesFilter([KnownType(typeof(FilterDatetime))][TypeScriptType("Array<string|Excel.FilterDatetime>")]object[] values);

		/// <summary>
		/// Apply a "Top Item" filter to the column for the given number of elements.
		/// </summary>
		/// <param name="count">The number of elements from the top to show.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Filter_TopItems)]
		void ApplyTopItemsFilter(int count);

		/// <summary>
		/// Apply a "Top Percent" filter to the column for the given percentage of elements.
		/// </summary>
		/// <param name="percent">The percentage of elements from the top to show.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Filter_TopPercent)]
		void ApplyTopPercentFilter(int percent);

		/// <summary>
		/// Apply a "Icon" filter to the column for the given icon.
		/// </summary>
		/// <param name="icon">The icons of the cells to show.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Filter_Icon)]
		void ApplyIconFilter(Icon icon);

		/// <summary>
		/// Apply a "Icon" filter to the column for the given criteria strings.
		/// </summary>
		/// <param name="criteria1">The first criteria string.</param>
		/// <param name="criteria2">The second criteria string.</param>
		/// <param name="oper">The operator that describes how the two criteria are joined.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Filter_Custom)]
		void ApplyCustomFilter(string criteria1, [Optional]string criteria2, [Optional]FilterOperator oper);

		/// <summary>
		/// Clear the filter on the given column.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Filter_Clear)]
		void Clear();

		/// <summary>
		/// The currently applied filter on the given column.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Filter_Criteria)]
		FilterCriteria Criteria { get; }
	}

	/// <summary>
	/// Represents the filtering criteria applied to a column.
	/// </summary>
	[ApiSet(Version = 1.2)]
	[ClientCallableComType(Name = "IFilterCriteria", InterfaceId = "BB994BE3-DDDA-4497-9906-8D22855491E8", CoClassName = "FilterCriteria", CoClassId = "9B468C31-9E41-467C-9E5F-24F20B7CB729")]
	public struct FilterCriteria
	{
		/// <summary>
		/// The first criterion used to filter data. Used as an operator in the case of "custom" filtering.
		/// For example ">50" for number greater than 50 or "=*s" for values ending in "s".
		///
		/// Used as a number in the case of top/bottom items/percents. E.g. "5" for the top 5 items if filterOn is set to "topItems"
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.FilterCriteria_Criterion1)]
		[Optional]
		string Criterion1 { get; set; }

		/// <summary>
		/// The second criterion used to filter data. Only used as an operator in the case of "custom" filtering.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.FilterCriteria_Criterion2)]
		[Optional]
		string Criterion2 { get; set; }

		/// <summary>
		/// The HTML color string used to filter cells. Used with "cellColor" and "fontColor" filtering. 
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.FilterCriteria_Color)]
		[Optional]
		string Color { get; set; }

		/// <summary>
		/// The operator used to combine criterion 1 and 2 when using "custom" filtering.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.FilterCriteria_Operator)]
		[Optional]
		FilterOperator Operator { get; set; }

		/// <summary>
		/// The icon used to filter cells. Used with "icon" filtering.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.FilterCriteria_Icon)]
		[Optional]
		Icon Icon { get; set; }

		/// <summary>
		/// The dynamic criteria from the Excel.DynamicFilterCriteria set to apply on this column. Used with "dynamic" filtering.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.FilterCriteria_DynamicCriteria)]
		[Optional]
		DynamicFilterCriteria DynamicCriteria { get; set; }

		/// <summary>
		/// The set of values to be used as part of "values" filtering.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.FilterCriteria_Values)]
		[Optional]
		[TypeScriptType("Array<string|Excel.FilterDatetime>")]
		[KnownType(typeof(FilterDatetime))]
		object[] Values { get; set; }

		/// <summary>
		/// The property used by the filter to determine whether the values should stay visible.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.FilterCriteria_FilterOn)]
		FilterOn FilterOn { get; set; }
	}

	/// <summary>
	/// Represents how to filter a date when filtering on values.
	/// </summary>
	[ApiSet(Version = 1.2)]
	[ClientCallableComType(Name = "IFilterDatetime", InterfaceId = "2F73B0F2-4627-4622-8494-640AF24FB44B", CoClassName = "FilterDatetime", CoClassId = "FFAB0D93-6B73-4F3F-8974-4932D35736E2")]
	public struct FilterDatetime
	{
		/// <summary>
		/// The date in ISO8601 format used to filter data.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.FilterDatetime_Date)]
		string Date { get; set; }

		/// <summary>
		/// How specific the date should be used to keep data. For example, if the date is 2005-04-02 and the specifity is set to "month", the filter operation will keep all rows with a date in the month of april 2009.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.FilterDatetime_Specificity)]
		FilterDatetimeSpecificity Specificity { get; set; }
	}

	#endregion Filter

	/// <summary>
	/// Represents a cell icon.
	/// </summary>
	[ApiSet(Version = 1.2)]
	[ClientCallableComType(Name = "IIcon", InterfaceId = "4FFBA2EE-8527-449C-9C81-739E2795182E", CoClassName = "Icon", CoClassId = "BB897B2C-9B30-4FCA-96B1-E7FFC576FC48")]
	public struct Icon
	{
		/// <summary>
		/// Represents the set that the icon is part of.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Icon_Set)]
		IconSet Set { get; set; }

		/// <summary>
		/// Represents the index of the icon in the given set.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = DispatchIds.Icon_Index)]
		int Index { get; set; }
	}

#region Enums
	[ApiSet(Version = 1.1)]
	public enum BindingType
	{
		Range = 0,
		Table = 1,
		Text = 2
	};

	[ApiSet(Version = 1.1)]
	public enum BorderIndex
	{
		EdgeTop = 0,
		EdgeBottom = 1,
		EdgeLeft = 2,
		EdgeRight = 3,
		InsideVertical = 4,
		InsideHorizontal = 5,
		DiagonalDown = 6,
		DiagonalUp = 7
	};

	[ApiSet(Version = 1.1)]
	public enum BorderLineStyle
	{
		None = 0,
		Continuous = 1,
		Dash = 2,
		DashDot = 3,
		DashDotDot = 4,
		Dot = 5,
		Double = 6,
		SlantDashDot = 7
	};

	[ApiSet(Version = 1.1)]
	public enum BorderWeight
	{
		Hairline = 0,
		Thin = 1,
		Medium = 2,
		Thick = 3
	};

	[ApiSet(Version = 1.1)]
	public enum CalculationMode
	{
		Automatic = 0,
		AutomaticExceptTables = 1,
		Manual = 2,
	}

	[ApiSet(Version = 1.1)]
	public enum CalculationType
	{
		Recalculate = 0,
		Full = 1,
		FullRebuild = 2,
	}

	[ApiSet(Version = 1.1)]
	public enum ClearApplyTo
	{
		All = 0,
		Formats = 1,
		Contents = 2,
	}

	/* Note that this enum must be kept in-sync with its corresponding enum in %SRCROOT%\otools\inc\chart\chartapi\ApiDataLabelPosition.h
	It is based off of the msoElementDataLabel____ set of enumeration values, with the only exceptions of:
	 * Adding an "Invalid", needed for purposes of the API pipeline
	 * Removing "msoElementDataLabelShow", since "show" is not a position... and the data labels will be shown anyway as soon as you set any property on them.
	 * Re-camelcasing "OutSideEnd" to "OutsideEnd", since that's more consisten with the rest of the enumerations.
	Note that for a given chart type, only some of these enumerations will be supported.*/
	[ApiSet(Version = 1.1)]
	public enum ChartDataLabelPosition
	{
		Invalid = 0, /* Needed for indicating "invalid argument" when reading enum values from a string */
		None = 1,
		Center = 2,
		InsideEnd = 3,
		InsideBase = 4,
		OutsideEnd = 5,
		Left = 6,
		Right = 7,
		Top = 8,
		Bottom = 9,
		BestFit = 10,
		Callout = 11,
	}


	/* Note that this enum must be kept in-sync with its corresponding enum in %SRCROOT%\otools\inc\chart\chartapi\ApiLegendPosition.h */
	[ApiSet(Version = 1.1)]
	public enum ChartLegendPosition
	{
		Invalid = 0, /* Needed for indicating "null argument" when reading enum values from a string */
		Top = 1,
		Bottom = 2,
		Left = 3,
		Right = 4,
		Corner = 5,
		Custom = 6
	}

	/* Note that this enum must be kept in-sync with its corresponding enum in %SRCROOT%\otools\inc\chart\chartapi\ApiSeriesBy.h */
	[ApiSet(Version = 1.1)]
	public enum ChartSeriesBy
	{
		Auto = 0, /* Auto is the default */
		Columns = 1,
		Rows = 2,
	}

	/* Note that unlike some of the other enums, this one is based on an existing XlChartType enumeration, 
	 and its value will be passed "as is" to Chart code. It must be kept in sync with the XlChartType
	 enumeration as defined for the public/VBA OM model (see XlChartType in otools\inc\mso\msochart12.odl).
	
	 Note that the enumeration names as defined here lack the "xl" prefix (e.g., just 
	 "ColumnClustered", not "xlColumnClustered"). Those names that begin with "xl3D" (e.g., xl3DPie)
	 are prefaced with "_" (otherwise it wouldn't be a valid variable name once "xl" is stripped).
	 The original names are kept to the side, commented out, to ensure that folks using Research will find
	 the string and update it.
	
	 The enumeration excludes a few of the types that appear to be hidden from VBA: namely, 
	     xlCombo = -4152,
	     xlComboColumnClusteredLine = 113,
	     xlComboColumnClusteredLineSecondaryAxis = 114,
	     xlComboAreaStackedColumnClustered = 115,
	     xlOtherCombinations = 116,
	     xlSuggestedChart = -2,  // this is not a chart type, but a flag specifying the most recommended chart be used
	
	 It currently also excludes the new "Ivy" types, which will be enabled in the future.
	     xlTreemap = 117,
	     xlHistogram = 118,
	     xlWaterfall = 119,
	     xlSunburst = 120,
	     xlBoxwhisker = 121,
	     xlPareto = 122,
	
	 Finally, the enumeration introduces a "None" type, which is needed for code-gen (and that will be
	 checked by our implementation to surface an "invalid argument" error). */
	[ApiSet(Version = 1.1)]
	public enum ChartType /* XlChartType */
	{
		Invalid = 0, /* Needed for indicating "invalid argument" when reading enum values from a string */

		ColumnClustered = 51, /* xlColumnClustered */
		ColumnStacked = 52, /* xlColumnStacked */
		ColumnStacked100 = 53, /* xlColumnStacked100 */
		_3DColumnClustered = 54, /* xl3DColumnClustered */
		_3DColumnStacked = 55, /* xl3DColumnStacked */
		_3DColumnStacked100 = 56, /* xl3DColumnStacked100 */
		BarClustered = 57, /* xlBarClustered */
		BarStacked = 58, /* xlBarStacked */
		BarStacked100 = 59, /* xlBarStacked100 */
		_3DBarClustered = 60, /* xl3DBarClustered */
		_3DBarStacked = 61, /* xl3DBarStacked */
		_3DBarStacked100 = 62, /* xl3DBarStacked100 */
		LineStacked = 63, /* xlLineStacked */
		LineStacked100 = 64, /* xlLineStacked100 */
		LineMarkers = 65, /* xlLineMarkers */
		LineMarkersStacked = 66, /* xlLineMarkersStacked */
		LineMarkersStacked100 = 67, /* xlLineMarkersStacked100 */
		PieOfPie = 68, /* xlPieOfPie */
		PieExploded = 69, /* xlPieExploded */
		_3DPieExploded = 70, /* xl3DPieExploded */
		BarOfPie = 71, /* xlBarOfPie */
		XYScatterSmooth = 72, /* xlXYScatterSmooth */
		XYScatterSmoothNoMarkers = 73, /* xlXYScatterSmoothNoMarkers */
		XYScatterLines = 74, /* xlXYScatterLines */
		XYScatterLinesNoMarkers = 75, /* xlXYScatterLinesNoMarkers */
		AreaStacked = 76, /* xlAreaStacked */
		AreaStacked100 = 77, /* xlAreaStacked100 */
		_3DAreaStacked = 78, /* xl3DAreaStacked */
		_3DAreaStacked100 = 79, /* xl3DAreaStacked100 */
		DoughnutExploded = 80, /* xlDoughnutExploded */
		RadarMarkers = 81, /* xlRadarMarkers */
		RadarFilled = 82, /* xlRadarFilled */
		Surface = 83, /* xlSurface */
		SurfaceWireframe = 84, /* xlSurfaceWireframe */
		SurfaceTopView = 85, /* xlSurfaceTopView */
		SurfaceTopViewWireframe = 86, /* xlSurfaceTopViewWireframe */
		Bubble = 15, /* xlBubble */
		Bubble3DEffect = 87, /* xlBubble3DEffect */
		StockHLC = 88, /* xlStockHLC */
		StockOHLC = 89, /* xlStockOHLC */
		StockVHLC = 90, /* xlStockVHLC */
		StockVOHLC = 91, /* xlStockVOHLC */
		CylinderColClustered = 92, /* xlCylinderColClustered */
		CylinderColStacked = 93, /* xlCylinderColStacked */
		CylinderColStacked100 = 94, /* xlCylinderColStacked100 */
		CylinderBarClustered = 95, /* xlCylinderBarClustered */
		CylinderBarStacked = 96, /* xlCylinderBarStacked */
		CylinderBarStacked100 = 97, /* xlCylinderBarStacked100 */
		CylinderCol = 98, /* xlCylinderCol */
		ConeColClustered = 99, /* xlConeColClustered */
		ConeColStacked = 100, /* xlConeColStacked */
		ConeColStacked100 = 101, /* xlConeColStacked100 */
		ConeBarClustered = 102, /* xlConeBarClustered */
		ConeBarStacked = 103, /* xlConeBarStacked */
		ConeBarStacked100 = 104, /* xlConeBarStacked100 */
		ConeCol = 105, /* xlConeCol */
		PyramidColClustered = 106, /* xlPyramidColClustered */
		PyramidColStacked = 107, /* xlPyramidColStacked */
		PyramidColStacked100 = 108, /* xlPyramidColStacked100 */
		PyramidBarClustered = 109, /* xlPyramidBarClustered */
		PyramidBarStacked = 110, /* xlPyramidBarStacked */
		PyramidBarStacked100 = 111, /* xlPyramidBarStacked100 */
		PyramidCol = 112, /* xlPyramidCol */
		_3DColumn = -4100, /* xl3DColumn */
		Line = 4, /* xlLine */
		_3DLine = -4101, /* xl3DLine */
		_3DPie = -4102, /* xl3DPie */
		Pie = 5, /* xlPie */
		XYScatter = -4169, /* xlXYScatter */
		_3DArea = -4098, /* xl3DArea */
		Area = 1, /* xlArea */
		Doughnut = -4120, /* xlDoughnut */
		Radar = -4151 /* xlRadar */
	}

	/* Right now we only support None or Single underlining. When chart moves to Oart formatting, we'll add the rest.
	 Then we'd be able to support all the style listed in %SRCROOT%\otools\inc\mso\mso.odl */
	[ApiSet(Version = 1.1)]
	public enum ChartUnderlineStyle
	{
		None = 0,
		Single = 1,
	}

	[ApiSet(Version = 1.1)]
	public enum DeleteShiftDirection
	{
		Up = 0,
		Left = 1,
	}

	[ApiSet(Version = 1.2)]
	public enum DynamicFilterCriteria
	{
		Unknown = 0,
		AboveAverage,
		AllDatesInPeriodApril,
		AllDatesInPeriodAugust,
		AllDatesInPeriodDecember,
		AllDatesInPeriodFebruray,
		AllDatesInPeriodJanuary,
		AllDatesInPeriodJuly,
		AllDatesInPeriodJune,
		AllDatesInPeriodMarch,
		AllDatesInPeriodMay,
		AllDatesInPeriodNovember,
		AllDatesInPeriodOctober,
		AllDatesInPeriodQuarter1,
		AllDatesInPeriodQuarter2,
		AllDatesInPeriodQuarter3,
		AllDatesInPeriodQuarter4,
		AllDatesInPeriodSeptember,
		BelowAverage,
		LastMonth,
		LastQuarter,
		LastWeek,
		LastYear,
		NextMonth,
		NextQuarter,
		NextWeek,
		NextYear,
		ThisMonth,
		ThisQuarter,
		ThisWeek,
		ThisYear,
		Today,
		Tomorrow,
		YearToDate,
		Yesterday
	}

	[ApiSet(Version = 1.2)]
	public enum FilterDatetimeSpecificity
	{
		Year,
		Month,
		Day,
		Hour,
		Minute,
		Second
	}

	[ApiSet(Version = 1.2)]
	public enum FilterOn
	{
		BottomItems,
		BottomPercent,
		CellColor,
		Dynamic,
		FontColor,
		Values,
		TopItems,
		TopPercent,
		Icon,
		Custom
	}

	[ApiSet(Version = 1.2)]
	public enum FilterOperator
	{
		And,
		Or,
	}

	[ApiSet(Version = 1.1)]
	public enum HorizontalAlignment
	{
		General = 0,
		Left = 1,
		Center = 2,
		Right = 3,
		Fill = 4,
		Justify = 5,
		CenterAcrossSelection = 6,
		Distributed = 7
	};

	/* This enum must be kept in sync with the KPISETS enum in xlshared/inc/cf/Enums.h
 There are however some transformations that have been applied. */
	[ApiSet(Version = 1.2)]
	public enum IconSet
	{
		Invalid = -1,
		ThreeArrows = 0,
		ThreeArrowsGray,
		ThreeFlags,
		ThreeTrafficLights1,
		ThreeTrafficLights2,
		ThreeSigns,
		ThreeSymbols,
		ThreeSymbols2,
		FourArrows,
		FourArrowsGray,
		FourRedToBlack,
		FourRating,
		FourTrafficLights,
		FiveArrows,
		FiveArrowsGray,
		FiveRating,
		FiveQuarters,
		ThreeStars,
		ThreeTriangles,
		FiveBoxes,
	}

	/* This enum must be kept synchronized with an identically-named enum in xlshared/inc/api/iapichartwrapper.h */
	[ApiSet(Version = 1.2)]
	public enum ImageFittingMode
	{
		Fit = 0,
		FitAndCenter = 1,
		Fill = 2,
	}

	[ApiSet(Version = 1.1)]
	public enum InsertShiftDirection
	{
		Down = 0,
		Right = 1,
	}

	[ApiSet(Version = 1.1)]
	public enum NamedItemType
	{
		String = 0,
		Integer = 1,
		Double = 2,
		Boolean = 3,
		Range = 4,
	}

	[ApiSet(Version = 1.1)]
	public enum RangeUnderlineStyle
	{
		None = 0,
		Single = 1,
		Double = 2,
		SingleAccountant = 3,
		DoubleAccountant = 4
	}

	[ApiSet(Version = 1.1)]
	public enum SheetVisibility
	{
		Visible = 0,
		Hidden = 1,
		VeryHidden = 2
	}

	[ApiSet(Version = 1.1)]
	public enum RangeValueType
	{
		Unknown = 0,
		Empty = 1,
		String = 2,
		Integer =3,
		Double = 4,
		Boolean = 5,
		Error = 6,
	}

	[ApiSet(Version = 1.2)]
	public enum SortOrientation
	{
		Rows = 0,
		Columns = 1,
	}

	[ApiSet(Version = 1.2)]
	public enum SortOn
	{
		Value = 0,
		CellColor = 1,
		FontColor = 2,
		Icon = 3
	}

	[ApiSet(Version = 1.2)]
	public enum SortDataOption
	{
		Normal = 0,
		TextAsNumber = 1
	}

	[ApiSet(Version = 1.2)]
	public enum SortMethod
	{
		PinYin = 0,
		StrokeCount = 1
	}

	[ApiSet(Version = 1.1)]
	public enum VerticalAlignment
	{
		Top = 0,
		Center = 1,
		Bottom = 2,
		Justify = 3,
		Distributed = 4
	}
#endregion Enums

}
