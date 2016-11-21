// --------------------------------------------------------------------------------------------------
//
// <copyright file="WdJscomApi.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// <summary>
// Contains the metadata of Word Jscom API that is currently implemented.
// The following is the workflow to add a new API
// 1) PM add the API to otools\inc\osfclient\RichApi\Word\Metadata\Spec\WdJscomApi.cs
// 2) DEV add the API to otools\inc\osfclient\RichApi\Word\Metadata\Current\WdJscomApi.cs
// 3a) DEV runs word\client\extension\codegen\WdJscomApiGen.bat to re-generate the following files for Word richclient
//      word\client\extension\WdJscomApi.h                COM CoClass header file
//      word\client\extension\WdJscomApi_i.h              COM interface header file
//      word\client\extension\WdJscomApi_i.cpp            COM GUIDs
//      word\client\extension\TypeRegistration.cpp        Type registration file
//      word\client\extension\*.disp.cpp                  COM IDispatch interface related implementation
//      word\client\extension\jscript\WdJscomApi.ts       TypeScript file
// 3b) DEV runs wac\src\WordEditor\Client\Packages\Extension\Codegen\WdWacJscomApiGen.bat to re-generate the following files for Word WAC
//      wac\src\WordEditor\Client\Packages\Extension\WdWacJscomApiTypeReg.cs        Type registration file
//      wac\src\WordEditor\Client\Packages\Extension\*.cs                           Wac API object related implementation
// 4) DEV implement the new API, update word\client\extension\sources.inc or wac\src\WordEditor\Client\Packages\Extension\Bin\sources if necessary.
// </summary>
//
// --------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using Microsoft.OfficeExtension.CodeGen;

[assembly: ClientCallableNamespaceMap(
	"Microsoft.WordServices",
	ComCoClassNamespaceName = "WdJscomApiImpl",
	ComInterfaceNamespaceName = "WdJscomApi",
	TypeScriptNamespaceName = "Word",
	WacNamespaceName = "WordEditor.Extension")]

// NOTE: "idsLcidLocal" will make WdJscomApiHost::GetLocalizedMessage return E_NOTIMPL
//       that will cause to fallback error string below.
// TODO: Add localized error strings and replace "idsLcidLocal".

// Default error (fallback if not uniquely mapped below)
[assembly: HResultDefaultError(HttpStatusCode.InternalServerError, Microsoft.WordServices.ErrorCodes.GeneralException, "idsLcidLocal")]

// Errors we specifically want to hide into general exception (500)
[assembly: HResultError("E_FAIL", HttpStatusCode.InternalServerError, Microsoft.WordServices.ErrorCodes.GeneralException, "idsLcidLocal")]
[assembly: HResultError("E_OUTOFMEMORY", HttpStatusCode.InternalServerError, Microsoft.WordServices.ErrorCodes.GeneralException, "idsLcidLocal")]

// Errors 400s
[assembly: HResultError("E_INVALIDARG", HttpStatusCode.BadRequest, Microsoft.WordServices.ErrorCodes.InvalidArgument, "idsLcidLocal")]
[assembly: HResultError("E_ACCESSDENIED", HttpStatusCode.Forbidden, Microsoft.WordServices.ErrorCodes.AccessDenied, "idsLcidLocal")]
[assembly: HResultError("TYPE_E_ELEMENTNOTFOUND", HttpStatusCode.NotFound, Microsoft.WordServices.ErrorCodes.ItemNotFound, "idsLcidLocal")]

// Errors 500s
[assembly: HResultError("E_NOTIMPL", HttpStatusCode.BadRequest, Microsoft.WordServices.ErrorCodes.NotImplemented, "idsLcidLocal")]

namespace Microsoft.WordServices
{
	// Fallback error strings if localized string is not available.
	internal static class ErrorCodes
	{
		internal const string GeneralException = "GeneralException";
		internal const string InvalidArgument = "InvalidArgument";
		internal const string NotImplemented = "NotImplemented";
		internal const string AccessDenied = "AccessDenied";
		internal const string ItemNotFound = "ItemNotFound";
	}

	/// <summary>
	/// Dispatch Ids
	/// </summary>
	/// <remarks>
	/// Please keep them ordered by type name, then ordered by the value of dispatch id.
	/// </remarks>
	internal static class DispatchIds
	{
		// NOTE: Property dispid values start from 1.
		//       Method dispid values start from 1001
		internal const int MethodBase = 1000;

		//===============================================================================
		// Application dispids
		//===============================================================================
		// Properties
		// Methods
		internal const int Application_CreateDocument = MethodBase + 1;

		//===============================================================================
		// Body dispids
		//===============================================================================
		// Properties
		internal const int Body_ReferenceId = 1;
		internal const int Body_Paragraphs = 2;
		internal const int Body_ContentControls = 3;
		internal const int Body_ParentContentControl = 4;
		internal const int Body_Font = 5;
		internal const int Body_Style = 6;
		internal const int Body_InlinePictures = 7;
		internal const int Body_Text = 8;
		internal const int Body_Type = 9;
		internal const int Body_ParentBody = 10;
		internal const int Body_Lists = 11;
		internal const int Body_Tables = 12;
		internal const int Body_ParentSection = 13;
		internal const int Body_StyleBuiltIn = 14;
		// Methods
		internal const int Body_OnAccess = MethodBase + 1;
		internal const int Body_KeepReference = MethodBase + 2;
		internal const int Body_InsertText = MethodBase + 3;
		internal const int Body_InsertHtml = MethodBase + 4;
		internal const int Body_InsertOoxml = MethodBase + 5;
		internal const int Body_InsertParagraph = MethodBase + 6;
		internal const int Body_InsertContentControl = MethodBase + 7;
		internal const int Body_InsertFileFromBase64 = MethodBase + 8;
		internal const int Body_InsertBreak = MethodBase + 9;
		internal const int Body_InsertInlinePictureFromBase64 = MethodBase + 10;
		internal const int Body_Clear = MethodBase + 11;
		internal const int Body_Search = MethodBase + 12;
		internal const int Body_GetHtml = MethodBase + 13;
		internal const int Body_GetOoxml = MethodBase + 14;
		internal const int Body_Select = MethodBase + 15;
		internal const int Body_GetRange = MethodBase + 16;
		internal const int Body_InsertTable = MethodBase + 17;

		//===============================================================================
		// ContentControl dispids
		//===============================================================================
		// Properties
		internal const int ContentControl_Id = 1;
		internal const int ContentControl_ReferenceId = 2;
		internal const int ContentControl_Title = 3;
		internal const int ContentControl_Tag = 4;
		internal const int ContentControl_PlaceholderText = 5;
		internal const int ContentControl_Type = 6;
		internal const int ContentControl_Appearance = 7;
		internal const int ContentControl_Color = 8;
		internal const int ContentControl_RemoveWhenEdited = 9;
		internal const int ContentControl_CannotDelete = 10;
		internal const int ContentControl_CannotEdit = 11;
		internal const int ContentControl_Font = 12;
		internal const int ContentControl_Style = 13;
		internal const int ContentControl_Paragraphs = 14;
		internal const int ContentControl_ContentControls = 15;
		internal const int ContentControl_ParentContentControl = 16;
		internal const int ContentControl_InlinePictures = 17;
		internal const int ContentControl_Text = 18;
		internal const int ContentControl_Lists = 19;
		internal const int ContentControl_Tables = 20;
		internal const int ContentControl_ParentTableCell = 21;
		internal const int ContentControl_ParentTable = 22;
		internal const int ContentControl_Subtype = 23;
		internal const int ContentControl_StyleBuiltIn = 24;
		internal const int ContentControl_ParentBody = 25;
		// Methods
		internal const int ContentControl_OnAccess = MethodBase + 1;
		internal const int ContentControl_KeepReference = MethodBase + 2;
		internal const int ContentControl_InsertText = MethodBase + 3;
		internal const int ContentControl_InsertHtml = MethodBase + 4;
		internal const int ContentControl_InsertOoxml = MethodBase + 5;
		internal const int ContentControl_InsertFileFromBase64 = MethodBase + 6;
		internal const int ContentControl_InsertParagraph = MethodBase + 7;
		internal const int ContentControl_InsertBreak = MethodBase + 8;
		internal const int ContentControl_InsertInlinePictureFromBase64 = MethodBase + 9;
		internal const int ContentControl_Clear = MethodBase + 10;
		internal const int ContentControl_Delete = MethodBase + 11;
		internal const int ContentControl_Select = MethodBase + 12;
		internal const int ContentControl_GetHtml = MethodBase + 13;
		internal const int ContentControl_GetOoxml = MethodBase + 14;
		internal const int ContentControl_Search = MethodBase + 15;
		internal const int ContentControl_GetRange = MethodBase + 16;
		internal const int ContentControl_Split = MethodBase + 17;
		internal const int ContentControl_InsertTable = MethodBase + 18;
		internal const int ContentControl_GetTextRanges = MethodBase + 19;

		//===============================================================================
		// ContentControlCollection dispids
		//===============================================================================
		// Properties
		internal const int ContentControlCollection_Indexer = 1;
		internal const int ContentControlCollection_ReferenceId = 2;
		// skip 3
		// Methods
		internal const int ContentControlCollection_KeepReference = MethodBase + 1;
		internal const int ContentControlCollection_GetById = MethodBase + 2;
		internal const int ContentControlCollection_GetByTitle = MethodBase + 3;
		internal const int ContentControlCollection_GetByTag = MethodBase + 4;
		internal const int ContentControlCollection_OnAccess = MethodBase + 5;
		internal const int ContentControlCollection_GetByTypes = MethodBase + 6;
		internal const int ContentControlCollection_GetFirst = MethodBase + 7;

		//===============================================================================
		// CustomProperty dispids
		//===============================================================================
		// Properties
		internal const int CustomProperty_ReferenceId = 1;
		internal const int CustomProperty_Key = 2;
		internal const int CustomProperty_Value = 3;
		internal const int CustomProperty_Type = 4;
		// Methods
		internal const int CustomProperty_OnAccess = MethodBase + 1;
		internal const int CustomProperty_KeepReference = MethodBase + 2;
		internal const int CustomProperty_Delete = MethodBase + 3;

		//===============================================================================
		// CustomPropertyCollection dispids
		//===============================================================================
		// Properties
		internal const int CustomPropertyCollection_Indexer = 1;
		internal const int CustomPropertyCollection_ReferenceId = 2;
		// Methods
		internal const int CustomPropertyCollection_OnAccess = MethodBase + 1;
		internal const int CustomPropertyCollection_KeepReference = MethodBase + 2;
		internal const int CustomPropertyCollection_Set = MethodBase + 3;
		internal const int CustomPropertyCollection_GetCount = MethodBase + 4;
		internal const int CustomPropertyCollection_DeleteAll = MethodBase + 5;

		//===============================================================================
		// Document dispids
		//===============================================================================
		// Properties
		internal const int Document_Saved = 1;
		internal const int Document_Sections = 2;
		internal const int Document_Body = 3;
		internal const int Document_ContentControls = 4;
		internal const int Document_ReferenceId = 5;
		internal const int Document_Properties = 6;
		internal const int Document_Settings = 7;
		// Methods
		internal const int Document_GetObjectByReferenceId = MethodBase + 1;
		internal const int Document_GetObjectTypeNameByReferenceId = MethodBase + 2;
		internal const int Document_RemoveReference = MethodBase + 3;
		internal const int Document_RemoveAllReferences = MethodBase + 4;
		internal const int Document_Save = MethodBase + 5;
		internal const int Document_GetSelection = MethodBase + 6;
		internal const int Document_OnAccess = MethodBase + 7;
		internal const int Document_KeepReference = MethodBase + 8;
		internal const int Document_Open = MethodBase + 9;
		internal const int Document_GetBookmarkRange = MethodBase + 10;
		internal const int Document_DeleteBookmark = MethodBase + 11;

		//===============================================================================
		// DocumentProperties dispids
		//===============================================================================
		// Properties
		internal const int DocumentProperties_ReferenceId = 1;
		internal const int DocumentProperties_CustomProperties = 2;
		internal const int DocumentProperties_Title = 3;
		internal const int DocumentProperties_Subject = 4;
		internal const int DocumentProperties_Author = 5;
		internal const int DocumentProperties_Keywords = 6;
		internal const int DocumentProperties_Comments = 7;
		internal const int DocumentProperties_Template = 8;
		internal const int DocumentProperties_LastAuthor = 9;
		internal const int DocumentProperties_RevisionNumber = 10;
		internal const int DocumentProperties_ApplicationName = 11;
		internal const int DocumentProperties_LastPrintDate = 12;
		internal const int DocumentProperties_CreationDate = 13;
		internal const int DocumentProperties_LastSaveTime = 14;
		internal const int DocumentProperties_Security = 19;
		internal const int DocumentProperties_Category = 20;
		internal const int DocumentProperties_Format = 21;
		internal const int DocumentProperties_Manager = 22;
		internal const int DocumentProperties_Company = 23;
		// Methods
		internal const int DocumentProperties_OnAccess = MethodBase + 1;
		internal const int DocumentProperties_KeepReference = MethodBase + 2;

		//===============================================================================
		// Font dispids
		//===============================================================================
		// Properties
		internal const int Font_ReferenceId = 1;
		internal const int Font_Name = 2;
		internal const int Font_Size = 3;
		internal const int Font_Bold = 4;
		internal const int Font_Italic = 5;
		internal const int Font_Color = 6;
		internal const int Font_Underline = 7;
		internal const int Font_Subscript = 8;
		internal const int Font_Superscript = 9;
		internal const int Font_StrikeThrough = 10;
		internal const int Font_DoubleStrikeThrough = 11;
		internal const int Font_HighlightColor = 12;
		// Methods
		internal const int Font_OnAccess = MethodBase + 1;
		internal const int Font_KeepReference = MethodBase + 2;

		//===============================================================================
		// InlinePicture dispids
		//===============================================================================
		// Properties
		internal const int InlinePicture_Id = 1;
		internal const int InlinePicture_ReferenceId = 2;
		internal const int InlinePicture_AltTextDescription = 3;
		internal const int InlinePicture_AltTextTitle = 4;
		internal const int InlinePicture_Height = 5;
		internal const int InlinePicture_Hyperlink = 6;
		internal const int InlinePicture_LockAspectRatio = 7;
		internal const int InlinePicture_Width = 8;
		internal const int InlinePicture_ParentContentControl = 9;
		internal const int InlinePicture_Paragraph = 10;
		internal const int InlinePicture_ImageFormat = 11;
		internal const int InlinePicture_ParentTableCell = 13;
		internal const int InlinePicture_ParentTable = 14;
		// Methods
		internal const int InlinePicture_OnAccess = MethodBase + 1;
		internal const int InlinePicture_KeepReference = MethodBase + 2;
		internal const int InlinePicture_GetBase64ImageSrc = MethodBase + 3;
		internal const int InlinePicture_InsertContentControl = MethodBase + 4;
		internal const int InlinePicture_InsertInlinePictureFromBase64 = MethodBase + 5;
		internal const int InlinePicture_InsertBreak = MethodBase + 6;
		internal const int InlinePicture_InsertText = MethodBase + 7;
		internal const int InlinePicture_InsertHtml = MethodBase + 8;
		internal const int InlinePicture_InsertOoxml = MethodBase + 9;
		internal const int InlinePicture_InsertParagraph = MethodBase + 10;
		internal const int InlinePicture_InsertFileFromBase64 = MethodBase + 11;
		internal const int InlinePicture_Delete = MethodBase + 12;
		internal const int InlinePicture_Select = MethodBase + 13;
		internal const int InlinePicture_GetRange = MethodBase + 14;
		internal const int InlinePicture_GetNext = MethodBase + 15;

		//===============================================================================
		// InlinePictureCollection dispids
		//===============================================================================
		// Properties
		internal const int InlinePictureCollection_Indexer = 1;
		internal const int InlinePictureCollection_ReferenceId = 2;
		// skip 3
		// Methods
		internal const int InlinePictureCollection_KeepReference = MethodBase + 1;
		internal const int InlinePictureCollection_OnAccess = MethodBase + 2;
		internal const int InlinePictureCollection_GetFirst = MethodBase + 3;

		//===============================================================================
		// List dispids
		//===============================================================================
		// Properties
		internal const int List_Id = 1;
		internal const int List_ReferenceId = 2;
		internal const int List_Paragraphs = 3;
		internal const int List_LevelTypes = 4;
		internal const int List_LevelExistences = 5;
		// Methods
		internal const int List_KeepReference = MethodBase + 1;
		internal const int List_OnAccess = MethodBase + 2;
		internal const int List_InsertParagraph = MethodBase + 3;
		internal const int List_GetLevelParagraphs = MethodBase + 4;
		internal const int List_SetLevelBullet = MethodBase + 5;
		internal const int List_SetLevelNumbering = MethodBase + 6;
		internal const int List_GetLevelString = MethodBase + 7;
		internal const int List_SetLevelPicture = MethodBase + 8;
		internal const int List_GetLevelPicture = MethodBase + 9;
		internal const int List_GetLevelFont = MethodBase + 10;
		internal const int List_ResetLevelFont = MethodBase + 11;
		internal const int List_SetLevelAlignment = MethodBase + 12;
		internal const int List_SetLevelIndents = MethodBase + 13;
		internal const int List_SetLevelStartingNumber = MethodBase + 14;

		//===============================================================================
		// ListCollection dispids
		//===============================================================================
		// Properties
		internal const int ListCollection_Indexer = 1;
		internal const int ListCollection_ReferenceId = 2;
		// skip 3
		// Methods
		internal const int ListCollection_KeepReference = MethodBase + 1;
		internal const int ListCollection_GetById = MethodBase + 2;
		internal const int ListCollection_OnAccess = MethodBase + 3;
		internal const int ListCollection_GetFirst = MethodBase + 4;

		//===============================================================================
		// ListItem dispids
		//===============================================================================
		// Properties
		internal const int ListItem_ReferenceId = 1;
		internal const int ListItem_SiblingIndex = 2;
		internal const int ListItem_ListString = 3;
		internal const int ListItem_Level = 4;
		// Methods
		internal const int ListItem_OnAccess = MethodBase + 1;
		internal const int ListItem_KeepReference = MethodBase + 2;
		internal const int ListItem_GetAncestor = MethodBase + 3;
		internal const int ListItem_GetDescendants = MethodBase + 4;

		//===============================================================================
		// Paragraph dispids
		//===============================================================================
		// Properties
		internal const int Paragraph_Id = 1;
		internal const int Paragraph_ReferenceId = 2;
		internal const int Paragraph_Font = 3;
		internal const int Paragraph_Style = 4;
		internal const int Paragraph_ContentControls = 5;
		internal const int Paragraph_ParentContentControl = 6;
		internal const int Paragraph_Alignment = 7;
		internal const int Paragraph_FirstLineIndent = 8;
		internal const int Paragraph_LeftIndent = 9;
		internal const int Paragraph_RightIndent = 10;
		internal const int Paragraph_LineSpacing = 11;
		internal const int Paragraph_OutlineLevel = 12;
		internal const int Paragraph_SpaceBefore = 13;
		internal const int Paragraph_SpaceAfter = 14;
		internal const int Paragraph_LineUnitBefore = 15;
		internal const int Paragraph_LineUnitAfter = 16;
		internal const int Paragraph_InlinePictures = 17;
		internal const int Paragraph_Text = 18;
		internal const int Paragraph_IsListItem = 19;
		internal const int Paragraph_TableNestingLevel = 22;
		internal const int Paragraph_ParentBody = 23;
		internal const int Paragraph_List = 24;
		internal const int Paragraph_ParentTableCell = 25;
		internal const int Paragraph_ParentTable = 26;
		internal const int Paragraph_ListItem = 27;
		internal const int Paragraph_IsLastParagraph = 28;
		internal const int Paragraph_StyleBuiltIn = 29;
		// Methods
		internal const int Paragraph_OnAccess = MethodBase + 1;
		internal const int Paragraph_KeepReference = MethodBase + 2;
		internal const int Paragraph_InsertInlinePictureFromBase64 = MethodBase + 3;
		internal const int Paragraph_InsertContentControl = MethodBase + 4;
		internal const int Paragraph_InsertText = MethodBase + 5;
		internal const int Paragraph_InsertHtml = MethodBase + 6;
		internal const int Paragraph_InsertOoxml = MethodBase + 7;
		internal const int Paragraph_InsertFileFromBase64 = MethodBase + 8;
		internal const int Paragraph_InsertParagraph = MethodBase + 9;
		internal const int Paragraph_InsertBreak = MethodBase + 10;
		internal const int Paragraph_Clear = MethodBase + 11;
		internal const int Paragraph_Delete = MethodBase + 12;
		internal const int Paragraph_Select = MethodBase + 13;
		internal const int Paragraph_GetHtml = MethodBase + 14;
		internal const int Paragraph_GetOoxml = MethodBase + 15;
		internal const int Paragraph_Search = MethodBase + 16;
		internal const int Paragraph_GetRange = MethodBase + 17;
		internal const int Paragraph_Split = MethodBase + 18;
		internal const int Paragraph_InsertTable = MethodBase + 19;
		internal const int Paragraph_GetTextRanges = MethodBase + 20;
		internal const int Paragraph_StartNewList = MethodBase + 21;
		internal const int Paragraph_AttachToList = MethodBase + 22;
		internal const int Paragraph_DetachFromList = MethodBase + 23;
		internal const int Paragraph_GetNext = MethodBase + 24;
		internal const int Paragraph_GetPrevious = MethodBase + 25;

		//===============================================================================
		// ParagraphCollection dispids
		//===============================================================================
		// Properties
		internal const int ParagraphCollection_Indexer = 1;
		internal const int ParagraphCollection_ReferenceId = 2;
		// skip 3
		// skip 4
		// Methods
		internal const int ParagraphCollection_KeepReference = MethodBase + 1;
		internal const int ParagraphCollection_OnAccess = MethodBase + 2;
		internal const int ParagraphCollection_GetFirst = MethodBase + 3;
		internal const int ParagraphCollection_GetLast = MethodBase + 4;

		//===============================================================================
		// Range dispids
		//===============================================================================
		// Properties
		internal const int Range_Id = 1;
		internal const int Range_ReferenceId = 2;
		internal const int Range_Font = 3;
		internal const int Range_Style = 4;
		internal const int Range_Paragraphs = 5;
		internal const int Range_ContentControls = 6;
		internal const int Range_ParentContentControl = 7;
		internal const int Range_InlinePictures = 8;
		internal const int Range_Text = 9;
		internal const int Range_IsEmpty = 10;
		internal const int Range_Hyperlink = 11;
		internal const int Range_Lists = 12;
		internal const int Range_Tables = 13;
		internal const int Range_ParentTableCell = 14;
		internal const int Range_ParentTable = 15;
		internal const int Range_ParentBody = 16;
		internal const int Range_StyleBuiltIn = 17;
		// Methods
		internal const int Range_OnAccess = MethodBase + 1;
		internal const int Range_KeepReference = MethodBase + 2;
		internal const int Range_InsertContentControl = MethodBase + 3;
		internal const int Range_InsertText = MethodBase + 4;
		internal const int Range_InsertHtml = MethodBase + 5;
		internal const int Range_InsertOoxml = MethodBase + 6;
		internal const int Range_InsertFileFromBase64 = MethodBase + 7;
		internal const int Range_InsertParagraph = MethodBase + 8;
		internal const int Range_InsertBreak = MethodBase + 9;
		internal const int Range_InsertInlinePictureFromBase64 = MethodBase + 10;
		internal const int Range_Clear = MethodBase + 11;
		internal const int Range_Delete = MethodBase + 12;
		internal const int Range_Select = MethodBase + 13;
		internal const int Range_GetHtml = MethodBase + 14;
		internal const int Range_GetOoxml = MethodBase + 15;
		internal const int Range_Search = MethodBase + 16;
		internal const int Range_GetRange = MethodBase + 17;
		internal const int Range_Split = MethodBase + 18;
		internal const int Range_CompareLocationWith = MethodBase + 19;
		internal const int Range_ExpandTo = MethodBase + 20;
		internal const int Range_IntersectWith = MethodBase + 21;
		internal const int Range_GetNextTextRange = MethodBase + 22;
		internal const int Range_GetHyperlinkRanges = MethodBase + 23;
		internal const int Range_InsertTable = MethodBase + 24;
		internal const int Range_GetTextRanges = MethodBase + 25;
		internal const int Range_GetBookmarks = MethodBase + 26;
		internal const int Range_InsertBookmark = MethodBase + 27;

		//===============================================================================
		// RangeCollection dispids
		//===============================================================================
		// Properties
		internal const int RangeCollection_Indexer = 1;
		internal const int RangeCollection_ReferenceId = 2;
		// skip 3
		// Methods
		internal const int RangeCollection_KeepReference = MethodBase + 1;
		internal const int RangeCollection_OnAccess = MethodBase + 2;
		internal const int RangeCollection_GetFirst = MethodBase + 3;

		//===============================================================================
		// SearchOptions dispids
		//===============================================================================
		// Properties
		internal const int SearchOptions_IgnorePunct = 1;
		internal const int SearchOptions_IgnoreSpace = 2;
		internal const int SearchOptions_MatchCase = 3;
		internal const int SearchOptions_MatchPrefix = 4;
		internal const int SearchOptions_MatchSuffix = 6;
		internal const int SearchOptions_MatchWildcards = 7;
		internal const int SearchOptions_MatchWholeWord = 8;

		//===============================================================================
		// Section dispids
		//===============================================================================
		// Properties
		internal const int Section_Id = 1;
		internal const int Section_ReferenceId = 2;
		internal const int Section_Body = 3;
		// skip 4
		// Methods
		internal const int Section_OnAccess = MethodBase + 1;
		internal const int Section_KeepReference = MethodBase + 2;
		internal const int Section_GetHeader = MethodBase + 3;
		internal const int Section_GetFooter = MethodBase + 4;
		internal const int Section_GetNext = MethodBase + 5;

		//===============================================================================
		// SectionCollection dispids
		//===============================================================================
		// Properties
		internal const int SectionCollection_Indexer = 1;
		internal const int SectionCollection_ReferenceId = 2;
		// skip 3
		// Methods
		internal const int SectionCollection_KeepReference = MethodBase + 1;
		internal const int SectionCollection_OnAccess = MethodBase + 2;
		internal const int SectionCollection_GetFirst = MethodBase + 3;

		//===============================================================================
		// Setting dispids
		//===============================================================================
		// Properties
		internal const int Setting_ReferenceId = 1;
		internal const int Setting_Key = 2;
		internal const int Setting_Value = 3;
		// Methods
		internal const int Setting_OnAccess = MethodBase + 1;
		internal const int Setting_KeepReference = MethodBase + 2;
		internal const int Setting_Delete = MethodBase + 3;

		//===============================================================================
		// SettingCollection dispids
		//===============================================================================
		// Properties
		internal const int SettingCollection_Indexer = 1;
		internal const int SettingCollection_ReferenceId = 2;
		// Methods
		internal const int SettingCollection_OnAccess = MethodBase + 1;
		internal const int SettingCollection_KeepReference = MethodBase + 2;
		internal const int SettingCollection_Set = MethodBase + 3;
		internal const int SettingCollection_GetCount = MethodBase + 4;
		internal const int SettingCollection_DeleteAll = MethodBase + 5;

		//===============================================================================
		// Table dispids
		//===============================================================================
		// Properties
		internal const int Table_Id = 1;
		internal const int Table_ReferenceId = 2;
		internal const int Table_Rows = 3;
		internal const int Table_IsUniform = 4;
		internal const int Table_Tables = 5;
		internal const int Table_NestingLevel = 6;
		internal const int Table_ParentTableCell = 7;
		internal const int Table_ParentTable = 8;
		internal const int Table_Values = 9;
		internal const int Table_Style = 10;
		internal const int Table_RowCount = 11;
		internal const int Table_HeaderRowCount = 12;
		internal const int Table_StyleTotalRow = 13;
		internal const int Table_StyleFirstColumn = 14;
		internal const int Table_StyleLastColumn = 15;
		internal const int Table_StyleBandedRows = 16;
		internal const int Table_StyleBandedColumns = 17;
		internal const int Table_ShadingColor = 18;
		internal const int Table_HorizontalAlignment = 22;
		internal const int Table_VerticalAlignment = 23;
		internal const int Table_Font = 24;
		internal const int Table_ParentContentControl = 25;
		internal const int Table_Height = 26;
		internal const int Table_Width = 27;
		internal const int Table_ParagraphBefore = 28;
		internal const int Table_ParagraphAfter = 29;
		internal const int Table_StyleBuiltIn = 31;
		internal const int Table_ParentBody = 32;
		// Methods
		internal const int Table_OnAccess = MethodBase + 1;
		internal const int Table_KeepReference = MethodBase + 2;
		internal const int Table_AddRows = MethodBase + 3;
		internal const int Table_AddColumns = MethodBase + 4;
		internal const int Table_GetCell = MethodBase + 5;
		internal const int Table_MergeCells = MethodBase + 6;
		internal const int Table_Delete = MethodBase + 7;
		internal const int Table_Clear = MethodBase + 8;
		internal const int Table_DeleteRows = MethodBase + 9;
		internal const int Table_DeleteColumns = MethodBase + 10;
		internal const int Table_AutoFitContents = MethodBase + 11;
		internal const int Table_AutoFitWindow = MethodBase + 12;
		internal const int Table_DistributeRows = MethodBase + 13;
		internal const int Table_DistributeColumns = MethodBase + 14;
		internal const int Table_GetBorder = MethodBase + 15;
		internal const int Table_Select = MethodBase + 17;
		internal const int Table_Search = MethodBase + 18;
		internal const int Table_GetRange = MethodBase + 19;
		internal const int Table_InsertContentControl = MethodBase + 20;
		internal const int Table_InsertTable = MethodBase + 21;
		internal const int Table_InsertParagraph = MethodBase + 22;
		internal const int Table_GetCellPadding = MethodBase + 23;
		internal const int Table_SetCellPadding = MethodBase + 24;
		internal const int Table_GetNext = MethodBase + 25;

		//===============================================================================
		// TableCollection dispids
		//===============================================================================
		// Properties
		internal const int TableCollection_Indexer = 1;
		internal const int TableCollection_ReferenceId = 2;
		// skip 3
		// Methods
		internal const int TableCollection_KeepReference = MethodBase + 1;
		internal const int TableCollection_OnAccess = MethodBase + 2;
		internal const int TableCollection_GetFirst = MethodBase + 3;

		//===============================================================================
		// TableRow dispids
		//===============================================================================
		// Properties
		internal const int TableRow_Id = 1;
		internal const int TableRow_ReferenceId = 2;
		internal const int TableRow_Cells = 3;
		internal const int TableRow_CellCount = 4;
		internal const int TableRow_ParentTable = 5;
		internal const int TableRow_RowIndex = 6;
		internal const int TableRow_Values = 7;
		internal const int TableRow_ShadingColor = 8;
		internal const int TableRow_HorizontalAlignment = 12;
		internal const int TableRow_VerticalAlignment = 13;
		internal const int TableRow_Font = 14;
		internal const int TableRow_ContentControls = 15;
		internal const int TableRow_IsHeader = 16;
		internal const int TableRow_PreferredHeight = 17;
		// skip 18
		// Methods
		internal const int TableRow_OnAccess = MethodBase + 1;
		internal const int TableRow_KeepReference = MethodBase + 2;
		internal const int TableRow_InsertRows = MethodBase + 3;
		internal const int TableRow_Merge = MethodBase + 4;
		internal const int TableRow_Delete = MethodBase + 5;
		internal const int TableRow_Clear = MethodBase + 6;
		internal const int TableRow_Select = MethodBase + 7;
		internal const int TableRow_Search = MethodBase + 8;
		internal const int TableRow_GetBorder = MethodBase + 9;
		internal const int TableRow_GetCellPadding = MethodBase + 10;
		internal const int TableRow_SetCellPadding = MethodBase + 11;
		internal const int TableRow_GetNext = MethodBase + 12;

		//===============================================================================
		// TableRowCollection dispids
		//===============================================================================
		// Properties
		internal const int TableRowCollection_Indexer = 1;
		internal const int TableRowCollection_ReferenceId = 2;
		// skip 3
		// Methods
		internal const int TableRowCollection_OnAccess = MethodBase + 1;
		internal const int TableRowCollection_KeepReference = MethodBase + 2;
		internal const int TableRowCollection_GetFirst = MethodBase + 3;

		//===============================================================================
		// TableCell dispids
		//===============================================================================
		// Properties
		internal const int TableCell_Id = 1;
		internal const int TableCell_ReferenceId = 2;
		internal const int TableCell_ParentTable = 3;
		internal const int TableCell_ParentRow = 4;
		internal const int TableCell_RowIndex = 5;
		internal const int TableCell_CellIndex = 6;
		internal const int TableCell_Value = 7;
		internal const int TableCell_Body = 8;
		internal const int TableCell_ShadingColor = 9;
		internal const int TableCell_HorizontalAlignment = 13;
		internal const int TableCell_VerticalAlignment = 14;
		internal const int TableCell_ColumnWidth = 15;
		internal const int TableCell_Width = 16;
		// skip 17
		// Methods
		internal const int TableCell_OnAccess = MethodBase + 1;
		internal const int TableCell_KeepReference = MethodBase + 2;
		internal const int TableCell_InsertRows = MethodBase + 3;
		internal const int TableCell_InsertColumns = MethodBase + 4;
		internal const int TableCell_Split = MethodBase + 5;
		internal const int TableCell_DeleteRow = MethodBase + 6;
		internal const int TableCell_DeleteColumn = MethodBase + 7;
		internal const int TableCell_GetBorder = MethodBase + 8;
		internal const int TableCell_GetCellPadding = MethodBase + 9;
		internal const int TableCell_SetCellPadding = MethodBase + 10;
		internal const int TableCell_GetNext = MethodBase + 11;

		//===============================================================================
		// TableCellCollection dispids
		//===============================================================================
		// Properties
		internal const int TableCellCollection_Indexer = 1;
		internal const int TableCellCollection_ReferenceId = 2;
		// skip 3
		// Methods
		internal const int TableCellCollection_OnAccess = MethodBase + 1;
		internal const int TableCellCollection_KeepReference = MethodBase + 2;
		internal const int TableCellCollection_GetFirst = MethodBase + 3;

		//===============================================================================
		// TableBorder dispids
		//===============================================================================
		// Properties
		internal const int TableBorder_ReferenceId = 1;
		internal const int TableBorder_Color = 2;
		internal const int TableBorder_Type = 3;
		internal const int TableBorder_Width = 4;
		// Methods
		internal const int TableBorder_OnAccess = MethodBase + 1;
		internal const int TableBorder_KeepReference = MethodBase + 2;
	}

	// Internal Helper class for TypeScriptType strings literals.
	internal static class ParamTypeStrings
	{
		// SearchOptions param TypeScriptType in Search method.
		internal const string SearchOptions = "Word.SearchOptions | {"
			+ "ignorePunct?: boolean; ignoreSpace?: boolean; matchCase?: boolean; matchPrefix?: boolean;"
			+ "matchSuffix?: boolean; matchWholeWord?: boolean; matchWildcards?: boolean}";
	}

	/// <summary>
	/// The Application object.
	/// </summary>
	[ClientCallableComType(Name = "IApplication", InterfaceId = "E033F092-AC87-40F7-A865-769C9FDACD72",
		CoClassName = "Application", CoClassId = "2C000FC1-C973-41FC-B5E3-8F3A0CC65C12")]
	[ClientCallableServiceRoot]
	[ApiSet(Version = 1.3)]
	public interface Application
	{
		//===============================================================================
		// Properties
		//===============================================================================

		//===============================================================================
		// Methods
		//===============================================================================
		/// <summary>
		/// Creates a new document by using a base64 encoded .docx file.
		/// </summary>
		/// <param name="base64File">Optional. The base64 encoded .docx file. The default value is null.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Application_CreateDocument)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		Document CreateDocument([Optional] string base64File);
	}

	/// <summary>
	/// Represents the body of a document or a section.
	/// </summary>
	[ClientCallableComType(Name = "IBody", InterfaceId = "5516BB1D-CF31-4AA9-97CD-CCEBC2EE9E6D", CoClassName = "Body")]
	[ApiSet(Version = 1.1)]
	public interface Body
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_ReferenceId)]
		string _ReferenceId { get; }

		/// <summary>
		/// Gets the collection of paragraph objects in the body. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_Paragraphs)]
		[ApiSet(Version = 1.1)]
		ParagraphCollection Paragraphs { get; }

		/// <summary>
		/// Gets the collection of rich text content control objects in the body. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_ContentControls)]
		[ApiSet(Version = 1.1)]
		ContentControlCollection ContentControls { get; }

		/// <summary>
		/// Gets the content control that contains the body. Returns a null object if there isn't a parent content control. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_ParentContentControl)]
		[ApiSet(Version = 1.1)]
		ContentControl ParentContentControl { get; }

		/// <summary>
		/// Gets the text format of the body. Use this to get and set font name, size, color and other properties. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_Font)]
		[JsonStringify()]
		[ApiSet(Version = 1.1)]
		Font Font { get; }

		/// <summary>
		/// Gets or sets the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_Style)]
		[ApiSet(Version = 1.1)]
		string Style { get; set; }

		/// <summary>
		/// Gets or sets the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_StyleBuiltIn)]
		[ApiSet(Version = 1.3)]
		Style StyleBuiltIn { get; set; }

		/// <summary>
		/// Gets the collection of inlinePicture objects in the body. The collection does not include floating images. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_InlinePictures)]
		[ApiSet(Version = 1.1)]
		InlinePictureCollection InlinePictures { get; }

		/// <summary>
		/// Gets the text of the body. Use the insertText method to insert text. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_Text)]
		[ApiSet(Version = 1.1)]
		string Text { get; }

		/// <summary>
		/// Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_Type)]
		[ApiSet(Version = 1.3)]
		BodyType Type { get; }

		/// <summary>
		/// Gets the parent body of the body. For example, a table cell body's parent body could be a header. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_ParentBody)]
		[ApiSet(Version = 1.3)]
		Body ParentBody { get; }

		/// <summary>
		/// Gets the collection of list objects in the body. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_Lists)]
		[ApiSet(Version = 1.3)]
		ListCollection Lists { get; }

		/// <summary>
		/// Gets the collection of table objects in the body. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_Tables)]
		[ApiSet(Version = 1.3)]
		TableCollection Tables { get; }

		/// <summary>
		/// Gets the parent section of the body. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_ParentSection)]
		[ApiSet(Version = 1.3)]
		Section ParentSection { get; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.Body_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.Body_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		/// <summary>
		/// Inserts text into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
		/// </summary>
		/// <param name="text">Required. Text to be inserted.</param>
		/// <param name="insertLocation">Required. The value can be 'Replace', 'Start' or 'End'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_InsertText)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.1)]
		Range InsertText(string text, InsertLocation insertLocation);

		/// <summary>
		/// Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
		/// </summary>
		/// <param name="html">Required. The HTML to be inserted in the document.</param>
		/// <param name="insertLocation">Required. The value can be 'Replace', 'Start' or 'End'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_InsertHtml)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.1)]
		Range InsertHtml(string html, InsertLocation insertLocation);

		/// <summary>
		/// Inserts OOXML at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'.
		/// </summary>
		/// <param name="ooxml">Required. The OOXML to be inserted.</param>
		/// <param name="insertLocation">Required. The value can be 'Replace', 'Start' or 'End'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_InsertOoxml)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.1)]
		Range InsertOoxml(string ooxml, InsertLocation insertLocation);

		/// <summary>
		/// Inserts a paragraph at the specified location. The insertLocation value can be 'Start' or 'End'.
		/// </summary>
		/// <param name="paragraphText">Required. The paragraph text to be inserted.</param>
		/// <param name="insertLocation">Required. The value can be 'Start' or 'End'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_InsertParagraph)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.1)]
		Paragraph InsertParagraph(string paragraphText, InsertLocation insertLocation);

		/// <summary>
		/// Inserts a document into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
		/// </summary>
		/// <param name="base64File">Required. The base64 encoded content of a .docx file.</param>
		/// <param name="insertLocation">Required. The value can be 'Replace', 'Start' or 'End'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_InsertFileFromBase64)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.1)]
		Range InsertFileFromBase64(string base64File, InsertLocation insertLocation);

		/// <summary>
		/// Inserts a break at the specified location in the main document. The insertLocation value can be 'Start' or 'End'.
		/// </summary>
		/// <param name="breakType">Required. The break type to add to the body.</param>
		/// <param name="insertLocation">Required. The value can be 'Start' or 'End'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_InsertBreak)]
		[ClientCallableOperation(WacAsync = true)]
		[ApiSet(Version = 1.1)]
		void InsertBreak(BreakType breakType, InsertLocation insertLocation);

		/// <summary>
		/// Wraps the body object with a Rich Text content control.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_InsertContentControl)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.1)]
		ContentControl InsertContentControl();

		/// <summary>
		/// Inserts a picture into the body at the specified location. The insertLocation value can be 'Start' or 'End'.
		/// </summary>
		/// <param name="base64EncodedImage">Required. The base64 encoded image to be inserted in the body.</param>
		/// <param name="insertLocation">Required. The value can be 'Start' or 'End'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_InsertInlinePictureFromBase64)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.2)]
		InlinePicture InsertInlinePictureFromBase64(string base64EncodedImage, InsertLocation insertLocation);

		/// <summary>
		/// Clears the contents of the body object. The user can perform the undo operation on the cleared content.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_Clear)]
		[ClientCallableOperation(WacAsync = true)]
		[ApiSet(Version = 1.1)]
		void Clear();

		/// <summary>
		/// Selects the body and navigates the Word UI to it.
		/// </summary>
		/// <param name="selectionMode">Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_Select)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.1)]
		void Select([Optional] SelectionMode selectionMode);

		/// <summary>
		/// Performs a search with the specified searchOptions on the scope of the body object. The search results are a collection of range objects.
		/// </summary>
		/// <param name="searchText">Required. The search text.</param>
		/// <param name="searchOptions">Optional. Options for the search.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_Search)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.1)]
		RangeCollection Search(string searchText, [Optional][TypeScriptType(ParamTypeStrings.SearchOptions)]SearchOptions searchOptions);

		/// <summary>
		/// Gets the HTML representation of the body object.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_GetHtml)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.1)]
		string GetHtml();

		/// <summary>
		/// Gets the OOXML (Office Open XML) representation of the body object.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_GetOoxml)]
		[ClientCallableOperation(OperationType = OperationType.Read, WacAsync = true)]
		[ApiSet(Version = 1.1)]
		string GetOoxml();

		/// <summary>
		/// Gets the whole body, or the starting or ending point of the body, as a range.
		/// </summary>
		/// <param name="rangeLocation">Optional. The range location can be 'Whole', 'Start', 'End', 'After' or 'Content'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_GetRange)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		Range GetRange([Optional] RangeLocation rangeLocation);

		/// <summary>
		/// Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Start' or 'End'.
		/// </summary>
		/// <param name="rowCount">Required. The number of rows in the table.</param>
		/// <param name="columnCount">Required. The number of columns in the table.</param>
		/// <param name="insertLocation">Required. The value can be 'Start' or 'End'.</param>
		/// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Body_InsertTable)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.3)]
		Table InsertTable(int rowCount, int columnCount, InsertLocation insertLocation, [Optional] string[][] values);
	}

	/// <summary>
	/// Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
	/// </summary>
	[ClientCallableType(ExposeIsNullProperty = true)]
	[ClientCallableComType(Name = "IContentControl", InterfaceId = "b86e5ae1-476e-4e56-825d-885468e549f3", CoClassName = "ContentControl")]
	[ApiSet(Version = 1.1)]
	public interface ContentControl
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// Gets an integer that represents the content control identifier. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_Id)]
		[ApiSet(Version = 1.1)]
		uint Id { get; }

		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_ReferenceId)]
		string _ReferenceId { get; }

		/// <summary>
		/// Gets or sets the title for a content control.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_Title)]
		[ApiSet(Version = 1.1)]
		string Title { get; set; }

		/// <summary>
		/// Gets or sets a tag to identify a content control.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_Tag)]
		[ApiSet(Version = 1.1)]
		string Tag { get; set; }

		/// <summary>
		/// Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_PlaceholderText)]
		[ApiSet(Version = 1.1)]
		string PlaceholderText { get; set; }

		/// <summary>
		/// Gets the content control type. Only rich text content controls are supported currently. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_Type)]
		[ApiSet(Version = 1.1)]
		ContentControlType Type { get; }

		/// <summary>
		/// Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_Subtype)]
		[ApiSet(Version = 1.3)]
		ContentControlType Subtype { get; }

		/// <summary>
		/// Gets or sets the appearance of the content control. The value can be 'boundingBox', 'tags' or 'hidden'.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_Appearance)]
		[ApiSet(Version = 1.1)]
		ContentControlAppearance Appearance { get; set; }

		/// <summary>
		/// Gets or sets the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_Color)]
		[ApiSet(Version = 1.1)]
		string Color { get; set; }

		/// <summary>
		/// Gets or sets a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_RemoveWhenEdited)]
		[ApiSet(Version = 1.1)]
		bool RemoveWhenEdited { get; set; }

		/// <summary>
		/// Gets or sets a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_CannotDelete)]
		[ApiSet(Version = 1.1)]
		bool CannotDelete { get; set; }

		/// <summary>
		/// Gets or sets a value that indicates whether the user can edit the contents of the content control.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_CannotEdit)]
		[ApiSet(Version = 1.1)]
		bool CannotEdit { get; set; }

		/// <summary>
		/// Gets the text format of the content control. Use this to get and set font name, size, color, and other properties. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_Font)]
		[JsonStringify()]
		[ApiSet(Version = 1.1)]
		Font Font { get; }

		/// <summary>
		/// Gets or sets the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_Style)]
		[ApiSet(Version = 1.1)]
		string Style { get; set; }

		/// <summary>
		/// Gets or sets the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_StyleBuiltIn)]
		[ApiSet(Version = 1.3)]
		Style StyleBuiltIn { get; set; }

		/// <summary>
		/// Get the collection of paragraph objects in the content control. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_Paragraphs)]
		[ApiSet(Version = 1.1)]
		ParagraphCollection Paragraphs { get; }

		/// <summary>
		/// Gets the collection of content control objects in the content control. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_ContentControls)]
		[ApiSet(Version = 1.1)]
		ContentControlCollection ContentControls { get; }

		/// <summary>
		/// Gets the content control that contains the content control. Returns a null object if there isn't a parent content control. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_ParentContentControl)]
		[ApiSet(Version = 1.1)]
		ContentControl ParentContentControl { get; }

		/// <summary>
		/// Gets the collection of inlinePicture objects in the content control. The collection does not include floating images. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_InlinePictures)]
		[ApiSet(Version = 1.1)]
		InlinePictureCollection InlinePictures { get; }

		/// <summary>
		/// Gets the text of the content control. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_Text)]
		[ApiSet(Version = 1.1)]
		string Text { get; }

		/// <summary>
		/// Gets the collection of list objects in the content control. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_Lists)]
		[ApiSet(Version = 1.3)]
		ListCollection Lists { get; }

		/// <summary>
		/// Gets the collection of table objects in the content control. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_Tables)]
		[ApiSet(Version = 1.3)]
		TableCollection Tables { get; }

		/// <summary>
		/// Gets the table cell that contains the content control. Returns a null object if it is not contained in a table cell. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_ParentTableCell)]
		[ApiSet(Version = 1.3)]
		TableCell ParentTableCell { get; }

		/// <summary>
		/// Gets the table that contains the content control. Returns a null object if it is not contained in a table. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_ParentTable)]
		[ApiSet(Version = 1.3)]
		Table ParentTable { get; }

		/// <summary>
		/// Gets the parent body of the content control. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_ParentBody)]
		[ApiSet(Version = 1.3)]
		Body ParentBody { get; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		/// <summary>
		/// Inserts text into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
		/// </summary>
		/// <param name="text">Required. The text to be inserted in to the content control.</param>
		/// <param name="insertLocation">Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_InsertText)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.1)]
		Range InsertText(string text, InsertLocation insertLocation);

		/// <summary>
		/// Inserts HTML into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
		/// </summary>
		/// <param name="html">Required. The HTML to be inserted in to the content control.</param>
		/// <param name="insertLocation">Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_InsertHtml)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.1)]
		Range InsertHtml(string html, InsertLocation insertLocation);

		/// <summary>
		/// Inserts OOXML into the content control at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'.
		/// </summary>
		/// <param name="ooxml">Required. The OOXML to be inserted in to the content control.</param>
		/// <param name="insertLocation">Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_InsertOoxml)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.1)]
		Range InsertOoxml(string ooxml, InsertLocation insertLocation);

		/// <summary>
		/// Inserts a document into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
		/// </summary>
		/// <param name="base64File">Required. The base64 encoded content of a .docx file.</param>
		/// <param name="insertLocation">Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_InsertFileFromBase64)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.1)]
		Range InsertFileFromBase64(string base64File, InsertLocation insertLocation);

		/// <summary>
		/// Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.
		/// </summary>
		/// <param name="paragraphText">Required. The paragrph text to be inserted.</param>
		/// <param name="insertLocation">Required. The value can be 'Start', 'End', 'Before' or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_InsertParagraph)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.1)]
		Paragraph InsertParagraph(string paragraphText, InsertLocation insertLocation);

		/// <summary>
		/// Inserts a break at the specified location in the main document. The insertLocation value can be 'Start', 'End', 'Before' or 'After'. This method cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
		/// </summary>
		/// <param name="breakType">Required. Type of break.</param>
		/// <param name="insertLocation">Required. The value can be 'Start', 'End', 'Before' or 'After'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_InsertBreak)]
		[ClientCallableOperation(WacAsync = true)]
		[ApiSet(Version = 1.1)]
		void InsertBreak(BreakType breakType, InsertLocation insertLocation);

		/// <summary>
		/// Inserts an inline picture into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
		/// </summary>
		/// <param name="base64EncodedImage">Required. The base64 encoded image to be inserted in the content control.</param>
		/// <param name="insertLocation">Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_InsertInlinePictureFromBase64)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.2)]
		InlinePicture InsertInlinePictureFromBase64(string base64EncodedImage, InsertLocation insertLocation);

		/// <summary>
		/// Clears the contents of the content control. The user can perform the undo operation on the cleared content.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_Clear)]
		[ApiSet(Version = 1.1)]
		void Clear();

		/// <summary>
		/// Deletes the content control and its content. If keepContent is set to true, the content is not deleted.
		/// </summary>
		/// <param name="keepContent">Required. Indicates whether the content should be deleted with the content control. If keepContent is set to true, the content is not deleted.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_Delete)]
		[ClientCallableOperation(WacName = "DeleteContentControl")]
		[ApiSet(Version = 1.1)]
		void Delete(bool keepContent);

		/// <summary>
		/// Selects the content control. This causes Word to scroll to the selection.
		/// </summary>
		/// <param name="selectionMode">Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_Select)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.1)]
		void Select([Optional] SelectionMode selectionMode);

		/// <summary>
		/// Gets the HTML representation of the content control object.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_GetHtml)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.1)]
		string GetHtml();

		/// <summary>
		/// Gets the Office Open XML (OOXML) representation of the content control object.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_GetOoxml)]
		[ClientCallableOperation(OperationType = OperationType.Read, WacAsync = true)]
		[ApiSet(Version = 1.1)]
		string GetOoxml();

		/// <summary>
		/// Performs a search with the specified searchOptions on the scope of the content control object. The search results are a collection of range objects.
		/// </summary>
		/// <param name="searchText">Required. The search text.</param>
		/// <param name="searchOptions">Optional. Options for the search.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_Search)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.1)]
		RangeCollection Search(string searchText, [Optional][TypeScriptType(ParamTypeStrings.SearchOptions)]SearchOptions searchOptions);

		/// <summary>
		/// Gets the whole content control, or the starting or ending point of the content control, as a range.
		/// </summary>
		/// <param name="rangeLocation">Optional. The range location can be 'Whole', 'Before', 'Start', 'End', 'After' or 'Content'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_GetRange)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		Range GetRange([Optional] RangeLocation rangeLocation);

		/// <summary>
		/// Splits the content control into child ranges by using delimiters.
		/// </summary>
		/// <param name="delimiters">Required. The delimiters as an array of strings.</param>
		/// <param name="multiParagraphs">Optional. Indicates whether a returned child range can cover multiple paragraphs. Default is false which indicates that the paragraph boundaries are also used as delimiters.</param>
		/// <param name="trimDelimiters">Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.</param>
		/// <param name="trimSpacing">Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_Split)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		RangeCollection Split(string[] delimiters, [Optional] bool multiParagraphs, [Optional] bool trimDelimiters, [Optional] bool trimSpacing);

		/// <summary>
		/// Inserts a table with the specified number of rows and columns into, or next to, a content control. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.
		/// </summary>
		/// <param name="rowCount">Required. The number of rows in the table.</param>
		/// <param name="columnCount">Required. The number of columns in the table.</param>
		/// <param name="insertLocation">Required. The value can be 'Start', 'End', 'Before' or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.</param>
		/// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_InsertTable)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.3)]
		Table InsertTable(int rowCount, int columnCount, InsertLocation insertLocation, [Optional] string[][] values);

		/// <summary>
		/// Gets the text ranges in the content control by using punctuation marks and/or other ending marks.
		/// </summary>
		/// <param name="endingMarks">Required. The punctuation marks and/or other ending marks as an array of strings.</param>
		/// <param name="trimSpacing">Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControl_GetTextRanges)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		RangeCollection GetTextRanges(string[] endingMarks, [Optional] bool trimSpacing);
	}

	/// <summary>
	/// Contains a collection of [contentControl](contentControl.md) objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
	/// </summary>
	[ClientCallableComType(Name = "IContentControlCollection", InterfaceId = "a6a87ba8-6037-4625-ba4b-ece14e708219",
		CoClassName = "ContentControlCollection", SupportEnumeration = true, SupportIEnumVARIANT = true)]
	[ApiSet(Version = 1.1)]
	public interface ContentControlCollection : IEnumerable<ContentControl>
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// Gets a content control by its index in the collection.
		/// </summary>
		/// <param name="index">The index.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControlCollection_Indexer)]
		[ApiSet(Version = 1.1)]
		ContentControl this[[TypeScriptType("number")]object index] { get; }

		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControlCollection_ReferenceId)]
		string _ReferenceId { get; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControlCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.ContentControlCollection_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		/// <summary>
		/// Gets a content control by its identifier.
		/// </summary>
		/// <param name="id">Required. A content control identifier.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControlCollection_GetById)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.1)]
		ContentControl GetById(int id);

		/// <summary>
		/// Gets the content controls that have the specified title.
		/// </summary>
		/// <param name="title">Required. The title of a content control.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControlCollection_GetByTitle)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.1)]
		ContentControlCollection GetByTitle(string title);

		/// <summary>
		/// Gets the content controls that have the specified tag.
		/// </summary>
		/// <param name="tag">Required. A tag set on a content control.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControlCollection_GetByTag)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.1)]
		ContentControlCollection GetByTag(string tag);

		/// <summary>
		/// Gets the content controls that have the specified types and/or subtypes.
		/// </summary>
		/// <param name="types">Required. An array of content control types and/or subtypes.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControlCollection_GetByTypes)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		ContentControlCollection GetByTypes(ContentControlType[] types);

		/// <summary>
		/// Gets the first content control in this collection.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ContentControlCollection_GetFirst)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		ContentControl GetFirst();
	}

	/// <summary>
	/// Represents a custom property.
	/// </summary>
	[ClientCallableType(ExposeIsNullProperty = true)]
	[ClientCallableComType(Name = "ICustomProperty", InterfaceId = "5EDBD9C9-449F-440C-BFAE-D59D2DB340BB", CoClassName = "CustomProperty")]
	[ApiSet(Version = 1.3)]
	public interface CustomProperty
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.CustomProperty_ReferenceId)]
		string _ReferenceId { get; }

		/// <summary>
		/// Gets the key of the custom property. Read only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.CustomProperty_Key)]
		[ApiSet(Version = 1.3)]
		string Key { get; }

		/// <summary>
		/// Gets or sets the value of the custom property.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.CustomProperty_Value)]
		[ApiSet(Version = 1.3)]
		object Value { get; set; }

		/// <summary>
		/// Gets the value type of the custom property. Read only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.CustomProperty_Type)]
		[ApiSet(Version = 1.3)]
		DocumentPropertyType Type { get; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.CustomProperty_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.CustomProperty_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		/// <summary>
		/// Deletes the custom property.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.CustomProperty_Delete)]
		[ApiSet(Version = 1.3)]
		void Delete();
	}

	/// <summary>
	/// Contains the collection of [customProperty](customProperty.md) objects.
	/// </summary>
	[ClientCallableComType(Name = "ICustomPropertyCollection", InterfaceId = "85869847-EF5B-4377-9A0B-0EB6DA2059E5",
		CoClassName = "CustomPropertyCollection", SupportEnumeration = true, SupportIEnumVARIANT = true)]
	[ApiSet(Version = 1.3)]
	public interface CustomPropertyCollection : IEnumerable<CustomProperty>
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// Gets a custom property object by its key, which is case-insensitive.
		/// </summary>
		/// <param name="key">The key that identifies the custom property object.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.CustomPropertyCollection_Indexer)]
		[ApiSet(Version = 1.3)]
		CustomProperty this[string key] { get; }

		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.CustomPropertyCollection_ReferenceId)]
		string _ReferenceId { get; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.CustomPropertyCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.CustomPropertyCollection_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		/// <summary>
		/// Creates or sets a custom property.
		/// </summary>
		/// <param name="key">Required. The custom property's key, which is case-insensitive.</param>
		/// <param name="value">Required. The custom property's value.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.CustomPropertyCollection_Set)]
		[ApiSet(Version = 1.3)]
		CustomProperty Set(string key, object value);

		/// <summary>
		/// Gets the count of custom properties.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.CustomPropertyCollection_GetCount)]
		[ApiSet(Version = 1.3)]
		int GetCount();

		/// <summary>
		/// Deletes all custom properties in this collection.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.CustomPropertyCollection_DeleteAll)]
		[ApiSet(Version = 1.3)]
		void DeleteAll();
	}

	/// <summary>
	/// The Document object is the top level object. A Document object contains one or more sections, content controls, and the body that contains the contents of the document.
	/// </summary>
	[ClientCallableComType(Name = "IDocument", InterfaceId = "1fd33ae5-86f2-4770-95f1-8bba5b1dce15", CoClassName = "Document")]
	[ClientCallableServiceRoot]
	[ApiSet(Version = 1.1)]
	public interface Document
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Document_ReferenceId)]
		string _ReferenceId { get; }

		/// <summary>
		/// Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Document_Saved)]
		[ApiSet(Version = 1.1)]
		bool Saved { get; }

		/// <summary>
		/// Gets the collection of section objects in the document. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Document_Sections)]
		[ApiSet(Version = 1.1)]
		SectionCollection Sections { get; }

		/// <summary>
		/// Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Document_Body)]
		[JsonStringify()]
		[ApiSet(Version = 1.1)]
		Body Body { get; }

		/// <summary>
		/// Gets the collection of content control objects in the current document. This includes content controls in the body of the document, headers, footers, textboxes, etc.. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Document_ContentControls)]
		[ApiSet(Version = 1.1)]
		ContentControlCollection ContentControls { get; }

		/// <summary>
		/// Gets the properties of the current document. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Document_Properties)]
		[JsonStringify()]
		[ApiSet(Version = 1.3)]
		DocumentProperties Properties { get; }

		/// <summary>
		/// Gets the add-in's settings in the current document. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Document_Settings)]
		[ApiSet(Version = 1.4)]
		SettingCollection Settings { get; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.Document_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.Document_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		[ClientCallableComMember(DispatchId = DispatchIds.Document_GetObjectByReferenceId)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		object _GetObjectByReferenceId(string referenceId);

		[ClientCallableComMember(DispatchId = DispatchIds.Document_GetObjectTypeNameByReferenceId)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		string _GetObjectTypeNameByReferenceId(string referenceId);

		[ClientCallableComMember(DispatchId = DispatchIds.Document_RemoveReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _RemoveReference(string referenceId);

		[ClientCallableComMember(DispatchId = DispatchIds.Document_RemoveAllReferences)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _RemoveAllReferences();

		/// <summary>
		/// Saves the document. This will use the Word default file naming convention if the document has not been saved before.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Document_Save)]
		[ApiSet(Version = 1.1)]
		void Save();

		/// <summary>
		/// Gets the current selection of the document. Multiple selections are not supported.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Document_GetSelection)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.1)]
		Range GetSelection();

		/// <summary>
		/// Open the document.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Document_Open)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		void Open();

		/// <summary>
		/// Gets a bookmark's range. Returns a null object if the bookmark does not exist.
		/// </summary>
		/// <param name="name">Required. The bookmark name, which is case-insensitive.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Document_GetBookmarkRange)]
		[ApiSet(Version = 1.4)]
		Range GetBookmarkRange(string name);

		/// <summary>
		/// Deletes a bookmark, if exists, from this document.
		/// </summary>
		/// <param name="name">Required. The bookmark name, which is case-insensitive.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Document_DeleteBookmark)]
		[ApiSet(Version = 1.4)]
		void DeleteBookmark(string name);
	}

	/// <summary>
	/// Represents document properties.
	/// </summary>
	[ClientCallableComType(Name = "IDocumentProperties", InterfaceId = "642E1471-0CAC-4DD2-BC76-1D1D0A3E130F", CoClassName = "DocumentProperties")]
	[ApiSet(Version = 1.3)]
	public interface DocumentProperties
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.DocumentProperties_ReferenceId)]
		string _ReferenceId { get; }

		/// <summary>
		/// Gets the collection of custom properties of the document. Read only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.DocumentProperties_CustomProperties)]
		[ApiSet(Version = 1.3)]
		CustomPropertyCollection CustomProperties { get; }

		/// <summary>
		/// Gets or sets the title of the document.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.DocumentProperties_Title)]
		[ApiSet(Version = 1.3)]
		string Title { get; set; }

		/// <summary>
		/// Gets or sets the subject of the document.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.DocumentProperties_Subject)]
		[ApiSet(Version = 1.3)]
		string Subject { get; set; }

		/// <summary>
		/// Gets or sets the author of the document.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.DocumentProperties_Author)]
		[ApiSet(Version = 1.3)]
		string Author { get; set; }

		/// <summary>
		/// Gets or sets the keywords of the document.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.DocumentProperties_Keywords)]
		[ApiSet(Version = 1.3)]
		string Keywords { get; set; }

		/// <summary>
		/// Gets or sets the comments of the document.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.DocumentProperties_Comments)]
		[ApiSet(Version = 1.3)]
		string Comments { get; set; }

		/// <summary>
		/// Gets the template of the document. Read only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.DocumentProperties_Template)]
		[ApiSet(Version = 1.3)]
		string Template { get; }

		/// <summary>
		/// Gets or sets the last author of the document.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.DocumentProperties_LastAuthor)]
		[ApiSet(Version = 1.3)]
		string LastAuthor { get; set; }

		/// <summary>
		/// Gets the revision number of the document. Read only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.DocumentProperties_RevisionNumber)]
		[ApiSet(Version = 1.3)]
		string RevisionNumber { get; }

		/// <summary>
		/// Gets the application name of the document. Read only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.DocumentProperties_ApplicationName)]
		[ApiSet(Version = 1.3)]
		string ApplicationName { get; }

		/// <summary>
		/// Gets the last print date of the document. Read only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.DocumentProperties_LastPrintDate)]
		[ApiSet(Version = 1.3)]
		DateTime? LastPrintDate { get; }

		/// <summary>
		/// Gets the creation date of the document. Read only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.DocumentProperties_CreationDate)]
		[ApiSet(Version = 1.3)]
		DateTime CreationDate { get; }

		/// <summary>
		/// Gets the last save time of the document. Read only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.DocumentProperties_LastSaveTime)]
		[ApiSet(Version = 1.3)]
		DateTime? LastSaveTime { get; }

		/// <summary>
		/// Gets the security of the document. Read only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.DocumentProperties_Security)]
		[ApiSet(Version = 1.3)]
		int? Security { get; }

		/// <summary>
		/// Gets or sets the category of the document.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.DocumentProperties_Category)]
		[ApiSet(Version = 1.3)]
		string Category { get; set; }

		/// <summary>
		/// Gets or sets the format of the document.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.DocumentProperties_Format)]
		[ApiSet(Version = 1.3)]
		string Format { get; set; }

		/// <summary>
		/// Gets or sets the manager of the document.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.DocumentProperties_Manager)]
		[ApiSet(Version = 1.3)]
		string Manager { get; set; }

		/// <summary>
		/// Gets or sets the company of the document.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.DocumentProperties_Company)]
		[ApiSet(Version = 1.3)]
		string Company { get; set; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.DocumentProperties_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.DocumentProperties_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();
	}

	/// <summary>
	/// Represents a font.
	/// </summary>
	[ClientCallableComType(Name = "IFont", InterfaceId = "896A961F-E647-4B59-82DA-5F1143B5AFC1", CoClassName = "Font")]
	[ApiSet(Version = 1.1)]
	public interface Font
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Font_ReferenceId)]
		string _ReferenceId { get; }

		/// <summary>
		/// Gets or sets a value that represents the name of the font.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Font_Name)]
		[ApiSet(Version = 1.1)]
		string Name { get; set; }

		/// <summary>
		/// Gets or sets a value that represents the font size in points.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Font_Size)]
		[ApiSet(Version = 1.1)]
		float? Size { get; set; }

		/// <summary>
		/// Gets or sets a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Font_Bold)]
		[ApiSet(Version = 1.1)]
		bool? Bold { get; set; }

		/// <summary>
		/// Gets or sets a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Font_Italic)]
		[ApiSet(Version = 1.1)]
		bool? Italic { get; set; }

		/// <summary>
		/// Gets or sets the color for the specified font. You can provide the value in the '#RRGGBB' format or the color name.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Font_Color)]
		[ApiSet(Version = 1.1)]
		string Color { get; set; }

		/// <summary>
		/// Gets or sets a value that indicates the font's underline type. 'None' if the font is not underlined.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Font_Underline)]
		[ApiSet(Version = 1.1)]
		UnderlineType Underline { get; set; }

		/// <summary>
		/// Gets or sets a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Font_Subscript)]
		[ApiSet(Version = 1.1)]
		bool? Subscript { get; set; }

		/// <summary>
		/// Gets or sets a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Font_Superscript)]
		[ApiSet(Version = 1.1)]
		bool? Superscript { get; set; }

		/// <summary>
		/// Gets or sets a value that indicates whether the font has a strike through. True if the font is formatted as strikethrough text, otherwise, false.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Font_StrikeThrough)]
		[ApiSet(Version = 1.1)]
		bool? StrikeThrough { get; set; }

		/// <summary>
		/// Gets or sets a value that indicates whether the font has a double strike through. True if the font is formatted as double strikethrough text, otherwise, false.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Font_DoubleStrikeThrough)]
		[ApiSet(Version = 1.1)]
		bool? DoubleStrikeThrough { get; set; }

		/// <summary>
		/// Gets or sets the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to null. The returned highlight color can be in the '#RRGGBB' format, or an empty string for mixed highlight colors, or null for no highlight color.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Font_HighlightColor)]
		[ApiSet(Version = 1.1)]
		string HighlightColor { get; set; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.Font_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.Font_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();
	}

	/// <summary>
	/// Represents an inline picture.
	/// </summary>
	[ClientCallableType(ExposeIsNullProperty = true)]
	[ClientCallableComType(Name = "IInlinePicture", InterfaceId = "EF524552-BE9E-413B-8030-E5EB3DB24F19", CoClassName = "InlinePicture")]
	[ApiSet(Version = 1.1)]
	public interface InlinePicture
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// ID
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_Id)]
		int _Id { get; }

		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_ReferenceId)]
		string _ReferenceId { get; }

		/// <summary>
		/// Gets or sets a string that represents the alternative text associated with the inline image
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_AltTextDescription)]
		[ApiSet(Version = 1.1)]
		string AltTextDescription { get; set; }

		/// <summary>
		/// Gets or sets a string that contains the title for the inline image.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_AltTextTitle)]
		[ApiSet(Version = 1.1)]
		string AltTextTitle { get; set; }

		/// <summary>
		/// Gets or sets a number that describes the height of the inline image.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_Height)]
		[ApiSet(Version = 1.1)]
		float Height { get; set; }

		/// <summary>
		/// Gets or sets a hyperlink on the image. Use a newline character ('\n') to separate the address part from the optional location part.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_Hyperlink)]
		[ApiSet(Version = 1.1)]
		string Hyperlink { get; set; }

		/// <summary>
		/// Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_LockAspectRatio)]
		[ApiSet(Version = 1.1)]
		bool LockAspectRatio { get; set; }

		/// <summary>
		/// Gets or sets a number that describes the width of the inline image.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_Width)]
		[ApiSet(Version = 1.1)]
		float Width { get; set; }

		/// <summary>
		/// Gets the content control that contains the inline image. Returns a null object if there isn't a parent content control. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_ParentContentControl)]
		[ApiSet(Version = 1.1)]
		ContentControl ParentContentControl { get; }

		/// <summary>
		/// Gets the parent paragraph that contains the inline image. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_Paragraph)]
		[ApiSet(Version = 1.2)]
		Paragraph Paragraph { get; }

		/// <summary>
		/// Gets the format of the inline image. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_ImageFormat)]
		[ApiSet(Version = 1.4)]
		ImageFormat ImageFormat { get; }

		/// <summary>
		/// Gets the table cell that contains the inline image. Returns a null object if it is not contained in a table cell. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_ParentTableCell)]
		[ApiSet(Version = 1.3)]
		TableCell ParentTableCell { get; }

		/// <summary>
		/// Gets the table that contains the inline image. Returns a null object if it is not contained in a table. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_ParentTable)]
		[ApiSet(Version = 1.3)]
		Table ParentTable { get; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		/// <summary>
		/// Gets the base64 encoded string representation of the inline image.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_GetBase64ImageSrc)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.1)]
		string GetBase64ImageSrc();

		/// <summary>
		/// Wraps the inline picture with a rich text content control.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_InsertContentControl)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.1)]
		ContentControl InsertContentControl();

		/// <summary>
		/// Inserts an inline picture at the specified location. The insertLocation value can be 'Replace', 'Before' or 'After'.
		/// </summary>
		/// <param name="base64EncodedImage">Required. The base64 encoded image to be inserted.</param>
		/// <param name="insertLocation">Required. The value can be 'Replace', 'Before' or 'After'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_InsertInlinePictureFromBase64)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.2)]
		InlinePicture InsertInlinePictureFromBase64(string base64EncodedImage, InsertLocation insertLocation);

		/// <summary>
		/// Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.
		/// </summary>
		/// <param name="breakType">Required. The break type to add.</param>
		/// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_InsertBreak)]
		[ClientCallableOperation(WacAsync = true)]
		[ApiSet(Version = 1.2)]
		void InsertBreak(BreakType breakType, InsertLocation insertLocation);

		/// <summary>
		/// Inserts text at the specified location. The insertLocation value can be 'Before' or 'After'.
		/// </summary>
		/// <param name="text">Required. Text to be inserted.</param>
		/// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_InsertText)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.2)]
		Range InsertText(string text, InsertLocation insertLocation);

		/// <summary>
		/// Inserts HTML at the specified location. The insertLocation value can be 'Before' or 'After'.
		/// </summary>
		/// <param name="html">Required. The HTML to be inserted.</param>
		/// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_InsertHtml)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.2)]
		Range InsertHtml(string html, InsertLocation insertLocation);

		/// <summary>
		/// Inserts OOXML at the specified location.  The insertLocation value can be 'Before' or 'After'.
		/// </summary>
		/// <param name="ooxml">Required. The OOXML to be inserted.</param>
		/// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_InsertOoxml)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.2)]
		Range InsertOoxml(string ooxml, InsertLocation insertLocation);

		/// <summary>
		/// Inserts a document at the specified location. The insertLocation value can be 'Before' or 'After'.
		/// </summary>
		/// <param name="base64File">Required. The base64 encoded content of a .docx file.</param>
		/// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_InsertFileFromBase64)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.2)]
		Range InsertFileFromBase64(string base64File, InsertLocation insertLocation);

		/// <summary>
		/// Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.
		/// </summary>
		/// <param name="paragraphText">Required. The paragraph text to be inserted.</param>
		/// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_InsertParagraph)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.2)]
		Paragraph InsertParagraph(string paragraphText, InsertLocation insertLocation);

		/// <summary>
		/// Deletes the inline picture from the document.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_Delete)]
		[ClientCallableOperation(WacName = "DeleteInlinePicture")]
		[ApiSet(Version = 1.2)]
		void Delete();

		/// <summary>
		/// Selects the inline picture. This causes Word to scroll to the selection.
		/// </summary>
		/// <param name="selectionMode">Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_Select)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.2)]
		void Select([Optional] SelectionMode selectionMode);

		/// <summary>
		/// Gets the picture, or the starting or ending point of the picture, as a range.
		/// </summary>
		/// <param name="rangeLocation">Optional. The range location can be 'Whole', 'Start' or 'End'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_GetRange)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		Range GetRange([Optional] RangeLocation rangeLocation);

		/// <summary>
		/// Gets the next inline image.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePicture_GetNext)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		InlinePicture GetNext();
	}

	/// <summary>
	/// Contains a collection of [inlinePicture](inlinePicture.md) objects.
	/// </summary>
	[ClientCallableType(HiddenIndexerMethod = true)]
	[ClientCallableComType(Name = "IInlinePictureCollection", InterfaceId = "EA74232B-E74B-4A8B-9D30-04D181C29CFA",
		CoClassName = "InlinePictureCollection", SupportEnumeration = true)]
	[ApiSet(Version = 1.1)]
	public interface InlinePictureCollection : IEnumerable<InlinePicture>
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// Gets an inline picture object by its index in the collection.
		/// </summary>
		/// <param name="index">A number that identifies the index location of an inline picture object.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePictureCollection_Indexer)]
		[ApiSet(Version = 1.1)]
		InlinePicture this[[TypeScriptType("number")]object index] { get; }

		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePictureCollection_ReferenceId)]
		string _ReferenceId { get; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePictureCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.InlinePictureCollection_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		/// <summary>
		/// Gets the first inline image in this collection.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.InlinePictureCollection_GetFirst)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		InlinePicture GetFirst();
	}

	/// <summary>
	/// Contains a collection of [paragraph](paragraph.md) objects.
	/// </summary>
	[ClientCallableType(ExposeIsNullProperty = true)]
	[ClientCallableComType(Name = "IList", InterfaceId = "92CB59A8-B889-4985-840C-42B2F740EF33", CoClassName = "List")]
	[ApiSet(Version = 1.3)]
	public interface List
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// Gets the list's id.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.List_Id)]
		[ApiSet(Version = 1.3)]
		int Id { get; }

		[ClientCallableComMember(DispatchId = DispatchIds.List_ReferenceId)]
		string _ReferenceId { get; }

		/// <summary>
		/// Gets paragraphs in the list. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.List_Paragraphs)]
		[ApiSet(Version = 1.3)]
		ParagraphCollection Paragraphs { get; }

		/// <summary>
		/// Gets all 9 level types in the list. Each type can be 'Bullet', 'Number' or 'Picture'. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.List_LevelTypes)]
		[ApiSet(Version = 1.3)]
		ListLevelType[] LevelTypes { get; }

		/// <summary>
		/// Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.List_LevelExistences)]
		[ApiSet(Version = 1.3)]
		bool[] LevelExistences { get; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.List_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		[ClientCallableComMember(DispatchId = DispatchIds.List_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.
		/// </summary>
		/// <param name="paragraphText">Required. The paragraph text to be inserted.</param>
		/// <param name="insertLocation">Required. The value can be 'Start', 'End', 'Before' or 'After'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.List_InsertParagraph)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.3)]
		Paragraph InsertParagraph(string paragraphText, InsertLocation insertLocation);

		/// <summary>
		/// Gets the paragraphs that occur at the specified level in the list.
		/// </summary>
		/// <param name="level">Required. The level in the list.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.List_GetLevelParagraphs)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		ParagraphCollection GetLevelParagraphs(int level);

		/// <summary>
		/// Sets the bullet format at the specified level in the list. If the bullet is 'Custom', the charCode is required.
		/// </summary>
		/// <param name="level">Required. The level in the list.</param>
		/// <param name="listBullet">Required. The bullet.</param>
		/// <param name="charCode">Optional. The bullet character's code value. Used only if the bullet is 'Custom'.</param>
		/// <param name="fontName">Optional. The bullet's font name. Used only if the bullet is 'Custom'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.List_SetLevelBullet)]
		[ApiSet(Version = 1.3)]
		void SetLevelBullet(int level, ListBullet listBullet, [Optional] int charCode, [Optional] string fontName);

		/// <summary>
		/// Sets the numbering format at the specified level in the list.
		/// </summary>
		/// <param name="level">Required. The level in the list.</param>
		/// <param name="listNumbering">Required. The ordinal format.</param>
		/// <param name="formatString">Optional. The numbering string format defined as an array of strings and/or integers. Each integer is a level of number type that is higher than or equal to this level. For example, an array of ["(", level - 1, ".", level, ")"] can define the format of "(2.c)", where 2 is the parent's item number and c is this level's item number.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.List_SetLevelNumbering)]
		[ApiSet(Version = 1.3)]
		void SetLevelNumbering(int level, ListNumbering listNumbering, [Optional] object[] formatString);

		/// <summary>
		/// Gets the bullet, number or picture at the specified level as a string.
		/// </summary>
		/// <param name="level">Required. The level in the list.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.List_GetLevelString)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		string GetLevelString(int level);

		/// <summary>
		/// Sets the picture at the specified level in the list.
		/// </summary>
		/// <param name="level">Required. The level in the list.</param>
		/// <param name="base64EncodedImage">Optional. The base64 encoded image to be set. If not given, the default picture is set.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.List_SetLevelPicture)]
		[ApiSet(Version = 1.4)]
		void SetLevelPicture(int level, [Optional] string base64EncodedImage);

		/// <summary>
		/// Gets the base64 encoded string representation of the picture at the specified level in the list.
		/// </summary>
		/// <param name="level">Required. The level in the list.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.List_GetLevelPicture)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.4)]
		string GetLevelPicture(int level);

		/// <summary>
		/// Gets the font of the bullet, number or picture at the specified level in the list.
		/// </summary>
		/// <param name="level">Required. The level in the list.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.List_GetLevelFont)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.4)]
		Font GetLevelFont(int level);

		/// <summary>
		/// Resets the font of the bullet, number or picture at the specified level in the list.
		/// </summary>
		/// <param name="level">Required. The level in the list.</param>
		/// <param name="resetFontName">Optional. Indicates whether to reset the font name. Default is false that indicates the font name is kept unchanged.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.List_ResetLevelFont)]
		[ApiSet(Version = 1.4)]
		void ResetLevelFont(int level, [Optional] bool resetFontName);

		/// <summary>
		/// Sets the alignment of the bullet, number or picture at the specified level in the list.
		/// </summary>
		/// <param name="level">Required. The level in the list.</param>
		/// <param name="alignment">Required. The level alignment that can be 'left', 'centered' or 'right'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.List_SetLevelAlignment)]
		[ApiSet(Version = 1.3)]
		void SetLevelAlignment(int level, Alignment alignment);

		/// <summary>
		/// Sets the two indents of the specified level in the list.
		/// </summary>
		/// <param name="level">Required. The level in the list.</param>
		/// <param name="textIndent">Required. The text indent in points. It is the same as paragraph left indent.</param>
		/// <param name="textIndent">Required. The relative indent, in points, of the bullet, number or picture. It is the same as paragraph first line indent.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.List_SetLevelIndents)]
		[ApiSet(Version = 1.3)]
		void SetLevelIndents(int level, float textIndent, float bulletNumberPictureIndent);

		/// <summary>
		/// Sets the starting number at the specified level in the list. Default value is 1.
		/// </summary>
		/// <param name="level">Required. The level in the list.</param>
		/// <param name="startingNumber">Required. The number to start with.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.List_SetLevelStartingNumber)]
		[ApiSet(Version = 1.3)]
		void SetLevelStartingNumber(int level, int startingNumber);
	}

	/// <summary>
	/// Contains a collection of [list](list.md) objects.
	/// </summary>
	[ClientCallableComType(Name = "IListCollection", InterfaceId = "D36F169F-D0B0-44E9-A933-62C4DCB1964F",
		CoClassName = "ListCollection", SupportEnumeration = true, SupportIEnumVARIANT = true)]
	[ApiSet(Version = 1.3)]
	public interface ListCollection : IEnumerable<List>
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// Gets a list object by its index in the collection.
		/// </summary>
		/// <param name="index">A number that identifies the index location of a list object.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ListCollection_Indexer)]
		[ApiSet(Version = 1.3)]
		List this[[TypeScriptType("number")]object index] { get; }

		[ClientCallableComMember(DispatchId = DispatchIds.ListCollection_ReferenceId)]
		string _ReferenceId { get; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.ListCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.ListCollection_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		/// <summary>
		/// Gets a list by its identifier.
		/// </summary>
		/// <param name="id">Required. A list identifier.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ListCollection_GetById)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		List GetById(int id);

		/// <summary>
		/// Gets the first list in this collection.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ListCollection_GetFirst)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		List GetFirst();
	}

	/// <summary>
	/// Represents the paragraph list item format.
	/// </summary>
	[ClientCallableComType(Name = "IListItem", InterfaceId = "F32748C0-D8DD-44AE-A550-92ED26A085A8", CoClassName = "ListItem")]
	[ApiSet(Version = 1.3)]
	public interface ListItem
	{
		//===============================================================================
		// Properties
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.ListItem_ReferenceId)]
		string _ReferenceId { get; }

		/// <summary>
		/// Gets the list item order number in relation to its siblings. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ListItem_SiblingIndex)]
		[ApiSet(Version = 1.3)]
		int SiblingIndex { get; }

		/// <summary>
		/// Gets the list item bullet, number or picture as a string. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ListItem_ListString)]
		[ApiSet(Version = 1.3)]
		string ListString { get; }

		/// <summary>
		/// Gets or sets the level of the item in the list.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ListItem_Level)]
		[ApiSet(Version = 1.3)]
		int Level { get; set; }

		//===============================================================================
		// Methods
		//===============================================================================

		[ClientCallableComMember(DispatchId = DispatchIds.ListItem_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.ListItem_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		/// <summary>
		/// Gets the list item parent, or the closest ancestor if the parent does not exist.
		/// </summary>
		/// <param name="parentOnly">Optional. Specified only the list item's parent will be returned. The default is false that specifies to get the lowest ancestor.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ListItem_GetAncestor)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		Paragraph GetAncestor([Optional] bool parentOnly);

		/// <summary>
		/// Gets all descendant list items of the list item.
		/// </summary>
		/// <param name="directChildrenOnly">Optional. Specified only the list item's direct children will be returned. The default is false that indicates to get all descendant items.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ListItem_GetDescendants)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		ParagraphCollection GetDescendants([Optional] bool directChildrenOnly);
	}

	/// <summary>
	/// Represents a single paragraph in a selection, range, content control, or document body.
	/// </summary>
	[ClientCallableType(ExposeIsNullProperty = true)]
	[ClientCallableComType(Name = "IParagraph", InterfaceId = "D0C9772C-BECF-4961-A582-8024EBDC614C", CoClassName = "Paragraph")]
	[ApiSet(Version = 1.1)]
	public interface Paragraph
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// ID
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_Id)]
		int _Id { get; }

		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_ReferenceId)]
		string _ReferenceId { get; }

		/// <summary>
		/// Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_Font)]
		[JsonStringify()]
		[ApiSet(Version = 1.1)]
		Font Font { get; }

		/// <summary>
		/// Gets or sets the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_Style)]
		[ApiSet(Version = 1.1)]
		string Style { get; set; }

		/// <summary>
		/// Gets or sets the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_StyleBuiltIn)]
		[ApiSet(Version = 1.3)]
		Style StyleBuiltIn { get; set; }

		/// <summary>
		/// Gets the collection of content control objects in the paragraph. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_ContentControls)]
		[ApiSet(Version = 1.1)]
		ContentControlCollection ContentControls { get; }

		/// <summary>
		/// Gets the content control that contains the paragraph. Returns a null object if there isn't a parent content control. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_ParentContentControl)]
		[ApiSet(Version = 1.1)]
		ContentControl ParentContentControl { get; }

		/// <summary>
		/// Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_FirstLineIndent)]
		[ApiSet(Version = 1.1)]
		float FirstLineIndent { get; set; }

		/// <summary>
		/// Gets or sets the left indent value, in points, for the paragraph.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_LeftIndent)]
		[ApiSet(Version = 1.1)]
		float LeftIndent { get; set; }

		/// <summary>
		/// Gets or sets the right indent value, in points, for the paragraph.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_RightIndent)]
		[ApiSet(Version = 1.1)]
		float RightIndent { get; set; }

		/// <summary>
		/// Gets or sets the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_Alignment)]
		[ApiSet(Version = 1.1)]
		Alignment Alignment { get; set; }

		/// <summary>
		/// Gets or sets the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_LineSpacing)]
		[ApiSet(Version = 1.1)]
		float LineSpacing { get; set; }

		/// <summary>
		/// Gets or sets the outline level for the paragraph.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_OutlineLevel)]
		[ApiSet(Version = 1.1)]
		int OutlineLevel { get; set; }

		/// <summary>
		/// Gets or sets the spacing, in points, before the paragraph.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_SpaceBefore)]
		[ApiSet(Version = 1.1)]
		float SpaceBefore { get; set; }

		/// <summary>
		/// Gets or sets the spacing, in points, after the paragraph.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_SpaceAfter)]
		[ApiSet(Version = 1.1)]
		float SpaceAfter { get; set; }

		/// <summary>
		/// Gets or sets the amount of spacing, in grid lines, before the paragraph.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_LineUnitBefore)]
		[ApiSet(Version = 1.1)]
		float LineUnitBefore { get; set; }

		/// <summary>
		/// Gets or sets the amount of spacing, in grid lines. after the paragraph.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_LineUnitAfter)]
		[ApiSet(Version = 1.1)]
		float LineUnitAfter { get; set; }

		/// <summary>
		/// Gets the collection of inlinePicture objects in the paragraph. The collection does not include floating images. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_InlinePictures)]
		[ApiSet(Version = 1.1)]
		InlinePictureCollection InlinePictures { get; }

		/// <summary>
		/// Gets the text of the paragraph. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_Text)]
		[ApiSet(Version = 1.1)]
		string Text { get; }

		/// <summary>
		/// Checks whether the paragraph is a list item. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_IsListItem)]
		[ApiSet(Version = 1.3)]
		bool IsListItem { get; }

		/// <summary>
		/// Gets the level of the paragraph's table. It returns 0 if the paragraph is not in a table. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_TableNestingLevel)]
		[ApiSet(Version = 1.3)]
		int TableNestingLevel { get; }

		/// <summary>
		/// Gets the parent body of the paragraph. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_ParentBody)]
		[ApiSet(Version = 1.3)]
		Body ParentBody { get; }

		/// <summary>
		/// Gets the List to which this paragraph belongs. Returns a null object if the paragraph is not in a list. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_List)]
		[ApiSet(Version = 1.3)]
		List List { get; }

		/// <summary>
		/// Gets the table cell that contains the paragraph. Returns a null object if it is not contained in a table cell. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_ParentTableCell)]
		[ApiSet(Version = 1.3)]
		TableCell ParentTableCell { get; }

		/// <summary>
		/// Gets the table that contains the paragraph. Returns a null object if it is not contained in a table. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_ParentTable)]
		[ApiSet(Version = 1.3)]
		Table ParentTable { get; }

		/// <summary>
		/// Gets the ListItem for the paragraph. Returns a null object if the paragraph is not part of a list. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_ListItem)]
		[JsonStringify()]
		[ApiSet(Version = 1.3)]
		ListItem ListItem { get; }

		/// <summary>
		/// Indicates the paragraph is the last one inside its parent body. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_IsLastParagraph)]
		[ApiSet(Version = 1.3)]
		bool IsLastParagraph { get; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		/// <summary>
		/// Inserts a picture into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
		/// </summary>
		/// <param name="base64EncodedImage">Required. The base64 encoded image to be inserted.</param>
		/// <param name="insertLocation">Required. The value can be 'Replace', 'Start' or 'End'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_InsertInlinePictureFromBase64)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.1)]
		InlinePicture InsertInlinePictureFromBase64(string base64EncodedImage, InsertLocation insertLocation);

		/// <summary>
		/// Wraps the paragraph object with a rich text content control.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_InsertContentControl)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.1)]
		ContentControl InsertContentControl();

		/// <summary>
		/// Inserts text into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
		/// </summary>
		/// <param name="text">Required. Text to be inserted.</param>
		/// <param name="insertLocation">Required. The value can be 'Replace', 'Start' or 'End'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_InsertText)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.1)]
		Range InsertText(string text, InsertLocation insertLocation);

		/// <summary>
		/// Inserts HTML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
		/// </summary>
		/// <param name="html">Required. The HTML to be inserted in the paragraph.</param>
		/// <param name="insertLocation">Required. The value can be 'Replace', 'Start' or 'End'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_InsertHtml)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.1)]
		Range InsertHtml(string html, InsertLocation insertLocation);

		/// <summary>
		/// Inserts OOXML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
		/// </summary>
		/// <param name="ooxml">Required. The OOXML to be inserted in the paragraph.</param>
		/// <param name="insertLocation">Required. The value can be 'Replace', 'Start' or 'End'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_InsertOoxml)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.1)]
		Range InsertOoxml(string ooxml, InsertLocation insertLocation);

		/// <summary>
		/// Inserts a document into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
		/// </summary>
		/// <param name="base64File">Required. The base64 encoded content of a .docx file.</param>
		/// <param name="insertLocation">Required. The value can be 'Replace', 'Start' or 'End'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_InsertFileFromBase64)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.1)]
		Range InsertFileFromBase64(string base64File, InsertLocation insertLocation);

		/// <summary>
		/// Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.
		/// </summary>
		/// <param name="paragraphText">Required. The paragraph text to be inserted.</param>
		/// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_InsertParagraph)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.1)]
		Paragraph InsertParagraph(string paragraphText, InsertLocation insertLocation);

		/// <summary>
		/// Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.
		/// </summary>
		/// <param name="breakType">Required. The break type to add to the document.</param>
		/// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_InsertBreak)]
		[ClientCallableOperation(WacAsync = true)]
		[ApiSet(Version = 1.1)]
		void InsertBreak(BreakType breakType, InsertLocation insertLocation);

		/// <summary>
		/// Clears the contents of the paragraph object. The user can perform the undo operation on the cleared content.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_Clear)]
		[ApiSet(Version = 1.1)]
		void Clear();

		/// <summary>
		/// Deletes the paragraph and its content from the document.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_Delete)]
		[ClientCallableOperation(WacName = "DeleteParagraph")]
		[ApiSet(Version = 1.1)]
		void Delete();

		/// <summary>
		/// Selects and navigates the Word UI to the paragraph.
		/// </summary>
		/// <param name="selectionMode">Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_Select)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.1)]
		void Select([Optional] SelectionMode selectionMode);

		/// <summary>
		/// Gets the HTML representation of the paragraph object.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_GetHtml)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.1)]
		string GetHtml();

		/// <summary>
		/// Gets the Office Open XML (OOXML) representation of the paragraph object.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_GetOoxml)]
		[ClientCallableOperation(OperationType = OperationType.Read, WacAsync = true)]
		[ApiSet(Version = 1.1)]
		string GetOoxml();

		/// <summary>
		/// Performs a search with the specified searchOptions on the scope of the paragraph object. The search results are a collection of range objects.
		/// </summary>
		/// <param name="searchText">Required. The search text.</param>
		/// <param name="searchOptions">Optional. Options for the search.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_Search)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.1)]
		RangeCollection Search(string searchText, [Optional][TypeScriptType(ParamTypeStrings.SearchOptions)]SearchOptions searchOptions);

		/// <summary>
		/// Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.
		/// </summary>
		/// <param name="rangeLocation">Optional. The range location can be 'Whole', 'Start', 'End', 'After' or 'Content'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_GetRange)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		Range GetRange([Optional] RangeLocation rangeLocation);

		/// <summary>
		/// Splits the paragraph into child ranges by using delimiters.
		/// </summary>
		/// <param name="delimiters">Required. The delimiters as an array of strings.</param>
		/// <param name="trimDelimiters">Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.</param>
		/// <param name="trimSpacing">Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_Split)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		RangeCollection Split(string[] delimiters, [Optional] bool trimDelimiters, [Optional] bool trimSpacing);

		/// <summary>
		/// Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.
		/// </summary>
		/// <param name="rowCount">Required. The number of rows in the table.</param>
		/// <param name="columnCount">Required. The number of columns in the table.</param>
		/// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
		/// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_InsertTable)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.3)]
		Table InsertTable(int rowCount, int columnCount, InsertLocation insertLocation, [Optional] string[][] values);

		/// <summary>
		/// Gets the text ranges in the paragraph by using punctuation marks and/or other ending marks.
		/// </summary>
		/// <param name="endingMarks">Required. The punctuation marks and/or other ending marks as an array of strings.</param>
		/// <param name="trimSpacing">Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_GetTextRanges)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		RangeCollection GetTextRanges(string[] endingMarks, [Optional] bool trimSpacing);

		/// <summary>
		/// Starts a new list with this paragraph. Fails if the paragraph is already a list item.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_StartNewList)]
		[ApiSet(Version = 1.3)]
		List StartNewList();

		/// <summary>
		/// Lets the paragraph join an existing list at the specified level. Fails if the paragraph cannot join the list or if the paragraph is already a list item.
		/// </summary>
		/// <param name="listId">Required. The ID of an existing list.</param>
		/// <param name="level">Required. The level in the list.</param>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_AttachToList)]
		List AttachToList(int listId, int level);

		/// <summary>
		/// Moves this paragraph out of its list, if the paragraph is a list item.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_DetachFromList)]
		[ApiSet(Version = 1.3)]
		void DetachFromList();

		/// <summary>
		/// Gets the next paragraph.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_GetNext)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		Paragraph GetNext();

		/// <summary>
		/// Gets the previous paragraph.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Paragraph_GetPrevious)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		Paragraph GetPrevious();
	}

	/// <summary>
	/// Contains a collection of [paragraph](paragraph.md) objects.
	/// </summary>
	[ClientCallableType(HiddenIndexerMethod = true)]
	[ClientCallableComType(Name = "IParagraphCollection", InterfaceId = "20A0D5B4-FB08-460C-96AA-6AF4AFE8F1BA",
		CoClassName = "ParagraphCollection", SupportEnumeration = true, SupportIEnumVARIANT = true)]
	[ApiSet(Version = 1.1)]
	public interface ParagraphCollection : IEnumerable<Paragraph>
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// Gets a paragraph object by its index in the collection.
		/// </summary>
		/// <param name="index">A number that identifies the index location of a paragraph object.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.ParagraphCollection_Indexer)]
		[ApiSet(Version = 1.1)]
		Paragraph this[[TypeScriptType("number")]object index] { get; }

		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ParagraphCollection_ReferenceId)]
		string _ReferenceId { get; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.ParagraphCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.ParagraphCollection_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		/// <summary>
		/// Gets the first paragraph in this collection.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ParagraphCollection_GetFirst)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		Paragraph GetFirst();

		/// <summary>
		/// Gets the last paragraph in this collection.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.ParagraphCollection_GetLast)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		Paragraph GetLast();
	}

	/// <summary>
	/// Represents a contiguous area in a document.
	/// </summary>
	[ClientCallableType(ExposeIsNullProperty = true)]
	[ClientCallableComType(Name = "IRange", InterfaceId = "4869FD78-C024-4E18-9405-FFA60E9BFD50", CoClassName = "Range")]
	[ApiSet(Version = 1.1)]
	public interface Range
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// ID
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Id)]
		int _Id { get; }

		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_ReferenceId)]
		string _ReferenceId { get; }

		/// <summary>
		/// Gets the text format of the range. Use this to get and set font name, size, color, and other properties. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Font)]
		[JsonStringify()]
		[ApiSet(Version = 1.1)]
		Font Font { get; }

		/// <summary>
		/// Gets or sets the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Style)]
		[ApiSet(Version = 1.1)]
		string Style { get; set; }

		/// <summary>
		/// Gets or sets the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_StyleBuiltIn)]
		[ApiSet(Version = 1.3)]
		Style StyleBuiltIn { get; set; }

		/// <summary>
		/// Gets the collection of paragraph objects in the range. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Paragraphs)]
		[ApiSet(Version = 1.1)]
		ParagraphCollection Paragraphs { get; }

		/// <summary>
		/// Gets the collection of content control objects in the range. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_ContentControls)]
		[ApiSet(Version = 1.1)]
		ContentControlCollection ContentControls { get; }

		/// <summary>
		/// Gets the content control that contains the range. Returns a null object if there isn't a parent content control. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_ParentContentControl)]
		[ApiSet(Version = 1.1)]
		ContentControl ParentContentControl { get; }

		/// <summary>
		/// Gets the collection of inline picture objects in the range. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_InlinePictures)]
		[ApiSet(Version = 1.2)]
		InlinePictureCollection InlinePictures { get; }

		/// <summary>
		/// Gets the text of the range. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Text)]
		[ApiSet(Version = 1.1)]
		string Text { get; }

		/// <summary>
		/// Checks whether the range length is zero. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_IsEmpty)]
		[ApiSet(Version = 1.3)]
		bool IsEmpty { get; }

		/// <summary>
		/// Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a newline character ('\n') to separate the address part from the optional location part.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Hyperlink)]
		[ApiSet(Version = 1.3)]
		string Hyperlink { get; set; }

		/// <summary>
		/// Gets the parent body of the range. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_ParentBody)]
		[ApiSet(Version = 1.3)]
		Body ParentBody { get; }

		/// <summary>
		/// Gets the collection of list objects in the range. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Lists)]
		[ApiSet(Version = 1.3)]
		ListCollection Lists { get; }

		/// <summary>
		/// Gets the collection of table objects in the range. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Tables)]
		[ApiSet(Version = 1.3)]
		TableCollection Tables { get; }

		/// <summary>
		/// Gets the table cell that contains the range. Returns a null object if it is not contained in a table cell. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_ParentTableCell)]
		[ApiSet(Version = 1.3)]
		TableCell ParentTableCell { get; }

		/// <summary>
		/// Gets the table that contains the range. Returns null if it is not contained in a table. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_ParentTable)]
		[ApiSet(Version = 1.3)]
		Table ParentTable { get; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.Range_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.Range_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		/// <summary>
		/// Clears the contents of the range object. The user can perform the undo operation on the cleared content.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Clear)]
		[ApiSet(Version = 1.1)]
		void Clear();

		/// <summary>
		/// Deletes the range and its content from the document.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Delete)]
		[ClientCallableOperation(WacName = "DeleteRange")]
		[ApiSet(Version = 1.1)]
		void Delete();

		/// <summary>
		/// Wraps the range object with a rich text content control.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_InsertContentControl)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.1)]
		ContentControl InsertContentControl();

		/// <summary>
		/// Inserts text at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
		/// </summary>
		/// <param name="text">Required. Text to be inserted.</param>
		/// <param name="insertLocation">Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_InsertText)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.1)]
		Range InsertText(string text, InsertLocation insertLocation);

		/// <summary>
		/// Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
		/// </summary>
		/// <param name="html">Required. The HTML to be inserted.</param>
		/// <param name="insertLocation">Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_InsertHtml)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.1)]
		Range InsertHtml(string html, InsertLocation insertLocation);

		/// <summary>
		/// Inserts OOXML at the specified location.  The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
		/// </summary>
		/// <param name="ooxml">Required. The OOXML to be inserted.</param>
		/// <param name="insertLocation">Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_InsertOoxml)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.1)]
		Range InsertOoxml(string ooxml, InsertLocation insertLocation);

		/// <summary>
		/// Inserts a document at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
		/// </summary>
		/// <param name="base64File">Required. The base64 encoded content of a .docx file.</param>
		/// <param name="insertLocation">Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_InsertFileFromBase64)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.1)]
		Range InsertFileFromBase64(string base64File, InsertLocation insertLocation);

		/// <summary>
		/// Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.
		/// </summary>
		/// <param name="paragraphText">Required. The paragraph text to be inserted.</param>
		/// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_InsertParagraph)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.1)]
		Paragraph InsertParagraph(string paragraphText, InsertLocation insertLocation);

		/// <summary>
		/// Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.
		/// </summary>
		/// <param name="breakType">Required. The break type to add.</param>
		/// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_InsertBreak)]
		[ClientCallableOperation(WacAsync = true)]
		[ApiSet(Version = 1.1)]
		void InsertBreak(BreakType breakType, InsertLocation insertLocation);

		/// <summary>
		/// Inserts a picture at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
		/// </summary>
		/// <param name="base64EncodedImage">Required. The base64 encoded image to be inserted.</param>
		/// <param name="insertLocation">Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_InsertInlinePictureFromBase64)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.2)]
		InlinePicture InsertInlinePictureFromBase64(string base64EncodedImage, InsertLocation insertLocation);

		/// <summary>
		/// Selects and navigates the Word UI to the range.
		/// </summary>
		/// <param name="selectionMode">Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Select)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.1)]
		void Select([Optional] SelectionMode selectionMode);

		/// <summary>
		/// Gets the HTML representation of the range object.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_GetHtml)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.1)]
		string GetHtml();

		/// <summary>
		/// Gets the OOXML representation of the range object.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_GetOoxml)]
		[ClientCallableOperation(OperationType = OperationType.Read, WacAsync = true)]
		[ApiSet(Version = 1.1)]
		string GetOoxml();

		/// <summary>
		/// Performs a search with the specified searchOptions on the scope of the range object. The search results are a collection of range objects.
		/// </summary>
		/// <param name="searchText">Required. The search text.</param>
		/// <param name="searchOptions">Optional. Options for the search.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Search)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.1)]
		RangeCollection Search(string searchText, [Optional][TypeScriptType(ParamTypeStrings.SearchOptions)]SearchOptions searchOptions);

		/// <summary>
		/// Clones the range, or gets the starting or ending point of the range as a new range.
		/// </summary>
		/// <param name="rangeLocation">Optional. The range location can be 'Whole', 'Start', 'End', 'After' or 'Content'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_GetRange)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		Range GetRange([Optional] RangeLocation rangeLocation);

		/// <summary>
		/// Splits the range into child ranges by using delimiters.
		/// </summary>
		/// <param name="delimiters">Required. The delimiters as an array of strings.</param>
		/// <param name="multiParagraphs">Optional. Indicates whether a returned child range can cover multiple paragraphs. Default is false which indicates that the paragraph boundaries are also used as delimiters.</param>
		/// <param name="trimDelimiters">Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.</param>
		/// <param name="trimSpacing">Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_Split)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		RangeCollection Split(string[] delimiters, [Optional] bool multiParagraphs, [Optional] bool trimDelimiters, [Optional] bool trimSpacing);

		/// <summary>
		/// Compares this range's location with another range's location.
		/// </summary>
		/// <param name="range">Required. The range to compare with this range.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_CompareLocationWith)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		LocationRelation CompareLocationWith(Range range);

		/// <summary>
		/// Returns a new range that extends from this range in either direction to cover another range. This range is not changed.
		/// </summary>
		/// <param name="range">Required. Another range.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_ExpandTo)]
		[ApiSet(Version = 1.3)]
		Range ExpandTo(Range range);

		/// <summary>
		/// Returns a new range as the intersection of this range with another range. This range is not changed.
		/// </summary>
		/// <param name="range">Required. Another range.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_IntersectWith)]
		[ApiSet(Version = 1.3)]
		Range IntersectWith(Range range);

		/// <summary>
		/// Gets the next text range by using punctuation marks and/or other ending marks.
		/// </summary>
		/// <param name="endingMarks">Required. The punctuation marks and/or other ending marks as an array of strings.</param>
		/// <param name="trimSpacing">Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the returned range. Default is false which indicates that spacing characters at the start and end of the range are included.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_GetNextTextRange)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		Range GetNextTextRange(string[] endingMarks, [Optional] bool trimSpacing);

		/// <summary>
		/// Gets hyperlink child ranges within the range.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_GetHyperlinkRanges)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		RangeCollection GetHyperlinkRanges();

		/// <summary>
		/// Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.
		/// </summary>
		/// <param name="rowCount">Required. The number of rows in the table.</param>
		/// <param name="columnCount">Required. The number of columns in the table.</param>
		/// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
		/// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_InsertTable)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.3)]
		Table InsertTable(int rowCount, int columnCount, InsertLocation insertLocation, [Optional] string[][] values);

		/// <summary>
		/// Gets the text child ranges in the range by using punctuation marks and/or other ending marks.
		/// </summary>
		/// <param name="endingMarks">Required. The punctuation marks and/or other ending marks as an array of strings.</param>
		/// <param name="trimSpacing">Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_GetTextRanges)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		RangeCollection GetTextRanges(string[] endingMarks, [Optional] bool trimSpacing);

		/// <summary>
		/// Gets the names all bookmarks in or overlapping the range. A bookmark is hidden if its name starts with an underscore character.
		/// </summary>
		/// <param name="includeHidden">Optional. Indicates whether to include hidden bookmarks. Default is false which indicates that the hidden bookmarks are excluded.</param>
		/// <param name="includeAdjacent">Optional. Indicates whether to include bookmarks that are adjacent to the range. Default is false which indicates that the adjacent bookmarks are excluded.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_GetBookmarks)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.4)]
		string[] GetBookmarks([Optional] bool includeHidden, [Optional] bool includeAdjacent);

		/// <summary>
		/// Inserts a bookmark on the range. If a bookmark of the same name exists, it is replaced.
		/// </summary>
		/// <param name="name">Required. The bookmark name, which is case-insensitive. If the name starts with an underscore character, the bookmark is an hidden one.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Range_InsertBookmark)]
		[ApiSet(Version = 1.4)]
		void InsertBookmark(string name);
	}

	/// <summary>
	/// Contains a collection of [range](range.md) objects.
	/// </summary>
	[ClientCallableType(HiddenIndexerMethod = true)]
	[ClientCallableComType(Name = "IRangeCollection", InterfaceId = "38A0DB4C-D855-40AA-879E-2251A0F17B1F",
		CoClassName = "RangeCollection", SupportEnumeration = true, SupportIEnumVARIANT = true)]
	[ApiSet(Version = 1.3)]
	public interface RangeCollection : IEnumerable<Range>
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// Gets a range object by its index in the collection.
		/// </summary>
		/// <param name="index">A number that identifies the index location of a range object.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.RangeCollection_Indexer)]
		[ApiSet(Version = 1.3)]
		Range this[[TypeScriptType("number")]object index] { get; }

		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.RangeCollection_ReferenceId)]
		string _ReferenceId { get; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.RangeCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.RangeCollection_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		/// <summary>
		/// Gets the first range in this collection.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.RangeCollection_GetFirst)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		Range GetFirst();
	}

	/// <summary>
	/// Specifies the options to be included in a search operation.
	/// </summary>
	[ClientCallableComType(Name = "ISearchOptions", InterfaceId = "957E194F-14AF-462D-BD95-E29A55DE3B4A",
		CoClassName = "SearchOptions", CoClassId = "88B0A400-950B-42B1-B183-258F504BF71F")]
	[ApiSet(Version = 1.1)]
	public interface SearchOptions
	{
		//===============================================================================
		// Properties
		//
		// IMPORTANT: Whenever new property is added here, update
		//     "ParamTypeStrings.SearchOptions" string so that fancy style search
		//     options correctly appears on JavaScript IntelliSense tooltip.
		//===============================================================================
		/// <summary>
		/// Gets or sets a value that indicates whether to ignore all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.SearchOptions_IgnorePunct)]
		[ApiSet(Version = 1.1)]
		bool IgnorePunct { get; set; }

		/// <summary>
		/// Gets or sets a value that indicates whether to ignore all whitespace between words. Corresponds to the Ignore whitespace characters check box in the Find and Replace dialog box.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.SearchOptions_IgnoreSpace)]
		[ApiSet(Version = 1.1)]
		bool IgnoreSpace { get; set; }

		/// <summary>
		/// Gets or sets a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box (Edit menu).
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.SearchOptions_MatchCase)]
		[ApiSet(Version = 1.1)]
		bool MatchCase { get; set; }

		/// <summary>
		/// Gets or sets a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.SearchOptions_MatchPrefix)]
		[ApiSet(Version = 1.1)]
		bool MatchPrefix { get; set; }

		/// <summary>
		/// Gets or sets a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.SearchOptions_MatchSuffix)]
		[ApiSet(Version = 1.1)]
		bool MatchSuffix { get; set; }

		/// <summary>
		/// Gets or sets a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.SearchOptions_MatchWildcards)]
		[ApiSet(Version = 1.1)]
		bool MatchWildcards { get; set; }

		/// <summary>
		/// Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.SearchOptions_MatchWholeWord)]
		[ApiSet(Version = 1.1)]
		bool MatchWholeWord { get; set; }

		//===============================================================================
		// Methods
		//===============================================================================
	}

	/// <summary>
	/// Represents a section in a Word document.
	/// </summary>
	[ClientCallableType(ExposeIsNullProperty = true)]
	[ClientCallableComType(Name = "ISection", InterfaceId = "A9D1403C-908B-4178-AA45-6B406AEB8CBF", CoClassName = "Section")]
	[ApiSet(Version = 1.1)]
	public interface Section
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// ID
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Section_Id)]
		int _Id { get; }

		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Section_ReferenceId)]
		string _ReferenceId { get; }

		/// <summary>
		/// Gets the body object of the section. This does not include the header/footer and other section metadata. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Section_Body)]
		[JsonStringify()]
		[ApiSet(Version = 1.1)]
		Body Body { get; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.Section_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.Section_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		/// <summary>
		/// Gets one of the section's headers.
		/// </summary>
		/// <param name="type">Required. The type of header to return. This value can be: 'primary', 'firstPage' or 'evenPages'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Section_GetHeader)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.1)]
		Body GetHeader(HeaderFooterType type);

		/// <summary>
		/// Gets one of the section's footers.
		/// </summary>
		/// <param name="type">Required. The type of footer to return. This value can be: 'primary', 'firstPage' or 'evenPages'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Section_GetFooter)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.1)]
		Body GetFooter(HeaderFooterType type);

		/// <summary>
		/// Gets the next section.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Section_GetNext)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		Section GetNext();
	}

	/// <summary>
	/// Contains the collection of the document's [section](section.md) objects.
	/// </summary>
	[ClientCallableType(HiddenIndexerMethod = true)]
	[ClientCallableComType(Name = "ISectionCollection", InterfaceId = "23EF4423-1C46-48EE-951A-CEC1D21D3A5B",
		CoClassName = "SectionCollection", SupportEnumeration = true, SupportIEnumVARIANT = true)]
	[ApiSet(Version = 1.1)]
	public interface SectionCollection : IEnumerable<Section>
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// Gets a section object by its index in the collection.
		/// </summary>
		/// <param name="index">A number that identifies the index location of a section object.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.SectionCollection_Indexer)]
		[ApiSet(Version = 1.1)]
		Section this[[TypeScriptType("number")]object index] { get; }

		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.SectionCollection_ReferenceId)]
		string _ReferenceId { get; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.SectionCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.SectionCollection_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		/// <summary>
		/// Gets the first section in this collection.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.SectionCollection_GetFirst)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		Section GetFirst();
	}

	/// <summary>
	/// Represents a setting of the add-in.
	/// </summary>
	[ClientCallableType(ExposeIsNullProperty = true)]
	[ClientCallableComType(Name = "ISetting", InterfaceId = "03764D8A-7C3A-4A5F-82EC-029FBC6F57EE", CoClassName = "Setting")]
	[ApiSet(Version = 1.4)]
	public interface Setting
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Setting_ReferenceId)]
		string _ReferenceId { get; }

		/// <summary>
		/// Gets the key of the setting. Read only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Setting_Key)]
		[ApiSet(Version = 1.4)]
		string Key { get; }

		/// <summary>
		/// Gets or sets the value of the setting.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Setting_Value)]
		[ApiSet(Version = 1.4)]
		object Value { get; set; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.Setting_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.Setting_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		/// <summary>
		/// Deletes the setting.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Setting_Delete)]
		[ApiSet(Version = 1.4)]
		void Delete();
	}

	/// <summary>
	/// Contains the collection of [setting](setting.md) objects.
	/// </summary>
	[ClientCallableComType(Name = "ISettingCollection", InterfaceId = "02C3B0B1-2468-48C1-8C30-C1E38C6734F7",
		CoClassName = "SettingCollection", SupportEnumeration = true, SupportIEnumVARIANT = true)]
	[ApiSet(Version = 1.4)]
	public interface SettingCollection : IEnumerable<Setting>
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.SettingCollection_ReferenceId)]
		string _ReferenceId { get; }

		/// <summary>
		/// Gets a setting object by its key, which is case-sensitive.
		/// </summary>
		/// <param name="key">The key that identifies the setting object.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.SettingCollection_Indexer)]
		[ApiSet(Version = 1.4)]
		Setting this[string key] { get; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.SettingCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.SettingCollection_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		/// <summary>
		/// Creates or sets a setting.
		/// </summary>
		/// <param name="key">Required. The setting's key, which is case-sensitive.</param>
		/// <param name="value">Required. The setting's value.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.SettingCollection_Set)]
		[ApiSet(Version = 1.4)]
		Setting Set(string key, object value);

		/// <summary>
		/// Gets the count of settings.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.SettingCollection_GetCount)]
		[ApiSet(Version = 1.4)]
		int GetCount();

		/// <summary>
		/// Deletes all settings in this add-in.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.SettingCollection_DeleteAll)]
		[ApiSet(Version = 1.4)]
		void DeleteAll();
	}

	/// <summary>
	/// Represents a table in a Word document.
	/// </summary>
	[ClientCallableType(ExposeIsNullProperty = true)]
	[ClientCallableComType(Name = "ITable", InterfaceId = "55EB3BC4-ACEC-4FED-B8D3-10414E633465", CoClassName = "Table")]
	[ApiSet(Version = 1.3)]
	public interface Table
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// ID
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_Id)]
		int _Id { get; }

		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_ReferenceId)]
		string _ReferenceId { get; }

		/// <summary>
		/// Gets all of the table rows. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_Rows)]
		[ApiSet(Version = 1.3)]
		TableRowCollection Rows { get; }

		/// <summary>
		/// Indicates whether all of the table rows are uniform. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_IsUniform)]
		[ApiSet(Version = 1.3)]
		bool IsUniform { get; }

		/// <summary>
		/// Gets the child tables nested one level deeper. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_Tables)]
		[ApiSet(Version = 1.3)]
		TableCollection Tables { get; }

		/// <summary>
		/// Gets the nesting level of the table. Top-level tables have level 1. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_NestingLevel)]
		[ApiSet(Version = 1.3)]
		int NestingLevel { get; }

		/// <summary>
		/// Gets the table cell that contains this table. Returns a null object if it is not contained in a table cell. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_ParentTableCell)]
		[ApiSet(Version = 1.3)]
		TableCell ParentTableCell { get; }

		/// <summary>
		/// Gets the table that contains this table. Returns a null object if it is not contained in a table. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_ParentTable)]
		[ApiSet(Version = 1.3)]
		Table ParentTable { get; }

		/// <summary>
		/// Gets and sets the text values in the table, as a 2D Javascript array.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_Values)]
		[ApiSet(Version = 1.3)]
		string[][] Values { get; set; }

		/// <summary>
		/// Gets or sets the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_Style)]
		[ApiSet(Version = 1.3)]
		string Style { get; set; }

		/// <summary>
		/// Gets or sets the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_StyleBuiltIn)]
		[ApiSet(Version = 1.3)]
		Style StyleBuiltIn { get; set; }

		/// <summary>
		/// Gets the number of rows in the table. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_RowCount)]
		[ApiSet(Version = 1.3)]
		int RowCount { get; }

		/// <summary>
		/// Gets and sets the number of header rows.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_HeaderRowCount)]
		[ApiSet(Version = 1.3)]
		int HeaderRowCount { get; set; }

		/// <summary>
		/// Gets and sets whether the table has a total (last) row with a special style.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_StyleTotalRow)]
		[ApiSet(Version = 1.3)]
		bool StyleTotalRow { get; set; }

		/// <summary>
		/// Gets and sets whether the table has a first column with a special style.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_StyleFirstColumn)]
		[ApiSet(Version = 1.3)]
		bool StyleFirstColumn { get; set; }

		/// <summary>
		/// Gets and sets whether the table has a last column with a special style.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_StyleLastColumn)]
		[ApiSet(Version = 1.3)]
		bool StyleLastColumn { get; set; }

		/// <summary>
		/// Gets and sets whether the table has banded rows.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_StyleBandedRows)]
		[ApiSet(Version = 1.3)]
		bool StyleBandedRows { get; set; }

		/// <summary>
		/// Gets and sets whether the table has banded columns.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_StyleBandedColumns)]
		[ApiSet(Version = 1.3)]
		bool StyleBandedColumns { get; set; }

		/// <summary>
		/// Gets and sets the shading color.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_ShadingColor)]
		[ApiSet(Version = 1.3)]
		string ShadingColor { get; set; }

		/// <summary>
		/// Gets and sets the horizontal alignment of every cell in the table. The value can be 'left', 'centered', 'right', or 'justified'.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_HorizontalAlignment)]
		[ApiSet(Version = 1.3)]
		Alignment HorizontalAlignment { get; set; }

		/// <summary>
		/// Gets and sets the vertical alignment of every cell in the table. The value can be 'top', 'center' or 'bottom'.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_VerticalAlignment)]
		[ApiSet(Version = 1.3)]
		VerticalAlignment VerticalAlignment { get; set; }

		/// <summary>
		/// Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_Font)]
		[JsonStringify()]
		[ApiSet(Version = 1.3)]
		Font Font { get; }

		/// <summary>
		/// Gets the content control that contains the table. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_ParentContentControl)]
		[ApiSet(Version = 1.3)]
		ContentControl ParentContentControl { get; }

		/// <summary>
		/// Gets the height of the table in points. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_Height)]
		[ApiSet(Version = 1.3)]
		float Height { get; }

		/// <summary>
		/// Gets and sets the width of the table in points.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_Width)]
		[ApiSet(Version = 1.3)]
		float Width { get; set; }

		/// <summary>
		/// Gets the paragraph before the table. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_ParagraphBefore)]
		[ApiSet(Version = 1.3)]
		Paragraph ParagraphBefore { get; }

		/// <summary>
		/// Gets the paragraph after the table. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_ParagraphAfter)]
		[ApiSet(Version = 1.3)]
		Paragraph ParagraphAfter { get; }

		/// <summary>
		/// Gets the parent body of the table. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_ParentBody)]
		[ApiSet(Version = 1.3)]
		Body ParentBody { get; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.Table_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.Table_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		/// <summary>
		/// Adds rows to the start or end of the table, using the first or last existing row as a template. The string values, if specified, are set in the newly inserted rows.
		/// </summary>
		/// <param name="insertLocation">Required. It can be 'Start' or 'End'.</param>
		/// <param name="rowCount">Required. Number of rows to add.</param>
		/// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_AddRows)]
		[ApiSet(Version = 1.3)]
		TableRowCollection AddRows(InsertLocation insertLocation, int rowCount, [Optional] string[][] values);

		/// <summary>
		/// Adds columns to the start or end of the table, using the first or last existing column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.
		/// </summary>
		/// <param name="insertLocation">Required. It can be 'Start' or 'End', corresponding to the appropriate side of the table.</param>
		/// <param name="columnCount">Required. Number of columns to add.</param>
		/// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_AddColumns)]
		[ApiSet(Version = 1.3)]
		void AddColumns(InsertLocation insertLocation, int columnCount, [Optional] string[][] values);

		/// <summary>
		/// Gets the table cell at a specified row and column.
		/// </summary>
		/// <param name="rowIndex">Required. The index of the row.</param>
		/// <param name="cellIndex">Required. The index of the cell in the row.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_GetCell)]
		[ApiSet(Version = 1.3)]
		TableCell GetCell(int rowIndex, int cellIndex);

		/// <summary>
		/// Merges the cells bounded inclusively by a first and last cell.
		/// </summary>
		/// <param name="topRow">Required. The row of the first cell</param>
		/// <param name="firstCell">Required. The index of the first cell in its row</param>
		/// <param name="bottomRow">Required. The row of the last cell</param>
		/// <param name="lastCell">Required. The index of the last cell in its row</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_MergeCells)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.4)]
		TableCell MergeCells(int topRow, int firstCell, int bottomRow, int lastCell);

		/// <summary>
		/// Deletes the entire table.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_Delete)]
		[ApiSet(Version = 1.3)]
		void Delete();

		/// <summary>
		/// Clears the contents of the table.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_Clear)]
		[ApiSet(Version = 1.3)]
		void Clear();

		/// <summary>
		/// Deletes specific rows.
		/// </summary>
		/// <param name="rowIndex">Required. The first row to delete.</param>
		/// <param name="rowCount">Optional. The number of rows to delete. Default 1.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_DeleteRows)]
		[ApiSet(Version = 1.3)]
		void DeleteRows(int rowIndex, [Optional] int? rowCount);

		/// <summary>
		/// Deletes specific columns. This is applicable to uniform tables.
		/// </summary>
		/// <param name="columnIndex">Required. The first column to delete.</param>
		/// <param name="columnCount">Optional. The number of columns to delete. Default 1.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_DeleteColumns)]
		[ApiSet(Version = 1.3)]
		void DeleteColumns(int columnIndex, [Optional] int? columnCount);

		/// <summary>
		/// Autofits the table columns to the width of their contents.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_AutoFitContents)]
		[ApiSet(Version = 1.3)]
		void AutoFitContents();

		/// <summary>
		/// Autofits the table columns to the width of the window.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_AutoFitWindow)]
		[ApiSet(Version = 1.3)]
		void AutoFitWindow();

		/// <summary>
		/// Distributes the row heights evenly.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_DistributeRows)]
		[ApiSet(Version = 1.3)]
		void DistributeRows();

		/// <summary>
		/// Distributes the column widths evenly.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_DistributeColumns)]
		[ApiSet(Version = 1.3)]
		void DistributeColumns();

		/// <summary>
		/// Gets the border style for the specified border.
		/// </summary>
		/// <param name="borderLocation">Required. The border location.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_GetBorder)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		TableBorder GetBorder(BorderLocation borderLocation);

		/// <summary>
		/// Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.
		/// </summary>
		/// <param name="selectionMode">Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_Select)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		void Select([Optional] SelectionMode selectionMode);

		/// <summary>
		/// Performs a search with the specified searchOptions on the scope of the table object. The search results are a collection of range objects.
		/// </summary>
		/// <param name="searchText">Required. The search text.</param>
		/// <param name="searchOptions">Optional. Options for the search.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_Search)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.3)]
		RangeCollection Search(string searchText, [Optional][TypeScriptType(ParamTypeStrings.SearchOptions)]SearchOptions searchOptions);

		/// <summary>
		/// Gets the range that contains this table, or the range at the start or end of the table. 
		/// </summary>
		/// <param name="rangeLocation">Optional. The range location can be 'Whole', 'Start', 'End' or 'After'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_GetRange)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		Range GetRange([Optional] RangeLocation rangeLocation);

		/// <summary>
		/// Inserts a content control on the table.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_InsertContentControl)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.3)]
		ContentControl InsertContentControl();

		/// <summary>
		/// Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.
		/// </summary>
		/// <param name="rowCount">Required. The number of rows in the table.</param>
		/// <param name="columnCount">Required. The number of columns in the table.</param>
		/// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
		/// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_InsertTable)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.3)]
		Table InsertTable(int rowCount, int columnCount, InsertLocation insertLocation, [Optional] string[][] values);

		/// <summary>
		/// Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.
		/// </summary>
		/// <param name="paragraphText">Required. The paragraph text to be inserted.</param>
		/// <param name="insertLocation">Required. The value can be 'Before' or 'After'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_InsertParagraph)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		[ApiSet(Version = 1.3)]
		Paragraph InsertParagraph(string paragraphText, InsertLocation insertLocation);

		/// <summary>
		/// Gets cell padding in points. 
		/// </summary>
		/// <param name="cellPaddingLocation">Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_GetCellPadding)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		float? GetCellPadding(CellPaddingLocation cellPaddingLocation);

		/// <summary>
		/// Sets cell padding in points.
		/// </summary>
		/// <param name="cellPaddingLocation">Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_SetCellPadding)]
		[ApiSet(Version = 1.3)]
		void SetCellPadding(CellPaddingLocation cellPaddingLocation, float cellPadding);

		/// <summary>
		/// Gets the next table.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_GetNext)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		Table GetNext();
	}

	/// <summary>
	/// Contains the collection of the document's Table objects.
	/// </summary>
	[ClientCallableType(HiddenIndexerMethod = true)]
	[ClientCallableComType(Name = "ITableCollection", InterfaceId = "793CA292-E6DC-48F7-B0E6-16311394A5B2",
		CoClassName = "TableCollection", SupportEnumeration = true, SupportIEnumVARIANT = true)]
	[ApiSet(Version = 1.3)]
	public interface TableCollection : IEnumerable<Table>
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// Gets a table object by its index in the collection.
		/// </summary>
		/// <param name="index">A number that identifies the index location of a table object.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCollection_Indexer)]
		[ApiSet(Version = 1.3)]
		Table this[[TypeScriptType("number")]object index] { get; }

		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCollection_ReferenceId)]
		string _ReferenceId { get; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.TableCollection_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		[ClientCallableComMember(DispatchId = DispatchIds.SectionCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Gets the first table in this collection.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCollection_GetFirst)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		Table GetFirst();
	}

	/// <summary>
	/// Represents a row in a Word document.
	/// </summary>
	[ClientCallableType(ExposeIsNullProperty = true)]
	[ClientCallableComType(Name = "ITableRow", InterfaceId = "7A28087B-2D9C-4DB1-95AF-B1CE69D1D4DB", CoClassName = "TableRow")]
	[ApiSet(Version = 1.3)]
	public interface TableRow
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// ID
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_Id)]
		int _Id { get; }

		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_ReferenceId)]
		string _ReferenceId { get; }

		/// <summary>
		/// Gets cells. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_Cells)]
		[ApiSet(Version = 1.3)]
		TableCellCollection Cells { get; }

		/// <summary>
		/// Gets the number of cells in the row. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_CellCount)]
		[ApiSet(Version = 1.3)]
		int CellCount { get; }

		/// <summary>
		/// Gets parent table. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_ParentTable)]
		[ApiSet(Version = 1.3)]
		Table ParentTable { get; }

		/// <summary>
		/// Gets the index of the row in its parent table. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_RowIndex)]
		[ApiSet(Version = 1.3)]
		int RowIndex { get; }

		/// <summary>
		/// Gets and sets the text values in the row, as a 1D Javascript array.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_Values)]
		[ApiSet(Version = 1.3)]
		string[] Values { get; set; }

		/// <summary>
		/// Gets and sets the shading color.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_ShadingColor)]
		[ApiSet(Version = 1.3)]
		string ShadingColor { get; set; }

		/// <summary>
		/// Gets and sets the horizontal alignment of every cell in the row. The value can be 'left', 'centered', 'right', or 'justified'.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_HorizontalAlignment)]
		[ApiSet(Version = 1.3)]
		Alignment HorizontalAlignment { get; set; }

		/// <summary>
		/// Gets and sets the vertical alignment of the cells in the row. The value can be 'top', 'center' or 'bottom'.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_VerticalAlignment)]
		[ApiSet(Version = 1.3)]
		VerticalAlignment VerticalAlignment { get; set; }

		/// <summary>
		/// Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_Font)]
		[JsonStringify()]
		[ApiSet(Version = 1.3)]
		Font Font { get; }

		/// <summary>
		/// Checks whether the row is a header row. Read-only. To set the number of header rows, use HeaderRowCount on the Table object.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_IsHeader)]
		[ApiSet(Version = 1.3)]
		bool IsHeader { get; }

		/// <summary>
		/// Gets and sets the preferred height of the row in points.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_PreferredHeight)]
		[ApiSet(Version = 1.3)]
		float PreferredHeight { get; set; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		/// <summary>
		/// Inserts rows using this row as a template. If values are specified, inserts the values into the new rows.
		/// </summary>
		/// <param name="insertLocation">Required. Where the new rows should be inserted, relative to the current row. It can be 'Before' or 'After'.</param>
		/// <param name="rowCount">Required. Number of rows to add</param>
		/// <param name="values">Optional. Strings to insert in the new rows, specified as a 2D array. The number of cells in each row must not exceed the number of cells in the existing row.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_InsertRows)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		TableRowCollection InsertRows(InsertLocation insertLocation, int rowCount, [Optional] string[][] values);

		/// <summary>
		/// Merges the row into one cell.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_Merge)]
		[ApiSet(Version = 1.4)]
		TableCell Merge();

		/// <summary>
		/// Deletes the entire row.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_Delete)]
		[ApiSet(Version = 1.3)]
		void Delete();

		/// <summary>
		/// Clears the contents of the row.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_Clear)]
		[ApiSet(Version = 1.3)]
		void Clear();

		/// <summary>
		/// Gets the border style of the cells in the row.
		/// </summary>
		/// <param name="borderLocation">Required. The border location.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_GetBorder)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		TableBorder GetBorder(BorderLocation borderLocation);

		/// <summary>
		/// Selects the row and navigates the Word UI to it.
		/// </summary>
		/// <param name="selectionMode">Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_Select)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		void Select([Optional] SelectionMode selectionMode);

		/// <summary>
		/// Performs a search with the specified searchOptions on the scope of the row. The search results are a collection of range objects.
		/// </summary>
		/// <param name="searchText">Required. The search text.</param>
		/// <param name="searchOptions">Optional. Options for the search.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_Search)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true, WacAsync = true)]
		[ApiSet(Version = 1.3)]
		RangeCollection Search(string searchText, [Optional][TypeScriptType(ParamTypeStrings.SearchOptions)]SearchOptions searchOptions);

		/// <summary>
		/// Gets cell padding in points. 
		/// </summary>
		/// <param name="cellPaddingLocation">Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_GetCellPadding)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		float? GetCellPadding(CellPaddingLocation cellPaddingLocation);

		/// <summary>
		/// Sets cell padding in points.
		/// </summary>
		/// <param name="cellPaddingLocation">Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_SetCellPadding)]
		[ApiSet(Version = 1.3)]
		void SetCellPadding(CellPaddingLocation cellPaddingLocation, float cellPadding);

		/// <summary>
		/// Gets the next row.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_GetNext)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		TableRow GetNext();
	}

	/// <summary>
	/// Contains the collection of the document's TableRow objects.
	/// </summary>
	[ClientCallableType(HiddenIndexerMethod = true)]
	[ClientCallableComType(Name = "ITableRowCollection", InterfaceId = "F752737F-A469-460B-A2C1-08B958EFB48F",
		CoClassName = "TableRowCollection", SupportEnumeration = true, SupportIEnumVARIANT = true)]
	[ApiSet(Version = 1.3)]
	public interface TableRowCollection : IEnumerable<TableRow>
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// Gets a table row object by its index in the collection.
		/// </summary>
		/// <param name="index">A number that identifies the index location of a table row object.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRowCollection_Indexer)]
		[ApiSet(Version = 1.3)]
		TableRow this[[TypeScriptType("number")]object index] { get; }

		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRowCollection_ReferenceId)]
		string _ReferenceId { get; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.TableRowCollection_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		[ClientCallableComMember(DispatchId = DispatchIds.TableRowCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Gets the first row in this collection.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRowCollection_GetFirst)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		TableRow GetFirst();
	}

	/// <summary>
	/// Represents a table cell in a Word document.
	/// </summary>
	[ClientCallableType(ExposeIsNullProperty = true)]
	[ClientCallableComType(Name = "ITableCell", InterfaceId = "276A08F8-2AFC-42FD-A336-8890D712E195", CoClassName = "TableCell")]
	[ApiSet(Version = 1.3)]
	public interface TableCell
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// ID
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCell_Id)]
		int _Id { get; }

		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCell_ReferenceId)]
		string _ReferenceId { get; }

		/// <summary>
		/// Gets the parent table of the cell. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCell_ParentTable)]
		[ApiSet(Version = 1.3)]
		Table ParentTable { get; }

		/// <summary>
		/// Gets the parent row of the cell. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCell_ParentRow)]
		[ApiSet(Version = 1.3)]
		TableRow ParentRow { get; }

		/// <summary>
		/// Gets the index of the cell's row in the table. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCell_RowIndex)]
		[ApiSet(Version = 1.3)]
		int RowIndex { get; }

		/// <summary>
		/// Gets the index of the cell in its row. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCell_CellIndex)]
		[ApiSet(Version = 1.3)]
		int CellIndex { get; }

		/// <summary>
		/// Gets and sets the text of the cell.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCell_Value)]
		[ApiSet(Version = 1.3)]
		string Value { get; set; }

		/// <summary>
		/// Gets the body object of the cell. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCell_Body)]
		[JsonStringify()]
		[ApiSet(Version = 1.3)]
		Body Body { get; }

		/// <summary>
		/// Gets or sets the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCell_ShadingColor)]
		[ApiSet(Version = 1.3)]
		string ShadingColor { get; set; }

		/// <summary>
		/// Gets and sets the horizontal alignment of the cell. The value can be 'left', 'centered', 'right', or 'justified'.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableRow_HorizontalAlignment)]
		[ApiSet(Version = 1.3)]
		Alignment HorizontalAlignment { get; set; }

		/// <summary>
		/// Gets and sets the vertical alignment of the cell. The value can be 'top', 'center' or 'bottom'.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCell_VerticalAlignment)]
		[ApiSet(Version = 1.3)]
		VerticalAlignment VerticalAlignment { get; set; }

		/// <summary>
		/// Gets and sets the width of the cell's column in points. This is applicable to uniform tables.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCell_ColumnWidth)]
		[ApiSet(Version = 1.3)]
		float? ColumnWidth { get; set; }

		/// <summary>
		/// Gets the width of the cell in points. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCell_Width)]
		[ApiSet(Version = 1.3)]
		float Width { get; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.TableCell_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.TableCell_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		/// <summary>
		/// Inserts rows above or below the cell, using the cell's row as a template. The string values, if specified, are set in the newly inserted rows.
		/// </summary>
		/// <param name="insertLocation">Required. It can be 'Before' or 'After'.</param>
		/// <param name="rowCount">Required. Number of rows to add.</param>
		/// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCell_InsertRows)]
		[ApiSet(Version = 1.3)]
		TableRowCollection InsertRows(InsertLocation insertLocation, int rowCount, [Optional] string[][] values);

		/// <summary>
		/// Adds columns to the left or right of the cell, using the cell's column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.
		/// </summary>
		/// <param name="insertLocation">Required. It can be 'Before' or 'After'.</param>
		/// <param name="columnCount">Required. Number of columns to add</param>
		/// <param name="values">Optional 2D array. Cells are filled if the corresponding strings are specified in the array.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCell_InsertColumns)]
		[ApiSet(Version = 1.3)]
		void InsertColumns(InsertLocation insertLocation, int columnCount, [Optional] string[][] values);

		/// <summary>
		/// Adds columns to the left or right of the cell, using the existing column as a template. The string values, if specified, are set in the newly inserted rows.
		/// </summary>
		/// <param name="rowCount">Required. The number of rows to split into. Must be a divisor of the number of underlying rows.</param>
		/// <param name="columnCount">Required. The number of columns to split into.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCell_Split)]
		[ApiSet(Version = 1.4)]
		void Split(int rowCount, int columnCount);

		/// <summary>
		/// Deletes the row containing this cell.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCell_DeleteRow)]
		[ApiSet(Version = 1.3)]
		void DeleteRow();

		/// <summary>
		/// Deletes the column containing this cell. This is applicable to uniform tables.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCell_DeleteColumn)]
		[ApiSet(Version = 1.3)]
		void DeleteColumn();

		/// <summary>
		/// Gets the border style for the specified border.
		/// </summary>
		/// <param name="borderLocation">Required. The border location.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCell_GetBorder)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		TableBorder GetBorder(BorderLocation borderLocation);

		/// <summary>
		/// Gets cell padding in points. 
		/// </summary>
		/// <param name="cellPaddingLocation">Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_GetCellPadding)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		float? GetCellPadding(CellPaddingLocation cellPaddingLocation);

		/// <summary>
		/// Sets cell padding in points.
		/// </summary>
		/// <param name="cellPaddingLocation">Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.Table_SetCellPadding)]
		[ApiSet(Version = 1.3)]
		void SetCellPadding(CellPaddingLocation cellPaddingLocation, float cellPadding);

		/// <summary>
		/// Gets the next cell.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCell_GetNext)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		TableCell GetNext();
	}

	/// <summary>
	/// Contains the collection of the document's TableCell objects.
	/// </summary>
	[ClientCallableType(HiddenIndexerMethod = true)]
	[ClientCallableComType(Name = "ITableCellCollection", InterfaceId = "67DFFDFA-3A5C-431B-8EA7-9600E82453E3",
		CoClassName = "TableCellCollection", SupportEnumeration = true, SupportIEnumVARIANT = true)]
	[ApiSet(Version = 1.3)]
	public interface TableCellCollection : IEnumerable<TableCell>
	{
		//===============================================================================
		// Properties
		//===============================================================================
		/// <summary>
		/// Gets a table cell object by its index in the collection.
		/// </summary>
		/// <param name="index">A number that identifies the index location of a table cell object.</param>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCellCollection_Indexer)]
		[ApiSet(Version = 1.3)]
		TableCell this[[TypeScriptType("number")]object index] { get; }

		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCellCollection_ReferenceId)]
		string _ReferenceId { get; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.TableCellCollection_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();

		[ClientCallableComMember(DispatchId = DispatchIds.TableCellCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Gets the first table cell in this collection.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableCellCollection_GetFirst)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.3)]
		TableCell GetFirst();
	}

	/// <summary>
	/// Specifies the border style
	/// </summary>
	[ClientCallableComType(Name = "ITableBorder", InterfaceId = "3E45A91B-F887-4FEB-8FEF-21E1E40F6FB6", CoClassName = "TableBorder")]
	[ApiSet(Version = 1.3)]
	public interface TableBorder
	{
		/// <summary>
		/// ReferenceId
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableBorder_ReferenceId)]
		string _ReferenceId { get; }

		/// <summary>
		/// Gets or sets the table border color, as a hex value or name.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableBorder_Color)]
		[ApiSet(Version = 1.3)]
		string Color { get; set; }

		/// <summary>
		/// Gets or sets the type of the table border.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableBorder_Type)]
		[ApiSet(Version = 1.3)]
		BorderType Type { get; set; }

		/// <summary>
		/// Gets or sets the width, in points, of the table border. Not applicable to table border types that have fixed widths.
		/// </summary>
		[ClientCallableComMember(DispatchId = DispatchIds.TableBorder_Width)]
		[ApiSet(Version = 1.3)]
		float Width { get; set; }

		//===============================================================================
		// Methods
		//===============================================================================
		[ClientCallableComMember(DispatchId = DispatchIds.TableBorder_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		[ClientCallableComMember(DispatchId = DispatchIds.TableBorder_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();
	}

	/// <summary>
	/// Specifies supported content control types and subtypes.
	/// </summary>
	public enum ContentControlType
	{
		Unknown = 0x0,
		// the following are subtypes only
		RichTextInline = 0x1,
		RichTextParagraphs = 0x2,
		RichTextTableCell = 0x4,   // contains whole cell
		RichTextTableRow = 0x8,    // contains whole row
		RichTextTable = 0x10,      // contains whole table
		PlainTextInline = 0x20,
		PlainTextParagraph = 0x40,
		// the following are both types and subtypes
		Picture = 0x80,
		BuildingBlockGallery = 0x100,
		CheckBox = 0x200,
		ComboBox = 0x400,
		DropDownList = 0x800,
		DatePicker = 0x1000,
		RepeatingSection = 0x2000,
		// the following are types only
		RichText = 0x1F,
		PlainText = 0x60
	}

	/// <summary>
	/// ContentControl appearance
	/// </summary>
	/// <remarks>
	/// Either bounding box, or tags, or hidden
	/// </remarks>
	public enum ContentControlAppearance
	{
		BoundingBox = 0,
		Tags = 1,
		Hidden = 2
	}

	/// <summary>
	/// Underline types
	/// </summary>
	/// <remarks>
	/// The supported styles for underline format. Must have same value as in word\inc\props.h
	/// </remarks>
	public enum UnderlineType
	{
		Mixed = -1,
		None = 0,

		[Obsolete("Hidden is no longer supported.")]
		Hidden = 5,
		[Obsolete("DotLine is no longer supported.")]
		DotLine = 8,

		Single = 0x1,               // wdUnderlineSingle = kulSingle
		Word = 0x2,                 // wdUnderlineWords = kulWord
		Double = 0x3,               // wdUnderlineDouble = kulDouble
		Thick = 0x6,                // wdUnderlineThick = kulThick
		Dotted = 0x4,               // wdUnderlineDotted = kulDotted
		DottedHeavy = 0x14,         // wdUnderlineDottedHeavy = kulDottedHeavy
		DashLine = 0x7,             // wdUnderlineDash = kulDashLine
		DashLineHeavy = 0x17,       // wdUnderlineDashHeavy = kulDashedHeavy
		DashLineLong = 0x27,        // wdUnderlineDashLong = kulDashLong
		DashLineLongHeavy = 0x37,   // wdUnderlineDashLongHeavy = kulDashLongHeavy
		DotDashLine = 0x9,          // wdUnderlineDotDash = kulDotDashLine
		DotDashLineHeavy = 0x19,    // wdUnderlineDotDashHeavy = kulDotDashHeavy
		TwoDotDashLine = 0xa,       // wdUnderlineDotDotDash = kul2DotDashLine
		TwoDotDashLineHeavy = 0x1a, // wdUnderlineDotDotDashHeavy = kul2DotDashHeavy
		Wave = 0xb,                 // wdUnderlineWavy = kulWave
		WaveHeavy = 0x1b,           // wdUnderlineWavyHeavy = kulWaveHeavy
		WaveDouble = 0x2b           // wdUnderlineWavyDouble = kulWaveDouble
	}

	/// <summary>
	/// Page break, line break, and four section breaks
	/// </summary>
	public enum BreakType
	{
		/// <summary>
		/// Page break.
		/// </summary>
		Page = 0,              // ipbPage

		[Obsolete("Use sectionNext instead.")]
		Next = 2,              // ipbNext

		/// <summary>
		/// Section break, with the new section starting on the next page.
		/// </summary>
		SectionNext = 1002,    // will be converted to ipbNext

		/// <summary>
		/// Section break, with the new section starting on the same page.
		/// </summary>
		SectionContinuous = 3, // ipbCont

		/// <summary>
		/// Section break, with the new section starting on the next even-numbered page.
		/// </summary>
		SectionEven = 4,       // ipbEven

		/// <summary>
		/// Section break, with the new section starting on the next odd-numbered page.
		/// </summary>
		SectionOdd = 5,        // ipbOdd

		/// <summary>
		/// Line break.
		/// </summary>
		Line = 6               // ipbLine
	}

	/// <summary>
	/// The insertion location types
	/// </summary>
	/// <remarks>
	/// For an API call
	///     obj.insertSomething(newStuff, location);
	/// If the location is Before or After, 'newStuff' will be outside of the modified 'obj'.
	/// If the location is Start or End, 'newStuff' will be included as part of the modified 'obj'.
	/// </remarks>
	public enum InsertLocation
	{
		Before = 0,
		After = 1,
		Start = 2,
		End = 3,
		Replace = 4
	}

	public enum Alignment
	{
		Mixed = -2,        // for table, row and cell
		Unknown = -1,      // for a paragraph
		Left = 0,          // jcLeft
		Centered = 1,      // jcCenter
		Right = 2,         // jcRight
		Justified = 3      // jcBoth
	}

	public enum HeaderFooterType
	{
		Primary = 0,
		FirstPage = 1,
		EvenPages = 2
	}

	public enum BodyType
	{
		Unknown = -1,
		MainDoc = 0,
		Section = 1,
		Header = 2,
		Footer = 3,
		TableCell = 4
		// Add more type here when supported
	}

	public enum SelectionMode
	{
		Select = 0,
		Start = 1,
		End = 2
	}

	public enum ImageFormat
	{
		Unsupported = -2,
		Undefined = -1,
		Bmp = 0,
		Jpeg = 1,
		Gif = 2,
		Tiff = 3,
		Png = 4,
		Icon = 5,
		Exif = 6,
		Wmf = 7,
		Emf = 8,
		Pict = 9,
		Pdf = 10,
		Svg = 11
	}

	public enum RangeLocation
	{
		Whole = 0,     // The object's whole range. If the object is a paragraph content control or table content control, the EOP or Table characters after the content control are also included.
		Start = 1,     // The starting point of the object. For content control, it is the point after the opening tag.
		End = 2,       // The ending point of the object. For paragraph, it is the point before the EOP. For content control, it is the point before the closing tag.
		Before = 3,    // For content control only. It is the point before the opening tag.
		After = 4,     // The point after the object. If the object is a paragraph content control or table content control, it is the point after the EOP or Table characters.
		Content = 5    // The range between 'Start' and 'End'.
	}

	public enum LocationRelation
	{
		Unrelated = 0,       // Indicates that this instance and the range are in different sub-documents.
		Equal = 1,           // Indicates that this instance and the range represent the same range.
		ContainsStart = 2,   // Indicates that this instance contains the range and that it shares the same start character. The range does not share the same end character as this instance.
		ContainsEnd = 3,     // Indicates that this instance contains the range and that it shares the same end character. The range does not share the same start character as this instance.
		Contains = 4,        // Indicates that this instance contains the range, with the exception of the start and end character of this instance.
		InsideStart = 5,     // Indicates that this instance is inside the range and that it shares the same start character. The range does not share the same end character as this instance.
		InsideEnd = 6,       // Indicates that this instance is inside the range and that it shares the same end character. The range does not share the same start character as this instance.
		Inside = 7,          // Indicates that this instance is inside the range. The range does not share the same start and end characters as this instance.
		AdjacentBefore = 8,  // Indicates that this instance occurs before, and is adjacent to, the range.
		OverlapsBefore = 9,  // Indicates that this instance starts before the range and overlaps the range’s first character.
		Before = 10,         // Indicates that this instance occurs before the range.
		AdjacentAfter = 11,  // Indicates that this instance occurs after, and is adjacent to, the range.
		OverlapsAfter = 12,  // Indicates that this instance starts inside the range and overlaps the range’s last character.
		After = 13           // Indicates that this instance occurs after the range.
	}

	public enum BorderLocation
	{
		Top = 0,
		Left = 1,
		Bottom = 2,
		Right = 3,
		InsideHorizontal = 4,
		InsideVertical = 5,
		Inside = 6,
		Outside = 7,
		All = 8
	}

	public enum CellPaddingLocation
	{
		Top = 0,      // ibrcTop
		Left = 1,     // ibrcLeft
		Bottom = 2,   // ibrcBottom
		Right = 3     // ibrcRight
	}

	// If you add a new enum, be sure it matches the one in table.h, and
	// add an assertion in TableBase.cpp
	public enum BorderType
	{
		Mixed = -1,
		None = 0,
		Single = 1,
		Double = 3,
		Dotted = 6,
		Dashed = 7,
		DotDashed = 8,
		Dot2Dashed = 9,
		Triple = 10,
		ThinThickSmall = 11,
		ThickThinSmall = 12,
		ThinThickThinSmall = 13,
		ThinThickMed = 14,
		ThickThinMed = 15,
		ThinThickThinMed = 16,
		ThinThickLarge = 17,
		ThickThinLarge = 18,
		ThinThickThinLarge = 19,
		Wave = 20,
		DoubleWave = 21,
		DashedSmall = 22,
		DashDotStroked = 23,
		ThreeDEmboss = 24,
		ThreeDEngrave = 25
	}

	// If you add a new enum, be sure it matches the one in table.h, and
	// add an assertion in TableBase.cpp
	public enum VerticalAlignment
	{
		Mixed = -1,
		Top = 0,
		Center = 1,
		Bottom = 2
	}

	// List level's type
	public enum ListLevelType
	{
		Bullet = 0,
		Number = 1,
		Picture = 2
	}

	// List bullet
	public enum ListBullet
	{
		Custom = 0,
		Solid = 0x3f0b7,       // ftcSymbol, 0xf0b7
		Hollow = 0x2006f,      // ftcCourierNew, 0x6f
		Square = 0xaf0a7,      // ftcWingdings, 0xf0a7
		Diamonds = 0xaf076,    // ftcWingdings, 0xf076
		Arrow = 0xaf0d8,       // ftcWingdings, 0xf0d8
		Checkmark = 0xaf0fc    // ftcWingdings, 0xf0fc
	}

	// List numbering
	public enum ListNumbering
	{
		None = -1,         // nfcNil
		Arabic = 0,        // nfcArabic
		UpperRoman = 1,    // nfcUCRoman
		LowerRoman = 2,    // nfcLCRoman
		UpperLetter = 3,   // nfcUCLetter
		LowerLetter = 4    // nfcLCLetter
	}

	public enum Style
	{
		/// <summary>
		/// Mixed styles or other style not in this list.
		/// </summary>
		Other = -1,

		/// <summary>
		/// Reset character and paragraph style to default.
		/// </summary>
		Normal = 0,                // stiNormalPara

		Heading1 = 1,              // stiHeading1
		Heading2 = 2,              // stiHeading2
		Heading3 = 3,              // stiHeading3
		Heading4 = 4,              // stiHeading4
		Heading5 = 5,              // stiHeading5
		Heading6 = 6,              // stiHeading6
		Heading7 = 7,              // stiHeading7
		Heading8 = 8,              // stiHeading8
		Heading9 = 9,              // stiHeading9

		/// <summary>
		/// Table-of-content level 1.
		/// </summary>
		Toc1 = 19,                 // stiToc1
		/// <summary>
		/// Table-of-content level 2.
		/// </summary>
		Toc2 = 20,                 // stiToc2
		/// <summary>
		/// Table-of-content level 3.
		/// </summary>
		Toc3 = 21,                 // stiToc3
		/// <summary>
		/// Table-of-content level 4.
		/// </summary>
		Toc4 = 22,                 // stiToc4
		/// <summary>
		/// Table-of-content level 5.
		/// </summary>
		Toc5 = 23,                 // stiToc5
		/// <summary>
		/// Table-of-content level 6.
		/// </summary>
		Toc6 = 24,                 // stiToc6
		/// <summary>
		/// Table-of-content level 7.
		/// </summary>
		Toc7 = 25,                 // stiToc7
		/// <summary>
		/// Table-of-content level 8.
		/// </summary>
		Toc8 = 26,                 // stiToc8
		/// <summary>
		/// Table-of-content level 9.
		/// </summary>
		Toc9 = 27,                 // stiToc9

		FootnoteText = 29,         // stiFtnText
		Header = 31,               // stiHeader
		Footer = 32,               // stiFooter
		Caption = 34,              // stiCaption
		FootnoteReference = 38,    // stiFtnRef
		EndnoteReference = 42,     // stiEdnRef
		EndnoteText = 43,          // stiEdnText
		Title = 62,                // stiTitle
		Subtitle = 74,             // stiSubtitle
		Hyperlink = 85,            // stiHyperlink
		Strong = 87,               // stiStrong
		Emphasis = 88,             // stiEmphasis
		NoSpacing = 157,           // stiNoSpacing
		ListParagraph = 179,       // stiListParagraph
		Quote = 180,               // stiQuote
		IntenseQuote = 181,        // stiIntenseQuote
		SubtleEmphasis = 260,      // stiSubtleEmphasis
		IntenseEmphasis = 261,     // stiIntenseEmphasis
		SubtleReference = 262,     // stiSubtleReference
		IntenseReference = 263,    // stiIntenseReference
		BookTitle = 264,           // stiBookTitle
		Bibliography = 265,        // stiBibliography

		/// <summary>
		/// Table-of-content heading.
		/// </summary>
		TocHeading = 266,          // stiTocHeading

		// Table styles
		TableGrid = 154,                  // stiTableGrid
		PlainTable1 = 267,                // stiTable15Plain1
		PlainTable2 = 268,                // stiTable15Plain2
		PlainTable3 = 269,                // stiTable15Plain3
		PlainTable4 = 270,                // stiTable15Plain4
		PlainTable5 = 271,                // stiTable15Plain5
		TableGridLight = 272,             // stiTableGridLight

		GridTable1Light = 111,            // stiTable15Grid1Light
		GridTable1Light_Accent1 = 280,    // stiTable15Grid1LightAccent1
		GridTable1Light_Accent2 = 287,    // stiTable15Grid1LightAccent2
		GridTable1Light_Accent3 = 294,    // stiTable15Grid1LightAccent3
		GridTable1Light_Accent4 = 301,    // stiTable15Grid1LightAccent4
		GridTable1Light_Accent5 = 308,    // stiTable15Grid1LightAccent5
		GridTable1Light_Accent6 = 315,    // stiTable15Grid1LightAccent6

		GridTable2 = 274,                 // stiTable15Grid2
		GridTable2_Accent1 = 281,         // stiTable15Grid2Accent1
		GridTable2_Accent2 = 288,         // stiTable15Grid2Accent2
		GridTable2_Accent3 = 295,         // stiTable15Grid2Accent3
		GridTable2_Accent4 = 302,         // stiTable15Grid2Accent4
		GridTable2_Accent5 = 309,         // stiTable15Grid2Accent5
		GridTable2_Accent6 = 316,         // stiTable15Grid2Accent6

		GridTable3 = 275,                 // stiTable15Grid3
		GridTable3_Accent1 = 282,         // stiTable15Grid3Accent1
		GridTable3_Accent2 = 289,         // stiTable15Grid3Accent2
		GridTable3_Accent3 = 296,         // stiTable15Grid3Accent3
		GridTable3_Accent4 = 303,         // stiTable15Grid3Accent4
		GridTable3_Accent5 = 310,         // stiTable15Grid3Accent5
		GridTable3_Accent6 = 317,         // stiTable15Grid3Accent6

		GridTable4 = 276,                 // stiTable15Grid4
		GridTable4_Accent1 = 283,         // stiTable15Grid4Accent1
		GridTable4_Accent2 = 290,         // stiTable15Grid4Accent2
		GridTable4_Accent3 = 297,         // stiTable15Grid4Accent3
		GridTable4_Accent4 = 304,         // stiTable15Grid4Accent4
		GridTable4_Accent5 = 311,         // stiTable15Grid4Accent5
		GridTable4_Accent6 = 318,         // stiTable15Grid4Accent6

		GridTable5Dark = 277,             // stiTable15Grid5Dark
		GridTable5Dark_Accent1 = 284,     // stiTable15Grid5DarkAccent1
		GridTable5Dark_Accent2 = 291,     // stiTable15Grid5DarkAccent2
		GridTable5Dark_Accent3 = 298,     // stiTable15Grid5DarkAccent3
		GridTable5Dark_Accent4 = 305,     // stiTable15Grid5DarkAccent4
		GridTable5Dark_Accent5 = 312,     // stiTable15Grid5DarkAccent5
		GridTable5Dark_Accent6 = 319,     // stiTable15Grid5DarkAccent6

		GridTable6Colorful = 278,             // stiTable15Grid6Colorful
		GridTable6Colorful_Accent1 = 285,     // stiTable15Grid6ColorfulAccent1
		GridTable6Colorful_Accent2 = 292,     // stiTable15Grid6ColorfulAccent2
		GridTable6Colorful_Accent3 = 299,     // stiTable15Grid6ColorfulAccent3
		GridTable6Colorful_Accent4 = 306,     // stiTable15Grid6ColorfulAccent4
		GridTable6Colorful_Accent5 = 313,     // stiTable15Grid6ColorfulAccent5
		GridTable6Colorful_Accent6 = 320,     // stiTable15Grid6ColorfulAccent6

		GridTable7Colorful = 279,             // stiTable15Grid7Colorful
		GridTable7Colorful_Accent1 = 286,     // stiTable15Grid7ColorfulAccent1
		GridTable7Colorful_Accent2 = 293,     // stiTable15Grid7ColorfulAccent2
		GridTable7Colorful_Accent3 = 300,     // stiTable15Grid7ColorfulAccent3
		GridTable7Colorful_Accent4 = 307,     // stiTable15Grid7ColorfulAccent4
		GridTable7Colorful_Accent5 = 314,     // stiTable15Grid7ColorfulAccent5
		GridTable7Colorful_Accent6 = 321,     // stiTable15Grid7ColorfulAccent6

		ListTable1Light = 322,                // stiTable15List1Light
		ListTable1Light_Accent1 = 329,        // stiTable15List1LightAccent1
		ListTable1Light_Accent2 = 336,        // stiTable15List1LightAccent2
		ListTable1Light_Accent3 = 343,        // stiTable15List1LightAccent3
		ListTable1Light_Accent4 = 350,        // stiTable15List1LightAccent4
		ListTable1Light_Accent5 = 357,        // stiTable15List1LightAccent5
		ListTable1Light_Accent6 = 364,        // stiTable15List1LightAccent6

		ListTable2 = 323,                     // stiTable15List2
		ListTable2_Accent1 = 330,             // stiTable15List2Accent1
		ListTable2_Accent2 = 337,             // stiTable15List2Accent2
		ListTable2_Accent3 = 344,             // stiTable15List2Accent3
		ListTable2_Accent4 = 351,             // stiTable15List2Accent4
		ListTable2_Accent5 = 358,             // stiTable15List2Accent5
		ListTable2_Accent6 = 365,             // stiTable15List2Accent6

		ListTable3 = 324,                     // stiTable15List3
		ListTable3_Accent1 = 331,             // stiTable15List3Accent1
		ListTable3_Accent2 = 338,             // stiTable15List3Accent2
		ListTable3_Accent3 = 345,             // stiTable15List3Accent3
		ListTable3_Accent4 = 352,             // stiTable15List3Accent4
		ListTable3_Accent5 = 359,             // stiTable15List3Accent5
		ListTable3_Accent6 = 366,             // stiTable15List3Accent6

		ListTable4 = 325,                     // stiTable15List4
		ListTable4_Accent1 = 332,             // stiTable15List4Accent1
		ListTable4_Accent2 = 339,             // stiTable15List4Accent2
		ListTable4_Accent3 = 346,             // stiTable15List4Accent3
		ListTable4_Accent4 = 353,             // stiTable15List4Accent4
		ListTable4_Accent5 = 360,             // stiTable15List4Accent5
		ListTable4_Accent6 = 367,             // stiTable15List4Accent6

		ListTable5Dark = 326,                 // stiTable15List5Dark
		ListTable5Dark_Accent1 = 333,         // stiTable15List5DarkAccent1
		ListTable5Dark_Accent2 = 340,         // stiTable15List5DarkAccent2
		ListTable5Dark_Accent3 = 347,         // stiTable15List5DarkAccent3
		ListTable5Dark_Accent4 = 354,         // stiTable15List5DarkAccent4
		ListTable5Dark_Accent5 = 361,         // stiTable15List5DarkAccent5
		ListTable5Dark_Accent6 = 368,         // stiTable15List5DarkAccent6

		ListTable6Colorful = 327,             // stiTable15List6Colorful
		ListTable6Colorful_Accent1 = 334,     // stiTable15List6ColorfulAccent1
		ListTable6Colorful_Accent2 = 341,     // stiTable15List6ColorfulAccent2
		ListTable6Colorful_Accent3 = 348,     // stiTable15List6ColorfulAccent3
		ListTable6Colorful_Accent4 = 355,     // stiTable15List6ColorfulAccent4
		ListTable6Colorful_Accent5 = 362,     // stiTable15List6ColorfulAccent5
		ListTable6Colorful_Accent6 = 369,     // stiTable15List6ColorfulAccent6

		ListTable7Colorful = 328,             // stiTable15List7Colorful
		ListTable7Colorful_Accent1 = 335,     // stiTable15List7ColorfulAccent1
		ListTable7Colorful_Accent2 = 342,     // stiTable15List7ColorfulAccent2
		ListTable7Colorful_Accent3 = 349,     // stiTable15List7ColorfulAccent3
		ListTable7Colorful_Accent4 = 356,     // stiTable15List7ColorfulAccent4
		ListTable7Colorful_Accent5 = 363,     // stiTable15List7ColorfulAccent5
		ListTable7Colorful_Accent6 = 370,     // stiTable15List7ColorfulAccent6
	}

	// Document property type
	public enum DocumentPropertyType
	{
		String = 0,
		Number = 1,
		Date = 2,
		Boolean = 3
	}

}   // namespace Microsoft.WordServices
