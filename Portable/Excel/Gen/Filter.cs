﻿/*
 * This is a generated file. 
 * If there are content placeholders, only edit content inside content placeholders.
 * If there are no content placeholders, do not edit this file directly.
 */
namespace Microsoft.ExcelServices
{
	using System;
	/* Begin_PlaceHolder_UsingHeader */
	/* End_PlaceHolder_UsingHeader */

	/* Begin_PlaceHolder_Header */
	/* End_PlaceHolder_Header */
	public class Filter: OfficeExtension.ClientObject
	{
		private Microsoft.ExcelServices.FilterCriteria m_criteria;

		/* Begin_PlaceHolder_Filter_Custom_Members */
		/* End_PlaceHolder_Filter_Custom_Members */
		public Filter(OfficeExtension.ClientRequestContext context, OfficeExtension.ObjectPath objectPath)
			: base(context, objectPath)
		{
		}
		

		public Microsoft.ExcelServices.FilterCriteria Criteria
		{
			get
			{
				OfficeExtension.Utility._ThrowIfNotLoaded("criteria", this.m_criteria);
				return this.m_criteria;
			}
		}

		public void Apply(Microsoft.ExcelServices.FilterCriteria criteria)
		{
			/* Begin_PlaceHolder_Filter_Apply */
			/* End_PlaceHolder_Filter_Apply */
			OfficeExtension.ActionFactory._CreateMethodAction(this.Context, this, "Apply", OfficeExtension.OperationType.Default, new object[] {criteria});
		}

		public void ApplyBottomItemsFilter(int count)
		{
			/* Begin_PlaceHolder_Filter_ApplyBottomItemsFilter */
			/* End_PlaceHolder_Filter_ApplyBottomItemsFilter */
			OfficeExtension.ActionFactory._CreateMethodAction(this.Context, this, "ApplyBottomItemsFilter", OfficeExtension.OperationType.Default, new object[] {count});
		}

		public void ApplyBottomPercentFilter(int percent)
		{
			/* Begin_PlaceHolder_Filter_ApplyBottomPercentFilter */
			/* End_PlaceHolder_Filter_ApplyBottomPercentFilter */
			OfficeExtension.ActionFactory._CreateMethodAction(this.Context, this, "ApplyBottomPercentFilter", OfficeExtension.OperationType.Default, new object[] {percent});
		}

		public void ApplyCellColorFilter(string color)
		{
			/* Begin_PlaceHolder_Filter_ApplyCellColorFilter */
			/* End_PlaceHolder_Filter_ApplyCellColorFilter */
			OfficeExtension.ActionFactory._CreateMethodAction(this.Context, this, "ApplyCellColorFilter", OfficeExtension.OperationType.Default, new object[] {color});
		}

		public void ApplyCustomFilter(string criteria1, string criteria2, string oper)
		{
			/* Begin_PlaceHolder_Filter_ApplyCustomFilter */
			/* End_PlaceHolder_Filter_ApplyCustomFilter */
			OfficeExtension.ActionFactory._CreateMethodAction(this.Context, this, "ApplyCustomFilter", OfficeExtension.OperationType.Default, new object[] {criteria1, criteria2, oper});
		}

		public void ApplyDynamicFilter(string criteria)
		{
			/* Begin_PlaceHolder_Filter_ApplyDynamicFilter */
			/* End_PlaceHolder_Filter_ApplyDynamicFilter */
			OfficeExtension.ActionFactory._CreateMethodAction(this.Context, this, "ApplyDynamicFilter", OfficeExtension.OperationType.Default, new object[] {criteria});
		}

		public void ApplyFontColorFilter(string color)
		{
			/* Begin_PlaceHolder_Filter_ApplyFontColorFilter */
			/* End_PlaceHolder_Filter_ApplyFontColorFilter */
			OfficeExtension.ActionFactory._CreateMethodAction(this.Context, this, "ApplyFontColorFilter", OfficeExtension.OperationType.Default, new object[] {color});
		}

		public void ApplyIconFilter(Microsoft.ExcelServices.Icon icon)
		{
			/* Begin_PlaceHolder_Filter_ApplyIconFilter */
			/* End_PlaceHolder_Filter_ApplyIconFilter */
			OfficeExtension.ActionFactory._CreateMethodAction(this.Context, this, "ApplyIconFilter", OfficeExtension.OperationType.Default, new object[] {icon});
		}

		public void ApplyTopItemsFilter(int count)
		{
			/* Begin_PlaceHolder_Filter_ApplyTopItemsFilter */
			/* End_PlaceHolder_Filter_ApplyTopItemsFilter */
			OfficeExtension.ActionFactory._CreateMethodAction(this.Context, this, "ApplyTopItemsFilter", OfficeExtension.OperationType.Default, new object[] {count});
		}

		public void ApplyTopPercentFilter(int percent)
		{
			/* Begin_PlaceHolder_Filter_ApplyTopPercentFilter */
			/* End_PlaceHolder_Filter_ApplyTopPercentFilter */
			OfficeExtension.ActionFactory._CreateMethodAction(this.Context, this, "ApplyTopPercentFilter", OfficeExtension.OperationType.Default, new object[] {percent});
		}

		public void ApplyValuesFilter(object[] values)
		{
			/* Begin_PlaceHolder_Filter_ApplyValuesFilter */
			/* End_PlaceHolder_Filter_ApplyValuesFilter */
			OfficeExtension.ActionFactory._CreateMethodAction(this.Context, this, "ApplyValuesFilter", OfficeExtension.OperationType.Default, new object[] {values});
		}

		public void Clear()
		{
			/* Begin_PlaceHolder_Filter_Clear */
			/* End_PlaceHolder_Filter_Clear */
			OfficeExtension.ActionFactory._CreateMethodAction(this.Context, this, "Clear", OfficeExtension.OperationType.Default, new object[] {});
		}

			/** Handle results returned from the document
			 */
		public override void _HandleResult(Newtonsoft.Json.Linq.JToken value)
		{
			if (OfficeExtension.Utility._IsNullOrUndefined(value))
			{
				return;
			}
			Newtonsoft.Json.Linq.JObject obj = value as Newtonsoft.Json.Linq.JObject;
			if (obj == null)
			{
				return;
			}
		
			OfficeExtension.Utility._FixObjectPathIfNecessary(this, obj);
			if (!OfficeExtension.Utility._IsUndefined(obj["Criteria"]))
			{
				this.m_criteria = obj["Criteria"].ToObject<Microsoft.ExcelServices.FilterCriteria>();
			}
		
		}
		
		/*
		 * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
		 */
		public Microsoft.ExcelServices.Filter Load(OfficeExtension.LoadOption option) 
		{
			OfficeExtension.Utility._Load(this, option);
			return this;
		}
	}
}
