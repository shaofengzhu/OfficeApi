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
	public class ChartAxis: OfficeExtension.ClientObject
	{
		private Microsoft.ExcelServices.ChartAxisFormat m_format;
		private Microsoft.ExcelServices.ChartGridlines m_majorGridlines;
		private object m_majorUnit;
		private object m_maximum;
		private object m_minimum;
		private Microsoft.ExcelServices.ChartGridlines m_minorGridlines;
		private object m_minorUnit;
		private Microsoft.ExcelServices.ChartAxisTitle m_title;

		/* Begin_PlaceHolder_ChartAxis_Custom_Members */
		/* End_PlaceHolder_ChartAxis_Custom_Members */
		public ChartAxis(OfficeExtension.ClientRequestContext context, OfficeExtension.ObjectPath objectPath)
			: base(context, objectPath)
		{
		}
		
		
		public Microsoft.ExcelServices.ChartAxisFormat Format
		{
			get
			{
				if (this.m_format == null)
				{
					this.m_format = new Microsoft.ExcelServices.ChartAxisFormat(this.Context, OfficeExtension.ObjectPathFactory._CreatePropertyObjectPath(this.Context, this, "Format", false /*isCollection*/, false /*isInvalidAfterRequest*/));	
				}
		
				return this.m_format;
			}
		}
		
		public Microsoft.ExcelServices.ChartGridlines MajorGridlines
		{
			get
			{
				if (this.m_majorGridlines == null)
				{
					this.m_majorGridlines = new Microsoft.ExcelServices.ChartGridlines(this.Context, OfficeExtension.ObjectPathFactory._CreatePropertyObjectPath(this.Context, this, "MajorGridlines", false /*isCollection*/, false /*isInvalidAfterRequest*/));	
				}
		
				return this.m_majorGridlines;
			}
		}
		
		public Microsoft.ExcelServices.ChartGridlines MinorGridlines
		{
			get
			{
				if (this.m_minorGridlines == null)
				{
					this.m_minorGridlines = new Microsoft.ExcelServices.ChartGridlines(this.Context, OfficeExtension.ObjectPathFactory._CreatePropertyObjectPath(this.Context, this, "MinorGridlines", false /*isCollection*/, false /*isInvalidAfterRequest*/));	
				}
		
				return this.m_minorGridlines;
			}
		}
		
		public Microsoft.ExcelServices.ChartAxisTitle Title
		{
			get
			{
				if (this.m_title == null)
				{
					this.m_title = new Microsoft.ExcelServices.ChartAxisTitle(this.Context, OfficeExtension.ObjectPathFactory._CreatePropertyObjectPath(this.Context, this, "Title", false /*isCollection*/, false /*isInvalidAfterRequest*/));	
				}
		
				return this.m_title;
			}
		}

		public object MajorUnit
		{
			get
			{
				OfficeExtension.Utility._ThrowIfNotLoaded(this, "majorUnit", this.m_majorUnit);
				return this.m_majorUnit;
			}

			set
			{
				this.m_majorUnit = value;
				OfficeExtension.ActionFactory._CreateSetPropertyAction(this.Context, this, "MajorUnit", value);
			}
		}

		public object Maximum
		{
			get
			{
				OfficeExtension.Utility._ThrowIfNotLoaded(this, "maximum", this.m_maximum);
				return this.m_maximum;
			}

			set
			{
				this.m_maximum = value;
				OfficeExtension.ActionFactory._CreateSetPropertyAction(this.Context, this, "Maximum", value);
			}
		}

		public object Minimum
		{
			get
			{
				OfficeExtension.Utility._ThrowIfNotLoaded(this, "minimum", this.m_minimum);
				return this.m_minimum;
			}

			set
			{
				this.m_minimum = value;
				OfficeExtension.ActionFactory._CreateSetPropertyAction(this.Context, this, "Minimum", value);
			}
		}

		public object MinorUnit
		{
			get
			{
				OfficeExtension.Utility._ThrowIfNotLoaded(this, "minorUnit", this.m_minorUnit);
				return this.m_minorUnit;
			}

			set
			{
				this.m_minorUnit = value;
				OfficeExtension.ActionFactory._CreateSetPropertyAction(this.Context, this, "MinorUnit", value);
			}
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
			if (!OfficeExtension.Utility._IsUndefined(obj["MajorUnit"]))
			{
				this.LoadedPropertyNames.Add("MajorUnit");
				this.m_majorUnit = obj["MajorUnit"].ToObject<object>();
			}
		
			if (!OfficeExtension.Utility._IsUndefined(obj["Maximum"]))
			{
				this.LoadedPropertyNames.Add("Maximum");
				this.m_maximum = obj["Maximum"].ToObject<object>();
			}
		
			if (!OfficeExtension.Utility._IsUndefined(obj["Minimum"]))
			{
				this.LoadedPropertyNames.Add("Minimum");
				this.m_minimum = obj["Minimum"].ToObject<object>();
			}
		
			if (!OfficeExtension.Utility._IsUndefined(obj["MinorUnit"]))
			{
				this.LoadedPropertyNames.Add("MinorUnit");
				this.m_minorUnit = obj["MinorUnit"].ToObject<object>();
			}
		
		    if (!OfficeExtension.Utility._IsUndefined(obj["Format"]))
			{
		        this.Format._HandleResult(obj["Format"]);
			}
		    if (!OfficeExtension.Utility._IsUndefined(obj["MajorGridlines"]))
			{
		        this.MajorGridlines._HandleResult(obj["MajorGridlines"]);
			}
		    if (!OfficeExtension.Utility._IsUndefined(obj["MinorGridlines"]))
			{
		        this.MinorGridlines._HandleResult(obj["MinorGridlines"]);
			}
		    if (!OfficeExtension.Utility._IsUndefined(obj["Title"]))
			{
		        this.Title._HandleResult(obj["Title"]);
			}
		}
		
		/*
		 * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
		 */
		public Microsoft.ExcelServices.ChartAxis Load(OfficeExtension.LoadOption option) 
		{
			OfficeExtension.Utility._Load(this, option);
			return this;
		}
	}
}

