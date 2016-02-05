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
	public class Chart: OfficeExtension.ClientObject
	{
		private Microsoft.ExcelServices.ChartAxes m_axes;
		private Microsoft.ExcelServices.ChartDataLabels m_dataLabels;
		private Microsoft.ExcelServices.ChartAreaFormat m_format;
		private double m_height;
		private double m_left;
		private Microsoft.ExcelServices.ChartLegend m_legend;
		private string m_name;
		private Microsoft.ExcelServices.ChartSeriesCollection m_series;
		private Microsoft.ExcelServices.ChartTitle m_title;
		private double m_top;
		private double m_width;
		private Microsoft.ExcelServices.Worksheet m_worksheet;

		/* Begin_PlaceHolder_Chart_Custom_Members */
		/* End_PlaceHolder_Chart_Custom_Members */
		public Chart(OfficeExtension.ClientRequestContext context, OfficeExtension.ObjectPath objectPath)
			: base(context, objectPath)
		{
		}
		
		
		public Microsoft.ExcelServices.ChartAxes Axes
		{
			get
			{
				if (this.m_axes == null)
				{
					this.m_axes = new Microsoft.ExcelServices.ChartAxes(this.Context, OfficeExtension.ObjectPathFactory._CreatePropertyObjectPath(this.Context, this, "Axes", false /*isCollection*/, false /*isInvalidAfterRequest*/));	
				}
		
				return this.m_axes;
			}
		}
		
		public Microsoft.ExcelServices.ChartDataLabels DataLabels
		{
			get
			{
				if (this.m_dataLabels == null)
				{
					this.m_dataLabels = new Microsoft.ExcelServices.ChartDataLabels(this.Context, OfficeExtension.ObjectPathFactory._CreatePropertyObjectPath(this.Context, this, "DataLabels", false /*isCollection*/, false /*isInvalidAfterRequest*/));	
				}
		
				return this.m_dataLabels;
			}
		}
		
		public Microsoft.ExcelServices.ChartAreaFormat Format
		{
			get
			{
				if (this.m_format == null)
				{
					this.m_format = new Microsoft.ExcelServices.ChartAreaFormat(this.Context, OfficeExtension.ObjectPathFactory._CreatePropertyObjectPath(this.Context, this, "Format", false /*isCollection*/, false /*isInvalidAfterRequest*/));	
				}
		
				return this.m_format;
			}
		}
		
		public Microsoft.ExcelServices.ChartLegend Legend
		{
			get
			{
				if (this.m_legend == null)
				{
					this.m_legend = new Microsoft.ExcelServices.ChartLegend(this.Context, OfficeExtension.ObjectPathFactory._CreatePropertyObjectPath(this.Context, this, "Legend", false /*isCollection*/, false /*isInvalidAfterRequest*/));	
				}
		
				return this.m_legend;
			}
		}
		
		public Microsoft.ExcelServices.ChartSeriesCollection Series
		{
			get
			{
				if (this.m_series == null)
				{
					this.m_series = new Microsoft.ExcelServices.ChartSeriesCollection(this.Context, OfficeExtension.ObjectPathFactory._CreatePropertyObjectPath(this.Context, this, "Series", true /*isCollection*/, false /*isInvalidAfterRequest*/));	
				}
		
				return this.m_series;
			}
		}
		
		public Microsoft.ExcelServices.ChartTitle Title
		{
			get
			{
				if (this.m_title == null)
				{
					this.m_title = new Microsoft.ExcelServices.ChartTitle(this.Context, OfficeExtension.ObjectPathFactory._CreatePropertyObjectPath(this.Context, this, "Title", false /*isCollection*/, false /*isInvalidAfterRequest*/));	
				}
		
				return this.m_title;
			}
		}
		
		public Microsoft.ExcelServices.Worksheet Worksheet
		{
			get
			{
				if (this.m_worksheet == null)
				{
					this.m_worksheet = new Microsoft.ExcelServices.Worksheet(this.Context, OfficeExtension.ObjectPathFactory._CreatePropertyObjectPath(this.Context, this, "Worksheet", false /*isCollection*/, false /*isInvalidAfterRequest*/));	
				}
		
				return this.m_worksheet;
			}
		}

		public double Height
		{
			get
			{
				OfficeExtension.Utility._ThrowIfNotLoaded("height", this.m_height);
				return this.m_height;
			}

			set
			{
				this.m_height = value;
				OfficeExtension.ActionFactory._CreateSetPropertyAction(this.Context, this, "Height", value);
			}
		}

		public double Left
		{
			get
			{
				OfficeExtension.Utility._ThrowIfNotLoaded("left", this.m_left);
				return this.m_left;
			}

			set
			{
				this.m_left = value;
				OfficeExtension.ActionFactory._CreateSetPropertyAction(this.Context, this, "Left", value);
			}
		}

		public string Name
		{
			get
			{
				OfficeExtension.Utility._ThrowIfNotLoaded("name", this.m_name);
				return this.m_name;
			}

			set
			{
				this.m_name = value;
				OfficeExtension.ActionFactory._CreateSetPropertyAction(this.Context, this, "Name", value);
			}
		}

		public double Top
		{
			get
			{
				OfficeExtension.Utility._ThrowIfNotLoaded("top", this.m_top);
				return this.m_top;
			}

			set
			{
				this.m_top = value;
				OfficeExtension.ActionFactory._CreateSetPropertyAction(this.Context, this, "Top", value);
			}
		}

		public double Width
		{
			get
			{
				OfficeExtension.Utility._ThrowIfNotLoaded("width", this.m_width);
				return this.m_width;
			}

			set
			{
				this.m_width = value;
				OfficeExtension.ActionFactory._CreateSetPropertyAction(this.Context, this, "Width", value);
			}
		}

		public void Delete()
		{
			/* Begin_PlaceHolder_Chart_Delete */
			/* End_PlaceHolder_Chart_Delete */
			OfficeExtension.ActionFactory._CreateMethodAction(this.Context, this, "Delete", OfficeExtension.OperationType.Default, new object[] {});
		}

		public OfficeExtension.ClientResult< string > GetImage(int width, int height, string fittingMode)
		{
			/* Begin_PlaceHolder_Chart_GetImage */
			/* End_PlaceHolder_Chart_GetImage */
			var action = OfficeExtension.ActionFactory._CreateMethodAction(this.Context, this, "GetImage", OfficeExtension.OperationType.Read, new object[] {width, height, fittingMode});
			var ret = new OfficeExtension.ClientResult< string >();
			OfficeExtension.Utility._AddActionResultHandler(this, action, ret);
			return ret;
		}

		public void SetData(object sourceData, string seriesBy)
		{
			/* Begin_PlaceHolder_Chart_SetData */
			/* End_PlaceHolder_Chart_SetData */
			OfficeExtension.ActionFactory._CreateMethodAction(this.Context, this, "SetData", OfficeExtension.OperationType.Default, new object[] {sourceData, seriesBy});
		}

		public void SetPosition(object startCell, object endCell)
		{
			/* Begin_PlaceHolder_Chart_SetPosition */
			/* End_PlaceHolder_Chart_SetPosition */
			OfficeExtension.ActionFactory._CreateMethodAction(this.Context, this, "SetPosition", OfficeExtension.OperationType.Default, new object[] {startCell, endCell});
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
			if (!OfficeExtension.Utility._IsUndefined(obj["Height"]))
			{
				this.m_height = obj["Height"].ToObject<double>();
			}
		
			if (!OfficeExtension.Utility._IsUndefined(obj["Left"]))
			{
				this.m_left = obj["Left"].ToObject<double>();
			}
		
			if (!OfficeExtension.Utility._IsUndefined(obj["Name"]))
			{
				this.m_name = obj["Name"].ToObject<string>();
			}
		
			if (!OfficeExtension.Utility._IsUndefined(obj["Top"]))
			{
				this.m_top = obj["Top"].ToObject<double>();
			}
		
			if (!OfficeExtension.Utility._IsUndefined(obj["Width"]))
			{
				this.m_width = obj["Width"].ToObject<double>();
			}
		
		    if (!OfficeExtension.Utility._IsUndefined(obj["Axes"]))
			{
		        this.Axes._HandleResult(obj["Axes"]);
			}
		    if (!OfficeExtension.Utility._IsUndefined(obj["DataLabels"]))
			{
		        this.DataLabels._HandleResult(obj["DataLabels"]);
			}
		    if (!OfficeExtension.Utility._IsUndefined(obj["Format"]))
			{
		        this.Format._HandleResult(obj["Format"]);
			}
		    if (!OfficeExtension.Utility._IsUndefined(obj["Legend"]))
			{
		        this.Legend._HandleResult(obj["Legend"]);
			}
		    if (!OfficeExtension.Utility._IsUndefined(obj["Series"]))
			{
		        this.Series._HandleResult(obj["Series"]);
			}
		    if (!OfficeExtension.Utility._IsUndefined(obj["Title"]))
			{
		        this.Title._HandleResult(obj["Title"]);
			}
		    if (!OfficeExtension.Utility._IsUndefined(obj["Worksheet"]))
			{
		        this.Worksheet._HandleResult(obj["Worksheet"]);
			}
		}
		
		/*
		 * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
		 */
		public Microsoft.ExcelServices.Chart Load(OfficeExtension.LoadOption option) 
		{
			OfficeExtension.Utility._Load(this, option);
			return this;
		}
	}
}
