using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeExtension;

namespace Excel.UnitTest
{
	internal class ExcelTestConfiguration
	{
		private static ExcelTestConfiguration s_instance;

		public string Url
		{
			get;
			private set;
		}

		private string ClientId
		{
			get;
			set;
		}

		private string RefreshToken
		{
			get;
			set;
		}

		public string AccessToken
		{
			get;
			private set;
		}

		public void SetupRequestContext(ClientRequestContext requestContext)
		{
			if (!string.IsNullOrEmpty(this.AccessToken))
			{
				requestContext.RequestHeaders["Authorization"] = "Bearer " + this.AccessToken;
			}
		}

		public static async Task<ExcelTestConfiguration> GetInstance()
		{
			ExcelTestConfiguration ret = s_instance;
			if (ret == null)
			{
				// ret = CreateConfigurationInstanceForWacDevMachine();
				ret = await CreateConfigurationInstanceForGraph(useProdEnvironment: false, filename: "AgaveTest.xlsx");
				// ret = await CreateConfigurationInstanceForGraph(useProdEnvironment: true, filename: "AgaveTest.xlsx");
				s_instance = ret;
			}

			return ret;
		}

		private static ExcelTestConfiguration CreateConfigurationInstanceForWacDevMachine()
		{
			ExcelTestConfiguration ret = new ExcelTestConfiguration();
			ret.Url = "http://shaozhu-ttvm8.redmond.corp.microsoft.com/th/WacRest.ashx/transport_wopi/Application_Excel/wachost_/Fi_anonymous~AgaveTest.xlsx/ak_1%7CGN=R3Vlc3Q=&SN=MTQ3NzYwMzQ5Mw==&IT=NTI0NzY0ODQ4NzI5ODI1NTE3Mw==&PU=MTQ3NzYwMzQ5Mw==&SR=YW5vbnltb3Vz&TZ=MTExOQ==&SA=RmFsc2U=&LE=RmFsc2U=&AG=VHJ1ZQ==&RH=yowrURj8H-VliWbIQ9aPYap80Goy0B5oU0dfzY7WEwM=/_api";
			return ret;
		}

		private static async Task<ExcelTestConfiguration> CreateConfigurationInstanceForGraph(bool useProdEnvironment, string filename)
		{
			ExcelTestConfiguration ret = new ExcelTestConfiguration();
			string clientId = "";
			string refreshToken = "";
			string clientSecret = "";
			if (useProdEnvironment)
			{
				clientId = "8563463e-ea18-4355-9297-41ff32200164";
				refreshToken = "AAABAAAAiL9Kn2Z27UubvWFPbm0gLVfeKP2hZcZ86a6lp2m4bh3dqnZBSxRvzCvCxo9KFBV1u9PmnK85adjyZvkRyEkITe5o9yrHbdf7TkS3OXzn1V3_KaOLQnvQDwezFlGSoDmx_oBb-R09ayJ_X1ukk6W8nLgjmxQU4-f4xtGEjNjin-VXtjsPoQ6oecBByZOWwaTrA1q4ypcBC2U3N0JSKI3wsrdt6BOpI1HPlR0iNLN0EdLNCO4FanLRj9pX9I3rDuOTl4ij6eaTNBj9VUhMjuJAKsWgqbWU3BnElF_WApmxVk3dKSqqoozgHSPW8quU7Zzl1xfEF3N04lzguXgzbrMvfPUQsqnb30BM8O-wBUOuoVlAcok-moYZhmPgc6Nrx6b8ecBMIQxpBGVHnQjDdDwZF1yhI_sQS5cw8mjkegAXnQAjcJDZd6kHr68007ab_5EoX_0XmzqBsFhwzXI2fD3yI6qsLsPlk9MwBXZrKitj_0FFP6YbHrO-hOdJGqlcXsagoPhKt2jTccdTPdDemXzcGojNOzO94Od2OD93Ask0U2RxvIglOv44AylPMhcu_ZNmg415E0FbpQFPo2zHl31mZKeYuTPHJOEHsh03jpTeZGCeNgM9m2N0cxQdmNzkpucVgeLw_lkYksRIOJu7xR7rwx7B4wIcbVOKjp6wSMchM_UgAA";
			} 
			else
			{
				clientId = "09d9cc54-6048-4c79-b468-99aa29c6e98d";
				refreshToken = "AAABAAAAo3ZCPl0FaU2WWRdLWLHpeuZwRI99asNFFGqI5jrtOaZrpea7p2Tjv7lBqu9pLvvEeWZ8Cqt-7ZsIDrUwRG0fA7NN087iBjX1sfs90I-uTROatDX4iMls7CFUwqO6SmUvsugPFIBQL9g3Gab956jIJJd9IyN5Zds24Ff-WJwb07ISCzq4akW8VtJOn54aRilOQbbsGGDUX4AeYb_Lazre3J3LAG_O8egjPrYrhf-ks9OjUzePoyxrntxGGs9h_wWIVnFPyaXaStNvik4MoihmPNAo6DzWekTKD8EjY7FlwiQKYnLyEGBVb1rilFHMX2WW6up0uoCQ-JOAiT3zXJ2FX2tnjBzt4JCLsTYxl4QSHu6BMRwzHNPK-LOOSNleLa3k0chaFpIjhpC7i_C0z-aRiSJXfw9BN_ZEOd79XXWcEdhsiNfQutFYRY9eNIn8SaS-fCt_yqD46JE2yIeE0FjK9X89qE2vLiOLKPJJcwtraLEiyKTsAlYrVSldsBX4bv-j7sQOwAHb62Ys2rBAFXGfErEEW1xN-G3j2sNe9jSlJAZWvzWKyhP6N4Z4VM1-Wp4LRpFykclpVI5EiKJzq90hgBytEj9bXKIj3Iqn2Tmd1rJfHTfvrqcdi2cSjaAcSKygzu4jWr2DXNZ1PPi50bz7tsE2_dxYzpIYOJtgFaboV2YgAA";
			}

			GraphApiFileInfo fileInfo = await GraphApiUtility.GetFileInfo(useProdEnvironment, filename, clientId, clientSecret, refreshToken);
			ret.ClientId = clientId;
			ret.RefreshToken = refreshToken;
			ret.AccessToken = fileInfo.AccessToken;
			ret.Url = fileInfo.FileWorkbookUrl;

			return ret;
		}
	}
}
