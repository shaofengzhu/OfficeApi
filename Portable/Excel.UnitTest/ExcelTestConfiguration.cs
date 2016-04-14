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
				ret = CreateConfigurationInstanceForWacDevMachine();
				// ret = await CreateConfigurationInstanceForGraph(useProdEnvironment: false, filename: "AgaveTest.xlsx");
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
				refreshToken = "AAABAAAAiL9Kn2Z27UubvWFPbm0gLXqWU3D8Nj-lXI3JrMrY7gm-_gD1iYCa3IcS2FrW4yEkYsRfjcLsrH1l025qodeCpbpjGLVrQbHFb-64rKgpl94doKKLIyIU8YIAEwaoMH2do-ZUIcK5UPgv5wNCP_gKJXaiu0maGXS6uOdktd9STFfjNOh4MKYmYIEeJL6U4kgnPIlICzoVBAAwDJXWtK0oKT9HxXJD1Iz5lo45dLzy3ftMTx6wl4ZtQoIJQy4ulbmCCvaA8hFs46XzunvvnKxukWTJfdJK7DQcJbJHNmoOJOyLdwu40jsD6jYOSQ6PFS_TGRDe2FCqSxWjgBYIj1iA4LdHpWNkhF6O-V26dkMLPrytkE4hg0Kfr3kfB1cAtMILyJuCdKJm6Mt8BeIgXxPRvitF7efdiQxcRY8vq_XOPZhy5HKrU4b_lL5YVYrgkfRNdAH0oLx4eCocezpimTf0C8ZtpEkdGpn6toNnhAQ2e5B_dqrr_ytHdkKHpGOvjDjLwikEKFAHz7VVy9-uLI-ONpJQKT-NBdWcUsDjffMXFwuzNlgt4P0RsG_BWT6JSC9apFiYwTcsjC0qJ31BFw1TuKnnr7GXBjkbXkw0st9ao9n1PUturqbLsm8S8qFc8sQbq3tPxtHbYXu292ROMzERohcjs6OD_hMpZhTqtfYGUP4gAA";
			} 
			else
			{
				clientId = "09d9cc54-6048-4c79-b468-99aa29c6e98d";
				refreshToken = "AAABAAAAo3ZCPl0FaU2WWRdLWLHpercJWcdVVO9k0TEoGH_oXnMqGLxnlQsxfrHKkujhJgF1K_3P8cRqVdjlJp9xTTSiKmty9MHZeGvLraX1se3Y9YQlEzjP3pPL3LOUzDafk-C7qBmcbcp0KSHy7ZvbP_e7KWj7sN8c1_UHYSPp7ppQ1mWUJWQY-hfNDOaYfszZvrNZyEUELEosKWkjbk0sG2FqTUycy-4wtIMYezwx3FzwU-y_XsjrP2YotKWwHJHrSmV0U08IZ5UeWc-QLX4WtNsEguHwyc7n-QmW6s_Ph6S0LC0jf8o8SXuL6g-Q9Tvz2-cXekhJlKblrf--CZeK7bfCzfMMn9ZMuiWJSSK7URFFnMQJpvtizFokrF-DsNXfkqnAd3n_cVGFZjXpVMUNenoRpg1y85NLZ9KbDrwBNJIYJVWfh8mTAzzKcbzp6621HpJJCbRk9IqRg0uLOUguzcleUKBSH4Mzy3LscblFWyvTQidUpnz3InN7Ltw2-SWytmHAbHKlYahSPgM_chFMySCQ0k1JTerK4IkxdHk3lQvIxNsw-Vdenp0SM6LqqCSAqt-DSCcTQs-rxSUTIOkOHJtuUWuiHT64JDBt6rUe-0si5zcOxhDjJF1oNMqP-EgTKUIjHbudZ2ljX8GYaANs9wL_oNNl-RoYr5yfK7pSWIHP_ywgAA";
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
