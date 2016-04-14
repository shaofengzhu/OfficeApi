using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Http;
using Newtonsoft.Json.Linq;

namespace OfficeExtension
{
	public static class GraphApiUtility
	{
		public static async Task<GraphApiFileInfo> GetFileInfo(
			bool useProductionEnvironment, 
			string filename,
			string clientId,
			string clientSecret,
			string refreshToken)
		{
			string graphRootUrl = "";

			if (useProductionEnvironment)
			{
				graphRootUrl = "https://graph.microsoft.com/testexcel";
			}
			else
			{
				graphRootUrl = "https://graph.microsoft-ppe.com/testexcel";
			}

			string accessToken = await GetAccessTokenFromRefreshToken(useProductionEnvironment, clientId, clientSecret, refreshToken);

			HttpClient httpClient = new HttpClient();
			HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Get, graphRootUrl + "/me/drive/root/children");
			httpRequestMessage.Headers.Add("Authorization", "Bearer " + accessToken);
			HttpResponseMessage httpResponseMessage = await httpClient.SendAsync(httpRequestMessage);
			if (httpResponseMessage.StatusCode != HttpStatusCode.OK)
			{
				throw new Exception("Cannot get files");
			}
			string responseBody = await httpResponseMessage.Content.ReadAsStringAsync();
			JObject jsonBody = JObject.Parse(responseBody);
			JArray jsonFiles = jsonBody["value"] as JArray;
			if (jsonFiles == null)
			{
				throw new Exception("Cannot get files");
			}

			string fileId = "";
			foreach (JToken jsonFile in jsonFiles)
			{
				if (jsonFile.Type == JTokenType.Object &&
					jsonFile["name"] != null &&
					jsonFile["name"].Type == JTokenType.String &&
					string.Equals(jsonFile["name"].Value<string>(), filename, StringComparison.OrdinalIgnoreCase))
				{
					fileId = jsonFile["id"].Value<string>();
					break;
				}
			}

			if (string.IsNullOrEmpty(fileId))
			{
				throw new Exception("Cannot find file");
			}

			GraphApiFileInfo ret = new GraphApiFileInfo()
			{
				FileId = fileId,
				AccessToken = accessToken,
				FileWorkbookUrl = graphRootUrl + "/me/drive/items/" + fileId + "/workbook"
			};

			return ret;
		}

		private static Task<string> GetAccessTokenFromRefreshToken(
			bool useProductionEnvironment,
			string clientId,
			string clientSecret,
			string refreshToken)
		{
			string tokenServiceUrl;
			if (useProductionEnvironment)
			{
				tokenServiceUrl = "https://login.windows.net/common/oauth2/token";
			}
			else
			{
				tokenServiceUrl = "https://login.windows-ppe.net/common/oauth2/token";
			}

			return GetAccessTokenFromRefreshToken(tokenServiceUrl, clientId, clientSecret, refreshToken);
		}

		private static async Task<string> GetAccessTokenFromRefreshToken(
			string tokenServiceUrl,
			string clientId,
			string clientSecret,
			string refreshToken)
		{
			HttpClient httpClient = new HttpClient();
			HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, tokenServiceUrl);
			httpRequestMessage.Content = new FormUrlEncodedContent(new Dictionary<string, string>()
			{
				{ "grant_type", "refresh_token" },
				{ "refresh_token", refreshToken },
				{ "client_id", clientId },
			});
			HttpResponseMessage httpResponseMessage = await httpClient.SendAsync(httpRequestMessage);
			if (httpResponseMessage.StatusCode != HttpStatusCode.OK)
			{
				throw new Exception("Cannot get access token");
			}
			string body = await httpResponseMessage.Content.ReadAsStringAsync();
			JObject jsonBody = JObject.Parse(body);
			return jsonBody["access_token"].Value<string>();
		}
	}
}
