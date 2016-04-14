using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Http;

namespace OfficeExtension
{
    public class HttpRequestExecutor: IRequestExecutor
    {
        public HttpRequestExecutor()
        {

        }
        public async Task<RequestExecutorResponseMessage> Execute(RequestExecutorRequestMessage request)
        {
            HttpClient client = new HttpClient();
            HttpRequestMessage httpRequest = new HttpRequestMessage(HttpMethod.Post, request.Url);
            foreach (var pair in request.Headers)
            {
                httpRequest.Headers.Add(pair.Key, pair.Value);
            }
            string body = string.Empty;
            if (request.Body != null)
            {
                body = request.Body;
            }

			string contentType = request.ContentType;
			if (string.IsNullOrEmpty(contentType))
			{
				contentType = Constants.ContentTypeApplicationJson;
			}

            httpRequest.Content = new StringContent(body, Encoding.UTF8, contentType);
            HttpResponseMessage httpResponse = await client.SendAsync(httpRequest);

            RequestExecutorResponseMessage ret = new RequestExecutorResponseMessage();
            ret.StatusCode = httpResponse.StatusCode;
            foreach (var pair in httpResponse.Headers)
            {
                ret.Headers[pair.Key] = string.Join(",", pair.Value.ToArray());
            }

            ret.Body = await httpResponse.Content.ReadAsStringAsync();
            return ret;
        }
    }
}
