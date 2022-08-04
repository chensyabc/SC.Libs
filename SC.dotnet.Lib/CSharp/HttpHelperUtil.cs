using Newtonsoft.Json.Linq;
using System.IO;
using System.Net;
using System.Text;

namespace SC.dotnet.Lib.CSharp
{
    public class HttpHelperUtil
    {
        /*
         *  Sample Call
            Uri mainUri = new Uri(apiBaseUrl);
            string callPath = $"{mainUri.AbsolutePath.TrimEnd('/')}/odata/Get(env='{environment}')";    // Api call suffix.
            url = new Uri(mainUri, callPath).AbsoluteUri;       // Form complete api call url by merging base and call path.

            string response = HttpHelperUtil.GetHttpResponse(url, null);
        */

        /// <summary>
        /// Call api to get http response
        /// </summary>
        public static string GetHttpResponse(string url, WebHeaderCollection headers)
        {
            HttpWebRequest webRequest = WebRequest.Create(url) as HttpWebRequest;

            webRequest.Method = WebRequestMethods.Http.Get;
            webRequest.ContentType = "application/json";
            webRequest.Headers = headers ?? new WebHeaderCollection();

            HttpWebResponse webResponse = webRequest.GetResponse() as HttpWebResponse;
            return GetResponseString(webResponse);
        }

        /// <summary>
        /// Call with OData
        /// </summary>
        private static string GetResponseString(HttpWebResponse webResponse)
        {
            using (var stream = webResponse.GetResponseStream())
            {
                string response = new StreamReader(stream, Encoding.UTF8).ReadToEnd();

                if (response.Contains("odata.metadata") || response.Contains("odata.context"))     // If odata response.. 
                {
                    response = JObject.Parse(response)["value"].ToString();     // .. capture only value part.
                }

                return response;
            }
        }
    }
}
