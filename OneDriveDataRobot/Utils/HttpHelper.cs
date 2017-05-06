/*
 * OneDrive Data Robot - Sample Code
 * Copyright (c) Microsoft Corporation
 * All rights reserved. 
 * 
 * MIT License
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy of 
 * this software and associated documentation files (the ""Software""), to deal in 
 * the Software without restriction, including without limitation the rights to use, 
 * copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the
 * Software, and to permit persons to whom the Software is furnished to do so, 
 * subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all 
 * copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, 
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A 
 * PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT 
 * HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION 
 * OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE 
 * SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */

namespace OneDriveDataRobot
{
    using Models;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Net.Http;
    using System.Threading.Tasks;

    public class HttpHelper
    {
        private const int MAX_UPLOAD_SIZE = 4 * 1024 * 1024;
        private HttpClient httpClient = new HttpClient();
        public static readonly HttpHelper Default = new HttpHelper();

      

        #region Metadata operations
        /// <summary>
        /// Update an existing item with the changes specified in patchBody.
        /// </summary>
        public async Task<T> PatchAsync<T>(object patchBody, string itemUrl, string accessToken) where T : class
        {
            var forcePatch = new Dictionary<string, string>();
            forcePatch["X-HTTP-Method"] = "PATCH";

            return await PostAsync<T>(patchBody, itemUrl, accessToken, forcePatch);
        }

        /// <summary>
        /// Post an object to a URL and return the response converted back into an object.
        /// </summary>
        public async Task<T> PostAsync<T>(object body, string itemUrl, string accessToken, Dictionary<string,string> additionalHeaders = null) where T : class
        {
            var requestMessage = new HttpRequestMessage(HttpMethod.Post, itemUrl);
            if (additionalHeaders != null)
            {
                foreach(var header in additionalHeaders)
                {
                    requestMessage.Headers.TryAddWithoutValidation(header.Key, header.Value);
                }
            }
            if (!string.IsNullOrEmpty(accessToken))
            {
                requestMessage.Headers.TryAddWithoutValidation("Authorization", "Bearer " + accessToken);
            }

            if (null != body)
            {
                var contentText = Newtonsoft.Json.JsonConvert.SerializeObject(body);
                var content = new StringContent(contentText, System.Text.Encoding.UTF8, "application/json");
                requestMessage.Content = content;
            }

            var responseMessage = await httpClient.SendAsync(requestMessage);
            responseMessage.EnsureSuccessStatusCode();

            return await ParseJsonFromResponseAsync<T>(responseMessage);
        }


        /// <summary>
        /// Return a new object of type T by performing a GET on the URL and converting to an object.
        /// </summary>
        public async Task<T> GetAsync<T>(string requestUri, string accessToken) where T : class
        {
            var requestMessage = new HttpRequestMessage(HttpMethod.Get, requestUri);

            if (!string.IsNullOrEmpty(accessToken))
            {
                requestMessage.Headers.TryAddWithoutValidation("Authorization", "Bearer " + accessToken);
            }

            var responseMessage = await httpClient.SendAsync(requestMessage);

            responseMessage.EnsureSuccessStatusCode();
            return await ParseJsonFromResponseAsync<T>(responseMessage);
        }

        #endregion

        #region Utility functions

        /// <summary>
        /// Convert the contents of an HttpResponseMessage into an object by using a JSON parser.
        /// </summary>
        private async Task<T> ParseJsonFromResponseAsync<T>(HttpResponseMessage response)
        {
            if (response.Content.Headers.ContentType.MediaType.ToLower() != "application/json")
            {
                throw new InvalidOperationException($"MediaType for the response message was {response.Content.Headers.ContentType.MediaType} instead of \"application/json\".");
            }

            var responseData = await response.Content.ReadAsStringAsync();
            return Newtonsoft.Json.JsonConvert.DeserializeObject<T>(responseData);
        }
        #endregion

    }
}
