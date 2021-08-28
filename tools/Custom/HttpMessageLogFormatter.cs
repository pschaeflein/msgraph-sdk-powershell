﻿// ------------------------------------------------------------------------------
//  Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

namespace Microsoft.Graph.PowerShell
{
    using Hyak.Common;
    using Newtonsoft.Json;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Xml.Linq;

    public static class HttpMessageLogFormatter
    {
        public static string GetHttpRequestLog(HttpRequestMessage request)
        {
            if (request == null) return string.Empty;

            string body = (request.Content == null) ? string.Empty : FormatString(request.Content.ReadAsStringAsync().Result);

            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.AppendLine($"============================ HTTP REQUEST ============================{Environment.NewLine}");
            stringBuilder.AppendLine(string.Format("HTTP Method:{0}{1}{0}", Environment.NewLine, request.Method.ToString()));
            stringBuilder.AppendLine(string.Format("Absolute Uri:{0}{1}{0}", Environment.NewLine, request.RequestUri.ToString()));
            stringBuilder.AppendLine($"Headers:{Environment.NewLine}{HeadersToString(ConvertHttpHeadersToCollection(request.Headers))}");
            stringBuilder.AppendLine(string.Format("Body:{0}{1}{0}", Environment.NewLine, SanitizeBody(body)));
            return stringBuilder.ToString();
        }

        public static string GetHttpResponseLog(HttpResponseMessage response)
        {
            if (response == null) return string.Empty;

            string body = (response.Content == null) ? string.Empty : FormatString(response.Content.ReadAsStringAsync().Result);

            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.AppendLine($"============================ HTTP RESPONSE ============================{Environment.NewLine}");
            stringBuilder.AppendLine($"Status Code:{Environment.NewLine}{response.StatusCode}{Environment.NewLine}");
            stringBuilder.AppendLine($"Headers:{Environment.NewLine}{HeadersToString(ConvertHttpHeadersToCollection(response.Headers))}{Environment.NewLine}");
            stringBuilder.AppendLine($"Body:{Environment.NewLine}{SanitizeBody(body)}{Environment.NewLine}");
            return stringBuilder.ToString();
        }

        private static object SanitizeBody(string body)
        {
            IList<Regex> regexList = new List<Regex>();
            // Remove access_token:*  instances in body.
            Regex regexPattern = new Regex("(\\s*\"access_token\"\\s*:\\s*)\"[^\"]+\"");
            regexList.Add(regexPattern);

            foreach (Regex matcher in regexList)
            {
                body = matcher.Replace(body, "$1\"<redacted>\"");
            }

            return body;
        }

        private static IDictionary<string, IEnumerable<string>> ConvertHttpHeadersToCollection(HttpHeaders headers)
        {
            headers.Remove("Authorization");
            return headers.ToDictionary(a => a.Key, a => a.Value);
        }

        private static object HeadersToString(IDictionary<string, IEnumerable<string>> headers)
        {
            StringBuilder stringBuilder = headers.Aggregate(new StringBuilder(),
                (sb, kvp) => sb.AppendLine(string.Format("{0,-30}: {1}", kvp.Key, String.Join(",", kvp.Value.ToArray()))));

            if (stringBuilder.Length > 0)
                stringBuilder.Remove(stringBuilder.Length - 2, 2);

            return stringBuilder.ToString();
        }

        private static string FormatString(string content)
        {
            try
            {
                if (CloudException.IsJson(content))
                {
                    return JsonConvert.SerializeObject(JsonConvert.DeserializeObject(content), Formatting.Indented);
                }
                if (CloudException.IsXml(content))
                {
                    return XDocument.Parse(content).ToString();
                }
            }
            catch
            {
                return content;
            }

            if (content.Length > 10240)
            {
                return content.Substring(0, 10240) + "\r\nDATA TRUNCATED DUE TO SIZE\r\n";
            }

            return content;
        }
    }
}
