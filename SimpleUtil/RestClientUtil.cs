using System;
using System.Collections.Generic;
using System.Text;
using System.Net;
using System.IO;

namespace com.github.yedijas.util
{
    class RestClientUtil
    {
        /// <summary>
        /// Upload a multipart with the method name already in request uri string.
        /// PS. this can only upload ad single file from form.
        /// </summary>
        /// <param name="RequestUriString">Request URI complete with the method name.</param>
        /// <param name="Parameters">Parameters to be passed.</param>
        /// <param name="FileInByte">File in byte.</param>
        /// <returns>HttpWebResponse for further process.</returns>
        public static HttpWebResponse UploadMultipart(string RequestUriString,
            List<string> Parameters,
            byte[] FileInByte)
        {
            HttpWebResponse result = UploadMultipart(RequestUriString,
                null,
                Parameters,
                FileInByte);
            return result;
        }

        /// <summary>
        /// Upload a multipart with the method name already in request uri string.
        /// PS. this can only upload ad single file from form.
        /// </summary>
        /// <param name="RequestUriString">Request URI without the method name.</param>
        /// <param name="MethodName">Name of method to be called.</param>
        /// <param name="Parameters">Parameters to be passed.</param>
        /// <param name="FileInByte">File in byte.</param>
        /// <returns>HttpWebResponse for further process.</returns>
        public static HttpWebResponse UploadMultipart(string RequestUriString,
            string MethodName,
            List<string> Parameters,
            byte[] FileInByte)
        {
            HttpWebResponse result = null;
            HttpWebRequest request = CreateWebRequest(
                ConstructUri(RequestUriString, Parameters, MethodName),
                "POST",
                "text/plain");
            request.ContentLength = FileInByte.Length;
            using (Stream requestStream = request.GetRequestStream())
            {
                requestStream.Write(FileInByte, 0, FileInByte.Length);
                requestStream.Close();
            }
            using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)
            {
                result = response;
            }
            return result;
        }

        /// <summary>
        /// Upload a multipart with the method name already in request uri string.
        /// PS. this can only upload ad single file from form.
        /// </summary>
        /// <param name="RequestUriString">Request URI without the method name.</param>
        /// <param name="FileInByte">File in byte.</param>
        /// <returns>HttpWebResponse for further process.</returns>
        public static HttpWebResponse UploadMultipart(string RequestUriString, 
            byte[] FileInByte)
        {
            HttpWebResponse result = UploadMultipart(
                RequestUriString, null, null, FileInByte);
            return result;
        }

        /// <summary>
        /// Construct a uri to be passed to HttpWebRequest.
        /// Method name can be embedded in RequestUriString or not.
        /// </summary>
        /// <param name="RequestUriString">Raw URI. eg. "http://url.com/". </param>
        /// <param name="Parameters">List of string containing parameter to be passed.
        /// Must be in order.</param>
        /// <param name="MethodName">Method name to process request. 
        /// Can be embedded in RequestUriString.</param>
        /// <returns>A request uri.</returns>
        public static string ConstructUri(string RequestUriString,
            List<string> Parameters,
            string MethodName)
        {
            if (!RequestUriString.Substring(RequestUriString.Length - 1, 1).Equals("/"))
            {
                RequestUriString += "/";
            }
            if (!string.IsNullOrEmpty(MethodName))
            {
                RequestUriString = RequestUriString + "/" + MethodName;
            }
            if (Parameters != null && Parameters.Count > 0)
            {
                foreach (string singleParameter in Parameters)
                {
                    RequestUriString = RequestUriString + "/" + singleParameter;
                }
            }
            return RequestUriString;
        }

        /// <summary>
        /// Create a web request to be used to call a method in service provider.
        /// </summary>
        /// <param name="RequestUriString">Request uri string.
        /// This one shuld be compelte along with the methods and params.</param>
        /// <param name="RequestMethod">HTTP method either "GET" or "POST".
        /// "GET" if null or empty is passed.</param>
        /// <param name="ContentType">Content type of request.
        /// "POST" if null or empty is passed.</param>
        /// <returns></returns>
        public static HttpWebRequest CreateWebRequest(string RequestUriString, 
            string RequestMethod,
            string ContentType)
        {
            HttpWebRequest request = HttpWebRequest.Create(RequestUriString) as HttpWebRequest;
            if (string.IsNullOrEmpty(RequestMethod))
            {
                RequestMethod = "GET";
            }
            request.Method = RequestMethod;
            if (string.IsNullOrEmpty(ContentType))
            {
                ContentType = "text/plain";
            }
            request.ContentType = ContentType;
            return request;
        }
    }
}
