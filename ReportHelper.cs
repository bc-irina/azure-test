using System;

using System.IO;
using System.Collections.Generic;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using SendGrid.Helpers.Mail;
using OfficeOpenXml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

using System.Text;

using System.Collections;
using System.Threading.Tasks;

using Microsoft.Azure.WebJobs.Extensions.Http;

using System.Net.Http;
using System.Net;
using System.Net.Http.Headers;

using Microsoft.Extensions.Configuration;

using System.Security.Cryptography;
using Microsoft.Azure.Documents.Client;
using Microsoft.Azure.Documents.Linq;
namespace Company.Function

{

    public static class ReportHelper
    {

        public static byte[] excelAttachment(IEnumerable<object> inputDocument)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            MemoryStream memorystream = new MemoryStream();
            using (ExcelPackage excel = new ExcelPackage())
            {

                var worksheet = excel.Workbook.Worksheets.Add("Report1");

                // build exel header
                var headerRow = new List<string[]>();
                var arr = Enum.GetNames(typeof(ReportHeader));
                headerRow.Add(arr);
                String upRange = (headerRow[0].Length + 64) > 90 ? "A" + Char.ConvertFromUtf32(headerRow[0].Length + 38) : Char.ConvertFromUtf32(headerRow[0].Length + 64);
                string headerRange = "A1:" + upRange + "1";
                worksheet.Cells[headerRange].LoadFromArrays(headerRow);

                // build report content
                int row = 2;
                int numCol = worksheet.Dimension.Columns;
                foreach (var documentItem in inputDocument)
                {
                    var json = JsonConvert.SerializeObject(documentItem);
                    JToken responseBody = JToken.FromObject(documentItem);
                    foreach (ReportHeader a in Enum.GetValues(typeof(ReportHeader)))
                    {
                        string cellValue = responseBody[a.ToString()] != null ? responseBody[a.ToString()].ToString().Replace("\n", "").Replace("\r", "") : null;
                        worksheet.Cells[row, (int)a + 1].Value = cellValue;
                    }
                    row = row + 1;
                }

                excel.SaveAs(memorystream);

            }
            return memorystream.ToArray();
        }


        public static byte[] CallGenerateExcelReport(IEnumerable<object> inputDocument, DocumentClient userDocument, ILogger log)
        {
            var task = generateExcelReport(inputDocument, userDocument, log);
            task.Wait();
            var result = task.Result;
            return result;
        }
        public static async Task<byte[]> generateExcelReport(IEnumerable<object> inputDocument, DocumentClient userDocument, ILogger log)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            MemoryStream memorystream = new MemoryStream();
            using (ExcelPackage excel = new ExcelPackage())
            {
                var worksheet = excel.Workbook.Worksheets.Add("Report1");

                // build exel header
                var headerRow = new List<string[]>();
                var arr = Enum.GetNames(typeof(ReportHeader));
                headerRow.Add(arr);
                String upRange = (headerRow[0].Length + 64) > 90 ? "A" + Char.ConvertFromUtf32(headerRow[0].Length + 38) : Char.ConvertFromUtf32(headerRow[0].Length + 64);
                string headerRange = "A1:" + upRange + "1";
                worksheet.Cells[headerRange].LoadFromArrays(headerRow);

                // build report content
                int row = 2;
                int numCol = worksheet.Dimension.Columns;
                foreach (var documentItem in inputDocument)
                {
                    var json = JsonConvert.SerializeObject(documentItem);
                    JToken responseBody = JToken.FromObject(documentItem);
                    string username = null;
                    string instanceName = null;
                    string service = null;
                    try
                    {
                        service = responseBody["service"].ToString();

                    }
                    catch
                    {
                        service = "payline";
                    }

                    if (service == null || service == "payline")
                    {
                        string uuid = responseBody["user_id"].ToString();

                        Uri collectionUri = UriFactory.CreateDocumentCollectionUri("outDatabase", "UserCollection");
                        IDocumentQuery<dynamic> query = userDocument.CreateDocumentQuery(collectionUri,
                        "SELECT * FROM c WHERE c.id='" + uuid + "'",
                        new FeedOptions
                        {
                            PopulateQueryMetrics = true,
                            MaxItemCount = -1,
                            MaxDegreeOfParallelism = -1,
                            EnableCrossPartitionQuery = true
                        }

                        ).AsDocumentQuery();
                        //query if user already in db
                        FeedResponse<dynamic> sqlResult = await query.ExecuteNextAsync();

                        if (sqlResult.Count == 0)
                        {
                            //call user api to get user_name and instance
                            JToken userInfo = getUserInfo(uuid, log);
                            if (userInfo != null)
                            {
                                username = userInfo["user_name"].ToString();
                                instanceName = userInfo["instance"].ToString();

                                await userDocument.CreateDocumentAsync(collectionUri, userInfo);

                            }
                        }
                        else
                        {
                            //populate user_name and instance from sql query
                            JToken jsonResult = JToken.FromObject(sqlResult);
                            username = jsonResult[0]["user_name"].ToString();
                            instanceName = jsonResult[0]["instance"].ToString();
                        }


                    }

                    foreach (ReportHeader a in Enum.GetValues(typeof(ReportHeader)))

                    {
                        string cellValue = responseBody[a.ToString()] != null ? responseBody[a.ToString()].ToString().Replace("\n", "").Replace("\r", "") : null;
                        worksheet.Cells[row, (int)a + 1].Value = cellValue;

                    }
                    if (username != null & instanceName != null)
                    {
                        worksheet.Cells[row, (int)ReportHeader.user_name + 1].Value = username;
                        worksheet.Cells[row, (int)ReportHeader.instance + 1].Value = instanceName;

                    }
                    worksheet.Cells[row, (int)ReportHeader.service + 1].Value = service;
                    row = row + 1;
                }

                excel.SaveAs(memorystream);
            }
            return memorystream.ToArray();

        }

        private static string getSignature(string date, string accessSecret)
        {
            SortedList ParamsList = new SortedList{
            { "timestamp", date }
            };

            string canonicalQuery = string.Empty;
            foreach (DictionaryEntry keyValuePair in ParamsList)
            {
                canonicalQuery = canonicalQuery + keyValuePair.Key.ToString() + "="
                + Uri.EscapeDataString(keyValuePair.Value.ToString()) + "&";
            }
            string finalQuery = canonicalQuery.Replace("+",
            "%20").Remove(canonicalQuery.Length - 1, 1);
            string payshieldSignature = hmacDigest(finalQuery, accessSecret);
            return payshieldSignature;

        }
        private static string hmacDigest(string msg, string keyString)
        {
            byte[] keyByte = Encoding.ASCII.GetBytes(keyString);
            byte[] messageBytes = Encoding.ASCII.GetBytes(msg);
            HMACSHA256 hmacsha256 = new HMACSHA256(keyByte);
            byte[] hashmessage = hmacsha256.ComputeHash(messageBytes);
            return Convert.ToBase64String(hashmessage);
        }

        private static JToken getUserInfo(string uuid, ILogger log)
        {
            string accessKey;
            string accessSecret;
            string apiEndpoint;
            try
            {
                accessKey = Environment.GetEnvironmentVariable("ACCESS_KEY");
                accessSecret = Environment.GetEnvironmentVariable("ACCESS_SECRET");
                apiEndpoint = Environment.GetEnvironmentVariable("USER_API_ENDPOINT");
                if (accessKey == null || accessSecret == null || apiEndpoint == null)
                {
                    log.LogWarning("API Secrets not found");
                    return null;
                }

            }
            catch (Exception e)
            {
                log.LogError(e.Message);
                return null;
            }

            var httpClient = new HttpClient();
            var settings = new JsonSerializerSettings
            {
                DateFormatString =
                    "yyyy-MM-ddTHH:mm:ss.fffZ"
            };
            var d = JsonConvert.SerializeObject(DateTime.UtcNow, settings);
            var date = d.Replace("\"", "");
            string signature = getSignature(date, accessSecret);

            httpClient.DefaultRequestHeaders.Add("Payshield-Signature", signature);
            httpClient.DefaultRequestHeaders.Add("Payshield-Key", accessKey);
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var url = apiEndpoint + "payline/users/" + uuid;
            var builder = new UriBuilder(url);
            builder.Port = -1;
            var query = System.Web.HttpUtility.ParseQueryString(builder.Query);
            query["timestamp"] = date;

            builder.Query = query.ToString();
            string finalUrl = builder.ToString();
            var response = httpClient.GetAsync(finalUrl).Result;

            if (response.IsSuccessStatusCode)
            {
                var dataObjects = response.Content.ReadAsAsync<object>().Result;
                JToken jsonResponse = JToken.FromObject(dataObjects);

                return jsonResponse;
            }
            return null;
        }


        public static DateTime convertDateToUTC(string payment_date)
        {
            payment_date = payment_date.Substring(0, payment_date.LastIndexOf(' '));
            payment_date = payment_date.Insert(payment_date.Length - 2, ":");

            DateTime convertedDate = DateTime.Parse(payment_date).ToUniversalTime();
            //string convertedString = convertDate(payment_date).ToString("yyyy-MM-dd HH:mm:ss");
            return convertedDate;

        }

    }
}