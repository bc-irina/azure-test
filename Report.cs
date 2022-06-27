using System;
using System.Text;
using System.IO;
using System.Collections;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Net.Http;
using System.Net;
using System.Net.Http.Headers;
using Newtonsoft.Json.Linq;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using System.Security.Cryptography;
using Microsoft.Azure.Documents.Client;
using Microsoft.Azure.Documents.Linq;

namespace Company.Function
{
    public enum ReportHeader
    {
        user_id,
        user_name,
        instance,
        service,
        response_code,
        response_text,
        approved,
        transaction_reference,
        txn_reference,
        receipt_number,
        transaction_type,
        token,
        merchant_reference,
        crn1,
        crn2,
        crn3,
        masked_card_number,
        payment_date,
        custom_id_name,
        custom_id_value,
        amount,
        currency,
        expiry_date,
        cardholder_name,
        card_type,
        cardholder_address_street,
        cardholder_address_street2,
        cardholder_address_city,
        cardholder_address_state,
        cardholder_address_postcode,
        cardholder_address_country,
        metadata
    }
    public static class Report
    {
        static HttpResponseMessage response;

        const string query = "SELECT * FROM c WHERE c.payment_date BETWEEN {from} AND CONCAT({to},' 23:59:60')";
        [FunctionName("Report")]
        public static async Task<HttpResponseMessage> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "report/{from}/{to}")] HttpRequestMessage req,
            [CosmosDB("outDatabase", "WebhookCollection", ConnectionStringSetting = "CosmosDbConnectionString", SqlQuery = query)] IEnumerable<object> inputDocument,
            [CosmosDB("outDatabase", "UserCollection", ConnectionStringSetting = "CosmosDbConnectionString")] DocumentClient userDocument,
            string from,
            string to,
            ILogger log)
        {

            var authHeader = req.Headers.Authorization;
            if (authHeader != null && authHeader.ToString().StartsWith("Basic"))
            {
                string encodedUsernamePassword = authHeader.ToString().Substring("Basic ".Length).Trim();

                //the coding should be iso or you could use ASCII and UTF-8 decoder
                Encoding encoding = Encoding.GetEncoding("iso-8859-1");
                string usernamePassword = encoding.GetString(Convert.FromBase64String(encodedUsernamePassword));
                var config = new ConfigurationBuilder()
                            .AddEnvironmentVariables()
                            .Build();
                string userNameKeyVault = Environment.GetEnvironmentVariable("WebHookAuth", EnvironmentVariableTarget.Process);
                if (usernamePassword != userNameKeyVault)
                {
                    return new HttpResponseMessage(HttpStatusCode.Unauthorized);
                }
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage excel = new ExcelPackage())
                {
                    MemoryStream memorystream = new MemoryStream();
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
                        string username = null;
                        string instanceName = null;
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
                        row = row + 1;
                    }

                    excel.SaveAs(memorystream);
                    response = new HttpResponseMessage(HttpStatusCode.OK);
                    //Set the Excel document content response
                    response.Content = new ByteArrayContent(memorystream.ToArray());
                    //Set the contentDisposition as attachment
                    response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                    {
                        FileName = "Report-from-" + from + "-to-" + to + ".xlsx"
                    };
                    //Set the content type as xlsx format mime type
                    response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheet.excel");
                }
                return response;
            }
            else
            {
                return new HttpResponseMessage(HttpStatusCode.Unauthorized);
            }
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
    }
}