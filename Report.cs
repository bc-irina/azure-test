using System;
using System.Text;
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
using System.Security.Cryptography;
using Microsoft.Azure.Documents.Client;

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

        const string query = "SELECT * FROM c WHERE c.payment_date_utc BETWEEN {from} AND CONCAT({to},' 23:59:60')";
        [FunctionName("Report")]
        public static async Task<HttpResponseMessage> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "report/{from}/{to}")] HttpRequestMessage req,
            [CosmosDB("outDatabase", "WebhookCollection", ConnectionStringSetting = "CosmosDbConnectionString")] DocumentClient inputDocument,
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


                //update empty utc
                ReportHelper.updatePaymentUTC(inputDocument, log);

                response = new HttpResponseMessage(HttpStatusCode.OK);
                //Set the Excel document content response


                DateTime fromDate = DateTime.Parse(from).ToUniversalTime();
                DateTime toDate = DateTime.Parse(to + " 23:59:59").ToUniversalTime();
                IEnumerable<object> sqlResult = await ReportHelper.queryDB(inputDocument, fromDate, toDate, log);
                response.Content = new ByteArrayContent(ReportHelper.CallGenerateExcelReport(sqlResult, userDocument, log));
                //Set the contentDisposition as attachment
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                {
                    FileName = "Report-from-" + from + "-to-" + to + ".xlsx"
                };
                //Set the content type as xlsx format mime type
                response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheet.excel");

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