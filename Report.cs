using System;
using System.Text;
using System.IO;
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
namespace Company.Function
{
    public enum ReportHeader
    {
        user_id,
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
                        foreach (ReportHeader a in Enum.GetValues(typeof(ReportHeader)))
                        {
                            string cellValue = responseBody[a.ToString()] != null ? responseBody[a.ToString()].ToString().Replace("\n", "").Replace("\r", "") : null;
                            worksheet.Cells[row, (int)a + 1].Value = cellValue;
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
    }
}
