using System;
using System.Threading.Tasks;
using System.IO;
using System.Collections.Generic;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using SendGrid.Helpers.Mail;
using OfficeOpenXml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.Azure.Documents.Client;
using Microsoft.Azure.Documents.Linq;
namespace Company.Function

{
    public static class SendEmail
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




        // const string query = "SELECT * FROM c WHERE c.payment_date BETWEEN DateTimeAdd(\"hh\", -48, GetCurrentDateTime())  AND GetCurrentDateTime()";

        [FunctionName("SendEmailTimer")]
        [return: SendGrid(ApiKey = "SendGridApiKey")]
        public static async Task<SendGridMessage> Run([TimerTrigger("%MessageQueuerOccurence%")] TimerInfo myTimer,
        [CosmosDB("outDatabase", "WebhookCollection", ConnectionStringSetting = "CosmosDbConnectionString")] DocumentClient webhookDocument,
        // IEnumerable<object> inputDocument,
        [CosmosDB("outDatabase", "UserCollection", ConnectionStringSetting = "CosmosDbConnectionString")] DocumentClient userDocument,
        ILogger log)
        {


            log.LogInformation($"SendEmailTimer executed at: {DateTime.Now}");

            DateTime nowDate = DateTime.Now;


            DateTime fromDate = nowDate.AddHours(-24);

            Uri collectionUri = UriFactory.CreateDocumentCollectionUri("outDatabase", "WebhookCollection");
            string q = "SELECT * FROM c WHERE c.payment_date BETWEEN '" + fromDate.ToString("yyyy-MM-dd HH:mm:ssfffffff") + "'  AND '" + nowDate.ToString("yyyy-MM-dd HH:mm:ssfffffff") + "'";

            IDocumentQuery<dynamic> query = userDocument.CreateDocumentQuery(collectionUri, q,
                            new FeedOptions
                            {
                                PopulateQueryMetrics = true,
                                MaxItemCount = -1,
                                MaxDegreeOfParallelism = -1,
                                EnableCrossPartitionQuery = true
                            }

                            ).AsDocumentQuery();
            FeedResponse<dynamic> sqlResult = await query.ExecuteNextAsync();

            var msg = new SendGridMessage()
            {
                From = new EmailAddress(Environment.GetEnvironmentVariable("SenderEmail")),

                Subject = "Transactions Report from " + fromDate.ToString("yyyy-MM-dd_HH:mm") + " to " + nowDate.ToString("yyyy-MM-dd_HH:mm"),

                PlainTextContent = "and easy to do anywhere, especially with C#",
                HtmlContent = "and easy to do anywhere, <strong>especially with C#</strong>"
            };

            msg.AddTo(new EmailAddress(Environment.GetEnvironmentVariable("RecipientEmail")));


            SendGrid.Helpers.Mail.Attachment att = new SendGrid.Helpers.Mail.Attachment
            {
                Content = Convert.ToBase64String(ReportHelper.CallGenerateExcelReport(sqlResult, userDocument, log)),
                Filename = "Report-from-" + fromDate.ToString("yyyy-MM-dd_HH:mm:ss") + "-to-" + nowDate.ToString("yyyy-MM-dd_HH:mm:ss") + ".xlsx",
                Type = "application/vnd.openxmlformats-officedocument.spreadsheet.excel",
                Disposition = "attachment"

            };
            msg.AddAttachment(att);

            return msg;
        }
    }
}