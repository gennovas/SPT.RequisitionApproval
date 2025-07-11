using System;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using Microsoft.SqlServer.Server;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json.Linq;
using System.Net;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

public partial class StoredProcedures
{
    [SqlProcedure]
    public static void GEN_MSGraphGetEmailSp(SqlString userEmail, SqlString subjectFilter)
    {
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

        try
        {
            var tenantId = "74e90f9e-b525-4f67-a69a-0b15eead3f74";
            var clientId = "f10fffb3-22ca-4eef-8e99-e432cf04bcfc";
            var clientSecret = "5uJ8Q~YkUkYwZLsBqzn5r8geAnHkhljJjPnZlaOy";

            var authority = $"https://login.microsoftonline.com/{tenantId}";
            var authContext = new AuthenticationContext(authority);
            var clientCred = new ClientCredential(clientId, clientSecret);

            var result = authContext.AcquireTokenAsync("https://graph.microsoft.com/", clientCred).Result;
            var accessToken = result.AccessToken;

            var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            var metaData = new SqlMetaData[]
            {
                new SqlMetaData("Id", SqlDbType.NVarChar, 200),
                new SqlMetaData("Subject", SqlDbType.NVarChar, 4000),
                new SqlMetaData("ReceivedDateTime", SqlDbType.DateTime),
                new SqlMetaData("FromEmail", SqlDbType.NVarChar, 500),
                new SqlMetaData("ToRecipients", SqlDbType.NVarChar, 2000),
                new SqlMetaData("Importance", SqlDbType.NVarChar, 50),
                new SqlMetaData("ConversationId", SqlDbType.NVarChar, 200),
                new SqlMetaData("BodyPreview", SqlDbType.NVarChar, 4000)
            };

            SqlContext.Pipe.SendResultsStart(new SqlDataRecord(metaData));
             
            string newFilter = subjectFilter.Value;
            string rawFilter = $"contains(subject, '{newFilter}')";
            string encodedFilter = Uri.EscapeDataString(rawFilter);
            string url = $"https://graph.microsoft.com/v1.0/users/{userEmail}/mailFolders/inbox/messages?$filter={encodedFilter}";

            while (!string.IsNullOrEmpty(url))
            {
                var response = httpClient.GetAsync(url).Result;
                var json = response.Content.ReadAsStringAsync().Result;
                var data = JObject.Parse(json);

                var messages = data["value"] as JArray;
                if (messages != null)
                {
                    foreach (var msg in messages)
                    {
                        var subject = msg["subject"]?.ToString() ?? "";
                        var record = new SqlDataRecord(metaData);

                        record.SetString(0, msg["id"]?.ToString() ?? "");
                        record.SetString(1, subject);
                        record.SetDateTime(2, msg["receivedDateTime"]?.ToObject<DateTime?>() ?? DateTime.MinValue);
                        record.SetString(3, msg["from"]?["emailAddress"]?["address"]?.ToString() ?? "");

                        // Check for toRecipients being null or empty
                        var toRecipients = msg["toRecipients"] as JArray;
                        string toList = "";
                        if (toRecipients != null)
                        {
                            // Manually loop through the toRecipients
                            foreach (var recipient in toRecipients)
                            {
                                var emailAddress = recipient["emailAddress"]?["address"]?.ToString();
                                if (!string.IsNullOrEmpty(emailAddress))
                                {
                                    if (!string.IsNullOrEmpty(toList))
                                    {
                                        toList += ";";
                                    }
                                    toList += emailAddress;
                                }
                            }
                        }

                        record.SetString(4, toList);
                        record.SetString(5, msg["importance"]?.ToString() ?? "");
                        record.SetString(6, msg["conversationId"]?.ToString() ?? "");
                        record.SetString(7, msg["bodyPreview"]?.ToString() ?? "");

                        SqlContext.Pipe.SendResultsRow(record);
                        //}
                    }
                }

                url = data["@odata.nextLink"]?.ToString(); // Get the next page link if present
            }

            SqlContext.Pipe.SendResultsEnd();
        }
        catch (Exception ex)
        {
            /*
            SqlContext.Pipe.Send("ERROR: " + ex.Message);
            if (ex.InnerException != null)
                SqlContext.Pipe.Send("INNER: " + ex.InnerException.Message);
            */
            throw ex.InnerException;
        }
    }
}