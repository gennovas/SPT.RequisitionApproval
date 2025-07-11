using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SqlServer.Server;
using Newtonsoft.Json.Linq;
using System;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;

public partial class StoredProcedures
{
    [SqlProcedure]
    public static void GEN_MSGraphSendEmailSp(SqlString fromUser, SqlString toUser, SqlString subject, SqlString body, SqlString tableRowPointer)
    {
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

        try
        {
            var tenantId = "74e90f9e-b525-4f67-a69a-0b15eead3f74";
            var clientId = "f10fffb3-22ca-4eef-8e99-e432cf04bcfc";
            var clientSecret = "5uJ8Q~YkUkYwZLsBqzn5r8geAnHkhljJjPnZlaOy";

            var authority = $"https://login.microsoftonline.com/{tenantId}";
            var authContext = new AuthenticationContext(authority);
            var credentials = new ClientCredential(clientId, clientSecret);

            var result = authContext.AcquireTokenAsync("https://graph.microsoft.com/", credentials).Result;
            var accessToken = result.AccessToken;

            //SqlContext.Pipe.Send("AccessToken: " + accessToken); // for debug

            var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);


            var toRecipients = new JArray();

            if (!string.IsNullOrWhiteSpace(toUser.Value))
            {
                var emails = toUser.Value
                    .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var email in emails)
                {
                    var trimmedEmail = email.Trim();

                    if (!string.IsNullOrWhiteSpace(trimmedEmail))
                    {
                        toRecipients.Add(new JObject
                        {
                            ["emailAddress"] = new JObject
                            {
                                ["address"] = trimmedEmail
                            }
                        });
                    }
                }

                if (toRecipients.Count == 0)
                {
                    SqlContext.Pipe.Send("ERROR: No valid email addresses found in 'toUser'.");
                    return;
                }
            }
            else
            {
                SqlContext.Pipe.Send("ERROR: 'toUser' is required.");
                return;
            }

            var emailMessage = new JObject
            {
                ["message"] = new JObject
                {
                    ["subject"] = subject.Value,
                    ["body"] = new JObject
                    {
                        ["contentType"] = "Text",
                        ["content"] = body.Value
                    },
                    ["toRecipients"] = toRecipients
                },
                ["saveToSentItems"] = true
            };

            if (!tableRowPointer.IsNull)
            {
                using (var conn = new SqlConnection("context connection=true"))
                {
                    conn.Open();
                    var cmd = new SqlCommand(@"
                    SELECT DocumentObject, DocumentName, MediaType
                    FROM DocumentObjectAndRefView
                    WHERE TableRowPointer = @RowPointer", conn);

                    cmd.Parameters.AddWithValue("@RowPointer", tableRowPointer.Value);

                    var reader = cmd.ExecuteReader();
                    var attachments = new JArray();

                    while (reader.Read())
                    {
                        var fileBytes = (byte[])reader["DocumentObject"];
                        var fileName = reader["DocumentName"] != DBNull.Value ? reader["DocumentName"].ToString() : $"attachment_{Guid.NewGuid()}";
                        var mediaType = reader["MediaType"] != DBNull.Value ? reader["MediaType"].ToString() : "application/octet-stream";
                        var base64Content = Convert.ToBase64String(fileBytes);

                        attachments.Add(new JObject
                        {
                            ["@odata.type"] = "#microsoft.graph.fileAttachment",
                            ["name"] = fileName,
                            ["contentType"] = mediaType,
                            ["contentBytes"] = base64Content
                        });
                    }

                    reader.Close();

                    if (attachments.Count > 0)
                    {
                        emailMessage["message"]["attachments"] = attachments;
                    }
                }
            }

            var requestUrl = $"https://graph.microsoft.com/v1.0/users/{fromUser.Value}/sendMail";

            var response = client.PostAsync(requestUrl, new StringContent(emailMessage.ToString(), System.Text.Encoding.UTF8, "application/json")).Result;

            if (response.IsSuccessStatusCode)
            {
                SqlContext.Pipe.Send("Email sent successfully.");
            }
            else
            {
                var errorMsg = response.Content.ReadAsStringAsync().Result;
                SqlContext.Pipe.Send("ERROR: " + errorMsg);
            }
        }
        catch (Exception ex)
        {
            SqlContext.Pipe.Send("EXCEPTION: " + ex.Message);
            if (ex.InnerException != null)
                SqlContext.Pipe.Send("INNER: " + ex.InnerException.Message);
        }
    }
}
