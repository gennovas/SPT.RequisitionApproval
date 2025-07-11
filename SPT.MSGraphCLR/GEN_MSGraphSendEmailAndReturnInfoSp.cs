using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SqlServer.Server;
using Newtonsoft.Json.Linq;
using System;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;

public partial class StoredProcedures
{
    [SqlProcedure]
    public static void GEN_MSGraphSendEmailAndReturnInfoSp(
        SqlString fromUser,
        SqlString toUser,
        SqlString ccUser,
        SqlString bccUser,
        SqlString subject,
        SqlString body,
        SqlString tableRowPointer,
        out SqlString messageId,
        out SqlString conversationId,
        out SqlString internetMessageId,
        out SqlString createdDateTime,
        out SqlString sentDateTime,
        out SqlString toRecipients,
        out SqlString ccRecipients,
        out SqlString bccRecipients,
        out SqlBoolean hasAttachments,
        out SqlBoolean isReadReceiptRequested,
        out SqlString bodyPreview
    )
    {
        // Set default values for output
        messageId = conversationId = internetMessageId = createdDateTime = sentDateTime =
            toRecipients = ccRecipients = bccRecipients = bodyPreview = SqlString.Null;
        hasAttachments = false;
        isReadReceiptRequested = false;

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

            var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            JArray CreateRecipientObject(string addresses)
            {
                var recipients = new JArray();
                if (!string.IsNullOrWhiteSpace(addresses))
                {
                    var emails = addresses.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var email in emails)
                    {
                        var trimmed = email.Trim();
                        if (!string.IsNullOrEmpty(trimmed))
                        {
                            recipients.Add(new JObject
                            {
                                ["emailAddress"] = new JObject { ["address"] = trimmed }
                            });
                        }
                    }
                }
                return recipients;
            }

            // Build email object
            var emailMessage = new JObject
            {
                ["subject"] = subject.Value,
                ["body"] = new JObject
                {
                    ["contentType"] = "Text",
                    ["content"] = body.Value
                },
                ["toRecipients"] = CreateRecipientObject(toUser.Value),
                ["ccRecipients"] = CreateRecipientObject(ccUser.Value),
                ["bccRecipients"] = CreateRecipientObject(bccUser.Value),
                ["isReadReceiptRequested"] = false // can be changed
            };

            toRecipients = new SqlString(emailMessage["toRecipients"].ToString(Newtonsoft.Json.Formatting.None));
            ccRecipients = new SqlString(emailMessage["ccRecipients"].ToString(Newtonsoft.Json.Formatting.None));
            bccRecipients = new SqlString(emailMessage["bccRecipients"].ToString(Newtonsoft.Json.Formatting.None));

            // attachments
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
                        var fileName = reader["DocumentName"] != DBNull.Value ? reader["DocumentName"].ToString() : $"file_{Guid.NewGuid()}";
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
                        hasAttachments = true;
                        emailMessage["attachments"] = attachments;
                    }
                }
            }

            // Step 1: create message
            var createUrl = $"https://graph.microsoft.com/v1.0/users/{fromUser.Value}/messages";
            var createRes = client.PostAsync(createUrl,
                new StringContent(emailMessage.ToString(), Encoding.UTF8, "application/json")).Result;

            if (!createRes.IsSuccessStatusCode)
            {
                SqlContext.Pipe.Send("ERROR (Create): " + createRes.Content.ReadAsStringAsync().Result);
                return;
            }

            var msgJson = JObject.Parse(createRes.Content.ReadAsStringAsync().Result);

            messageId = new SqlString(msgJson["id"]?.ToString() ?? "");
            conversationId = new SqlString(msgJson["conversationId"]?.ToString() ?? "");
            internetMessageId = new SqlString(msgJson["internetMessageId"]?.ToString() ?? "");
            createdDateTime = new SqlString(msgJson["createdDateTime"]?.ToString() ?? "");
            sentDateTime = SqlString.Null;
            isReadReceiptRequested = msgJson["isReadReceiptRequested"]?.ToObject<bool>() ?? false;
            bodyPreview = new SqlString(msgJson["bodyPreview"]?.ToString() ?? "");

            // Step 2: send message
            var sendUrl = $"https://graph.microsoft.com/v1.0/users/{fromUser.Value}/messages/{messageId.Value}/send";
            var sendRes = client.PostAsync(sendUrl, null).Result;

            if (sendRes.IsSuccessStatusCode)
            {
                sentDateTime = new SqlString(DateTime.UtcNow.ToString("o"));
                SqlContext.Pipe.Send("Email sent successfully.");
            }
            else
            {
                SqlContext.Pipe.Send("ERROR (Send): " + sendRes.Content.ReadAsStringAsync().Result);
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