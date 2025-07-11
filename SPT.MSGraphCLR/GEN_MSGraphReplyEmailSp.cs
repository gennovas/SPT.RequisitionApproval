
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SqlServer.Server;
using Newtonsoft.Json.Linq;
using System;
using System.Data.SqlTypes;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;

public partial class StoredProcedures
{
    [SqlProcedure]
    public static void GEN_MSGraphReplyEmailSp(
        SqlString fromUser,
        SqlString originalMessageId,
        SqlString replyBody,
        SqlBoolean replyAll,
        out SqlString replyMessageId,
        out SqlString conversationId,
        out SqlString internetMessageId,
        out SqlString createdDateTime,
        out SqlString sentDateTime
    )
    {
        replyMessageId = conversationId = internetMessageId = createdDateTime = sentDateTime = SqlString.Null;

        try
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

            string tenantId = "74e90f9e-b525-4f67-a69a-0b15eead3f74";
            string clientId = "f10fffb3-22ca-4eef-8e99-e432cf04bcfc";
            string clientSecret = "5uJ8Q~YkUkYwZLsBqzn5r8geAnHkhljJjPnZlaOy";
            string authority = $"https://login.microsoftonline.com/{tenantId}";

            var authContext = new AuthenticationContext(authority);
            var credentials = new ClientCredential(clientId, clientSecret);
            var result = authContext.AcquireTokenAsync("https://graph.microsoft.com/", credentials).Result;
            var accessToken = result.AccessToken;

            var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            // Step 1: Create draft reply (normal or replyAll)
            string replyType = replyAll.IsTrue ? "createReplyAll" : "createReply";
            string createReplyUrl = $"https://graph.microsoft.com/v1.0/users/{fromUser.Value}/messages/{originalMessageId.Value}/{replyType}";
            var createReplyResponse = client.PostAsync(createReplyUrl, null).Result;

            if (!createReplyResponse.IsSuccessStatusCode)
            {
                SqlContext.Pipe.Send("ERROR (CreateReply): " + createReplyResponse.Content.ReadAsStringAsync().Result);
                return;
            }

            var replyMessageJson = JObject.Parse(createReplyResponse.Content.ReadAsStringAsync().Result);
            string replyId = replyMessageJson["id"]?.ToString();

            // Step 2: Update body content
            var patchBody = new JObject
            {
                ["body"] = new JObject
                {
                    ["contentType"] = "Text",
                    ["content"] = replyBody.Value
                }
            };

            string updateUrl = $"https://graph.microsoft.com/v1.0/users/{fromUser.Value}/messages/{replyId}";
            var patchRequest = new HttpRequestMessage(new HttpMethod("PATCH"), updateUrl)
            {
                Content = new StringContent(patchBody.ToString(), System.Text.Encoding.UTF8, "application/json")
            };
            var patchResponse = client.SendAsync(patchRequest).Result;

            if (!patchResponse.IsSuccessStatusCode)
            {
                SqlContext.Pipe.Send("ERROR (PatchReply): " + patchResponse.Content.ReadAsStringAsync().Result);
                return;
            }

            // Step 3: Send reply
            string sendUrl = $"https://graph.microsoft.com/v1.0/users/{fromUser.Value}/messages/{replyId}/send";
            var sendResponse = client.PostAsync(sendUrl, null).Result;

            if (!sendResponse.IsSuccessStatusCode)
            {
                SqlContext.Pipe.Send("ERROR (SendReply): " + sendResponse.Content.ReadAsStringAsync().Result);
                return;
            }

            // Step 4: Return info
            replyMessageId = new SqlString(replyId);
            conversationId = new SqlString(replyMessageJson["conversationId"]?.ToString() ?? "");
            internetMessageId = new SqlString(replyMessageJson["internetMessageId"]?.ToString() ?? "");
            createdDateTime = new SqlString(replyMessageJson["createdDateTime"]?.ToString() ?? "");
            sentDateTime = new SqlString(DateTime.UtcNow.ToString("o"));

            SqlContext.Pipe.Send("Reply email sent successfully.");
        }
        catch (Exception ex)
        {
            SqlContext.Pipe.Send("EXCEPTION: " + ex.Message);
            if (ex.InnerException != null)
                SqlContext.Pipe.Send("INNER: " + ex.InnerException.Message);
        }
    }
}