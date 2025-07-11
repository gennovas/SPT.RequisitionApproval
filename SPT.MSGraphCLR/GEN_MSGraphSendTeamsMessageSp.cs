using System;
using System.Data.SqlTypes;
using Microsoft.SqlServer.Server;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json.Linq;
using System.Net;

public partial class StoredProcedures
{
    [SqlProcedure]
    public static void GEN_MSGraphSendTeamsMessageSp(SqlString teamId, SqlString channelId, SqlString messageText)
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
            var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            string url = $"https://graph.microsoft.com/v1.0/teams/{teamId}/channels/{channelId}/messages";

            var messageBody = new JObject
            {
                ["body"] = new JObject
                {
                    ["content"] = messageText.Value
                }
            };

            var content = new StringContent(messageBody.ToString());
            content.Headers.ContentType = new MediaTypeHeaderValue("application/json");

            var response = client.PostAsync(url, content).Result;
            var resultContent = response.Content.ReadAsStringAsync().Result;

            SqlContext.Pipe.Send($"Status: {response.StatusCode}");
            SqlContext.Pipe.Send("Response: " + resultContent);
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