namespace Services
{
    public class MicrosoftOAuthService
    {
        readonly IDataProvider dataProvider;
        public MicrosoftOAuthService(IDataProvider dataProvider)
        {
            this.dataProvider = dataProvider;
        }

        private readonly string MsClientId = "INSERT_HERE";
        private readonly string MsClientSecret = "INSERT_HERE";        

        public bool AuthorizationHandler(int UserId, string AuthCode)
        {
            bool obtainedToken = false;
            DateTime CurrentTime = DateTime.UtcNow;

            string ReqType = "&grant_type=authorization_code";
            string RedirectUri = "&redirect_uri=YOUR_URL_HERE";
            string Scope = "&scope=https://graph.microsoft.com/calendars.readwrite";
            string ClientId = MsClientId;
            string ClientSecret = MsClientSecret;
            string MSOAuthApi = "https://login.microsoftonline.com/common/oauth2/v2.0/token";

            string postBody = "code=" + AuthCode + ReqType + Scope + RedirectUri + "&client_id=" + ClientId + "&client_secret=" + ClientSecret;
            string responseObject;

            HttpWebRequest tokenRequest = (HttpWebRequest)WebRequest.Create(MSOAuthApi);
            tokenRequest.Method = "POST";
            tokenRequest.ContentType = "application/x-www-form-urlencoded";
            try
            {
                using (var streamWriter = new StreamWriter(tokenRequest.GetRequestStream()))
                {
                    streamWriter.Write(postBody);

                }
                using (HttpWebResponse webResp = (HttpWebResponse)tokenRequest.GetResponse())
                {
                    using (Stream stream = webResp.GetResponseStream())
                    {
                        StreamReader reader = new StreamReader(stream, System.Text.Encoding.UTF8);
                        responseObject = reader.ReadToEnd();
                    }
                    var respString = (JObject)JsonConvert.DeserializeObject(responseObject);
                    System.Diagnostics.Debug.WriteLine(respString);


                    if (respString["refresh_token"] == null)
                    // Case: App Already Has Permission
                    {
                        string AccessToken = respString["access_token"].Value<string>();
                        string ExpiresIn = respString["expires_in"].Value<string>();

                        int ExpirationSeconds = Int16.Parse(ExpiresIn);
                        DateTime AccessExpiration = CurrentTime.AddSeconds(ExpirationSeconds);

                        StoreMicrosoftToken(UserId, 1, AccessToken, AccessExpiration);

                    }
                    else
                    // Case: Brand New Authorization With Authorization + Refresh Token
                    // Set User as Microsoft Authorized + Send Tokens To SQL
                    {
                        string AccessToken = respString["access_token"].Value<string>();
                        string ExpiresIn = respString["expires_in"].Value<string>();
                        string RefreshToken = respString["refresh_token"].Value<string>();

                        int ExpirationSeconds = Int16.Parse(ExpiresIn);
                        DateTime AccessExpiration = CurrentTime.AddSeconds(ExpirationSeconds);

                        SetUserMicrosoftStatus(UserId, true);
                        StoreMicrosoftToken(UserId, 0, RefreshToken, null);
                        StoreMicrosoftToken(UserId, 1, AccessToken, AccessExpiration);
                    }
                    obtainedToken = true;
                }
            }
            catch (WebException ex)
            {
                string resp;
                using (var streamReader = new StreamReader(ex.Response.GetResponseStream()))
                {
                    resp = streamReader.ReadToEnd();
                }
                System.Diagnostics.Debug.WriteLine(resp);                
            }
            return obtainedToken;
        }        

        public string GetNewAccessToken(int UserId)
        {
            string AccessToken = null;
            DateTime CurrentTime = DateTime.UtcNow;
            string ReqType = "&grant_type=refresh_token";
            string RedirectUri = "&redirect_uri=YOUR_URL_HERE";
            string Scope = "&scope=https://graph.microsoft.com/calendars.readwrite";
            string ClientId = MsClientId;
            string ClientSecret = MsClientSecret;
            string MSOAuthApi = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
            string RefreshToken = GetMicrosoftRefreshToken(UserId);

            string postBody = "refresh_token=" + RefreshToken + ReqType + Scope + RedirectUri + "&client_id=" + ClientId + "&client_secret=" + ClientSecret;

            HttpWebRequest tokenRequest = (HttpWebRequest)WebRequest.Create(MSOAuthApi);
            tokenRequest.Method = "POST";
            tokenRequest.ContentType = "application/x-www-form-urlencoded";
            string respObject;

            if (RefreshToken != null)
            {
                try
                {
                    using (var streamWriter = new StreamWriter(tokenRequest.GetRequestStream()))
                    {
                        streamWriter.Write(postBody);

                    }
                    using (HttpWebResponse webResp = (HttpWebResponse)tokenRequest.GetResponse())
                    {
                        using (Stream stream = webResp.GetResponseStream())
                        {
                            StreamReader reader = new StreamReader(stream, System.Text.Encoding.UTF8);
                            respObject = reader.ReadToEnd();
                        }
                        var respString = (JObject)JsonConvert.DeserializeObject(respObject);
                        System.Diagnostics.Debug.WriteLine(respString);


                        AccessToken = respString["access_token"].Value<string>();
                        string ExpiresIn = respString["expires_in"].Value<string>();

                        int ExpirationSeconds = Int16.Parse(ExpiresIn);
                        DateTime AccessExpiration = CurrentTime.AddSeconds(ExpirationSeconds);

                        StoreMicrosoftToken(UserId, 1, AccessToken, AccessExpiration);
                    }
                }
                catch (WebException ex)
                {
                    string resp;
                    using (var streamReader = new StreamReader(ex.Response.GetResponseStream()))
                    {
                        resp = streamReader.ReadToEnd();
                    }
                    System.Diagnostics.Debug.WriteLine(resp);
                    SetUserMicrosoftStatus(UserId, false);
                }
            }
            return AccessToken;
        }
    }
}
