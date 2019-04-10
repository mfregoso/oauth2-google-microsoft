namespace Services
{
    public class GoogleOAuthService
    {
        readonly IDataProvider dataProvider;
        public GoogleOAuthService(IDataProvider dataProvider)
        {
            this.dataProvider = dataProvider;
        }

        public readonly string GoogleApplicationName = "INSERT_HERE";
        public readonly string GoogleClientId = "INSERT_HERE";
        public readonly string GoogleClientSecret = "INSERT_HERE";

        public bool AuthorizationHandler(int UserId, string AuthCode)
        {
            bool obtainedToken = false;
            DateTime CurrentTime = DateTime.UtcNow;

            string ReqType = "&grant_type=authorization_code";
            string RedirectUri = "&redirect_uri=YOUR_URL_HERE";
            string ClientId = GoogleClientId;
            string ClientSecret = GoogleClientSecret;
            string GoogleOAuthApi = "https://www.googleapis.com/oauth2/v4/token?code=";

            string gapiRespObject;
            
            HttpWebRequest tokenRequest = (HttpWebRequest)WebRequest.Create(GoogleOAuthApi + AuthCode + ReqType + RedirectUri + "&client_id=" + ClientId + "&client_secret=" + ClientSecret);
            tokenRequest.Method = "POST";
            tokenRequest.ContentType = "application/x-www-form-urlencoded";
            tokenRequest.ContentLength = 0;
            try
            {
                using (HttpWebResponse webResp = (HttpWebResponse)tokenRequest.GetResponse())
                {
                    using (Stream stream = webResp.GetResponseStream())
                    {
                        StreamReader reader = new StreamReader(stream, System.Text.Encoding.UTF8);
                        gapiRespObject = reader.ReadToEnd();
                    }
                    var gapiRespString = (JObject)JsonConvert.DeserializeObject(gapiRespObject);
                    System.Diagnostics.Debug.WriteLine(gapiRespString);

                   
                    if (gapiRespString["refresh_token"] == null)
                    // Case: Already Have Permission, Unnecessary Authorization Request
                    // This Response ONLY Has Access Token :: So Do NOT try to read non-existent Refresh Token
                    {
                        string AccessToken = gapiRespString["access_token"].Value<string>();
                        string ExpiresIn = gapiRespString["expires_in"].Value<string>();

                        int ExpirationSeconds = Int16.Parse(ExpiresIn);
                        DateTime AccessExpiration = CurrentTime.AddSeconds(ExpirationSeconds);

                        StoreGoogleToken(UserId, 1, AccessToken, AccessExpiration);
                   
                    } else
                    // Case: Brand New Authorization With Authorization + Refresh Token
                    // Set User as Google Authorization + Send Tokens To SQL
                    {
                        string AccessToken = gapiRespString["access_token"].Value<string>();
                        string ExpiresIn = gapiRespString["expires_in"].Value<string>();
                        string RefreshToken = gapiRespString["refresh_token"].Value<string>();

                        int ExpirationSeconds = Int16.Parse(ExpiresIn);
                        DateTime AccessExpiration = CurrentTime.AddSeconds(ExpirationSeconds);

                        SetUserGoogleStatus(UserId, true);
                        StoreGoogleToken(UserId, 0, RefreshToken, null);
                        StoreGoogleToken(UserId, 1, AccessToken, AccessExpiration);
                    }
                    
                    obtainedToken = true;
                }
            } catch (WebException ex)
            {
                string resp;
                using (var streamReader = new StreamReader(ex.Response.GetResponseStream()))
                {
                    resp = streamReader.ReadToEnd();
                }                
                System.Diagnostics.Debug.WriteLine(resp);
                System.Diagnostics.Debug.WriteLine("Service Error: no tokens from POST");
            }
            return obtainedToken;
        }

        public string GetNewAccessToken(int UserId)
        {
            string AccessToken = null;
            DateTime CurrentTime = DateTime.UtcNow;
            string ReqType = "&grant_type=refresh_token";
            string ClientId = GoogleClientId;
            string ClientSecret = GoogleClientSecret;
            string GoogleOAuthApi = "https://www.googleapis.com/oauth2/v4/token?refresh_token=";
            string RefreshToken = GetGoogleRefreshToken(UserId);

            HttpWebRequest tokenRequest = (HttpWebRequest)WebRequest.Create(GoogleOAuthApi + RefreshToken + ReqType + "&client_id=" + ClientId + "&client_secret=" + ClientSecret);
            tokenRequest.Method = "POST";
            tokenRequest.ContentType = "application/x-www-form-urlencoded";
            tokenRequest.ContentLength = 0;
            string gapiRespObject;

            if (RefreshToken != null)
            {
                try
                {
                    using (HttpWebResponse webResp = (HttpWebResponse)tokenRequest.GetResponse())
                    {
                        using (Stream stream = webResp.GetResponseStream())
                        {
                            StreamReader reader = new StreamReader(stream, System.Text.Encoding.UTF8);
                            gapiRespObject = reader.ReadToEnd();
                        }
                        var gapiRespString = (JObject)JsonConvert.DeserializeObject(gapiRespObject);
                        System.Diagnostics.Debug.WriteLine(gapiRespString);

                        
                            AccessToken = gapiRespString["access_token"].Value<string>(); // set the "global" variable up above
                            string ExpiresIn = gapiRespString["expires_in"].Value<string>();

                            int ExpirationSeconds = Int16.Parse(ExpiresIn);
                            DateTime AccessExpiration = CurrentTime.AddSeconds(ExpirationSeconds);

                            StoreGoogleToken(UserId, 1, AccessToken, AccessExpiration); // STORE IN SQL
                    }
                }
                catch (WebException ex)
                // CASE: Invalid Refresh Token = App No Longer Authorized By User / Must Re-Authorize For Google Functionality
                {
                    string resp;
                    using (var streamReader = new StreamReader(ex.Response.GetResponseStream()))
                    {
                        resp = streamReader.ReadToEnd();
                    }                    
                    System.Diagnostics.Debug.WriteLine(resp);
                    System.Diagnostics.Debug.WriteLine("Refresh FAILED: Invalid Refresh Token / User Must Re-Authorize");
                    SetUserGoogleStatus(UserId, false); // user flagged as NOT googly
                }
            }

            return AccessToken;
        }
    }
}
