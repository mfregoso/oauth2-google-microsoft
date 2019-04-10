namespace Web.Controllers
{
    [RoutePrefix ("api/oauth")]
    public class OAuthController : ApiController
    {
        readonly GoogleOAuthService gOAuthService; 
        readonly MicrosoftOAuthService msOAuthService;
        readonly IAuthenticationService authenticationService;

        public OAuthController(GoogleOAuthService gOAuthService, IAuthenticationService authenticationService, MicrosoftOAuthService msOAuthService)
        {
            this.gOAuthService = gOAuthService;
            this.authenticationService = authenticationService;
            this.msOAuthService = msOAuthService;
        }

        [Route ("google"), HttpGet]
        public IHttpActionResult Log(string error="", string code="", string scope="", string state="")
        {
            bool obtainedAccessToken = false;
            int UserId = User.Identity.GetId().Value;
            string redirectUrl = "/#/app/calendar/connect?provider=Google&success=false";

            if (code != "")
            {
                obtainedAccessToken = gOAuthService.AuthorizationHandler(UserId, code);                
            }
            if (obtainedAccessToken)
            {
                redirectUrl = "/#/app/calendar/connect?provider=Google&success=true";
            }

            return Redirect(new Uri(redirectUrl, UriKind.Relative));
        }

        [Route("microsoft"), HttpGet]
        public IHttpActionResult Website(string error = "", string code = "", string scope = "", string state = "")
        {
            bool obtainedAccessToken = false;
            int UserId = User.Identity.GetId().Value;
            string redirectUrl = "/#/app/calendar/connect?provider=Microsoft&success=false";

            if (code != "")
            {
                obtainedAccessToken = msOAuthService.AuthorizationHandler(UserId, code);                
            }
            if (obtainedAccessToken)
            {
                redirectUrl = "/#/app/calendar/connect?provider=Microsoft&success=true";
            }

            return Redirect(new Uri(redirectUrl, UriKind.Relative));
        }
    }
}
