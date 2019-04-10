namespace Web.Controllers
{
    [RoutePrefix("api/calendar")]
    public class CalendarController : ApiController
    {
        readonly ICalendarService calendarService;
        readonly IAuthenticationService authenticationService;
        readonly ExternalCalendarHelperService exCalHelperService;
        readonly OAuthService oAuthService;
        readonly MicrosoftOAuthService msOAuthService;
        readonly MicrosoftCalendarService msCalService;
        readonly GoogleCalendarService googleCalendarService;

        public CalendarController(ICalendarService calendarService, IAuthenticationService authenticationService, ExternalCalendarHelperService exCalHelperService, OAuthService oAuthService, GoogleCalendarService googleCalendarService, MicrosoftOAuthService msOAuthService, MicrosoftCalendarService msCalService)
        {
            this.calendarService = calendarService;
            this.authenticationService = authenticationService;
            this.exCalHelperService = exCalHelperService;
            this.oAuthService = oAuthService;
            this.googleCalendarService = googleCalendarService;
            this.msOAuthService = msOAuthService;
            this.msCalService = msCalService;
        }

        [Route, HttpPost]
        public HttpResponseMessage Create(TimeBlockCreateRequest createRequest)
        {
            if (createRequest == null)
            {
                ModelState.AddModelError("", "No data supplied!");
            }

            if (!ModelState.IsValid)
            {
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ModelState);
            }

            int UserId = User.Identity.GetId().Value;
            createRequest.UserId = UserId;
            int newId = calendarService.CreateTimeBlock(createRequest);

            Task.Run(() =>
            {
                UserConnectedCalendars userCalendar = exCalHelperService.CheckConnectedCalendars(UserId);
                if (userCalendar.hasGoogle) { googleCalendarService.CreateEvent(UserId, newId, createRequest); }
                if (userCalendar.hasMicrosoft) { msCalService.CreateEvent(UserId, newId, createRequest); }
            });
            
            return Request.CreateResponse(HttpStatusCode.Created, new ItemResponse<int> { Item = newId });
        }

        [Route("{Id:int}"), HttpPut]
        public HttpResponseMessage Update(int Id, TimeBlockUpdateRequest updateRequest)
        {
            if (updateRequest == null)
            {
                ModelState.AddModelError("", "No data supplied!");
            }

            if (!ModelState.IsValid)
            {
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ModelState);
            }

            if (Id == updateRequest.Id)
            {
                calendarService.UpdateTimeBlock(updateRequest);

                int UserId = User.Identity.GetId().Value;

                Task.Run(() =>
                {
                    UserConnectedCalendars userCalendar = exCalHelperService.CheckConnectedCalendars(UserId);
                    if (userCalendar.hasGoogle) { googleCalendarService.UpdateEvent(UserId, Id, updateRequest); }
                    if (userCalendar.hasMicrosoft) { msCalService.UpdateEvent(UserId, Id, updateRequest); }
                });

                return Request.CreateResponse(HttpStatusCode.OK, new SuccessResponse());
            } else
            {
                ModelState.AddModelError("", "Id in URL does not match Id in request body");
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ModelState);
            }
        }

        [Route("{Id:int}"),HttpDelete]
        public HttpResponseMessage Delete(int Id)
        {
            int UserId = User.Identity.GetId().Value;
            calendarService.DeleteTimeBlock(Id, UserId);           

            Task.Run(() =>
            {
                UserConnectedCalendars userCalendar = exCalHelperService.CheckConnectedCalendars(UserId);
                if (userCalendar.hasGoogle) { googleCalendarService.DeleteEvent(UserId, Id); }
                if (userCalendar.hasMicrosoft) { msCalService.DeleteEvent(UserId, Id); }
            });

            return Request.CreateResponse(HttpStatusCode.OK, new SuccessResponse());
        }

        [Route("{Id:int}"), HttpGet]
        public HttpResponseMessage GetById(int Id)
        {
            int UserId = User.Identity.GetId().Value;
            ItemResponse < TimeBlock > itemResponse = calendarService.GetById(Id, UserId);
            return Request.CreateResponse(HttpStatusCode.OK, itemResponse);
        }
        
        [Route("user"), HttpGet]
        public HttpResponseMessage GetByUserDateRange(string Start, string End)
        {
            int UserId = User.Identity.GetId().Value;
            ItemsResponse<TimeBlock> itemsResponse = calendarService.GetByUserDateRange(UserId, Start, End);
            return Request.CreateResponse(HttpStatusCode.OK, itemsResponse);
        }    
    }
}