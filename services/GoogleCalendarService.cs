namespace Services
{
    public class GoogleCalendarService
    {
        readonly IDataProvider dataProvider;
        readonly OAuthService oAuthService;
        readonly ExternalCalendarHelperService exCalHelperService;
        public GoogleCalendarService(IDataProvider dataProvider, OAuthService oAuthService, ExternalCalendarHelperService exCalHelperService)
        {
            this.dataProvider = dataProvider;
            this.oAuthService = oAuthService;
            this.exCalHelperService = exCalHelperService;
        }        

        public async void BatchInsertFromTemplate(int UserId, string Timezone, List<WeeklyTemplateEventData> fromTemplateEvents)
        {
            string AccessToken = oAuthService.GetGoogleAccessToken(UserId);

            if (AccessToken != null && fromTemplateEvents.Count != 0)
            {
                System.Diagnostics.Debug.WriteLine("Calendar has FromTemplate events!... attemping Google batch insert!");

                string[] scopes = new string[] { "https://www.googleapis.com/auth/calendar" };
                var flow = new GoogleAuthorizationCodeFlow(new GoogleAuthorizationCodeFlow.Initializer
                {
                    ClientSecrets = new ClientSecrets
                    {
                        ClientId = oAuthService.GoogleClientId,
                        ClientSecret = oAuthService.GoogleClientSecret
                    },
                    Scopes = scopes,
                    DataStore = new FileDataStore("Store")
                });
                var token = new TokenResponse
                {
                    AccessToken = AccessToken,
                    RefreshToken = oAuthService.GetGoogleRefreshToken(UserId)
                };
                var credential = new UserCredential(flow, UserId.ToString(), token);

                var service = new CalendarService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = oAuthService.GoogleApplicationName,
                });

                var request = new BatchRequest(service);

                foreach (WeeklyTemplateEventData template in fromTemplateEvents)
                {
                    request.Queue<Event>(service.Events.Insert(
                new Event
                {
                    Summary = template.Title,
                    Start = new EventDateTime() { DateTimeRaw = template.StartDateTime, TimeZone = Timezone },
                    End = new EventDateTime() { DateTimeRaw = template.EndDateTime, TimeZone = Timezone },
                    ColorId = exCalHelperService.GetGoogleColor(template.CategoryId)
                }, "primary"),
                    (content, error, i, message) =>
                    {
                        exCalHelperService.StoreExternalEventId(UserId, template.LocalId, content.Id, 1, content.ETag);
                    });
                }
                await request.ExecuteAsync();
            }
        }

        public async void BatchDeleteFromTemplateEvents(int UserId, List<ExternalEventIds> externalEventIds)
        {
            string AccessToken = oAuthService.GetGoogleAccessToken(UserId);

            if (AccessToken != null && externalEventIds.Count != 0)
            {
                string[] scopes = new string[] { "https://www.googleapis.com/auth/calendar" };
                var flow = new GoogleAuthorizationCodeFlow(new GoogleAuthorizationCodeFlow.Initializer
                {
                    ClientSecrets = new ClientSecrets
                    {
                        ClientId = oAuthService.GoogleClientId,
                        ClientSecret = oAuthService.GoogleClientSecret
                    },
                    Scopes = scopes,
                    DataStore = new FileDataStore("Store")
                });
                var token = new TokenResponse
                {
                    AccessToken = AccessToken,
                    RefreshToken = oAuthService.GetGoogleRefreshToken(UserId)
                };
                var credential = new UserCredential(flow, UserId.ToString(), token);

                var service = new CalendarService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = oAuthService.GoogleApplicationName,
                });

                var request = new BatchRequest(service);

                foreach (ExternalEventIds template in externalEventIds)
                {
                    if (template.GoogleId != null)
                    {
                        request.Queue<Event>(service.Events.Delete("primary", template.GoogleId),
                        (content, error, i, message) =>
                        {
                            exCalHelperService.DeleteExternalEvent(template.GoogleId);
                        });
                    }
                }
                await request.ExecuteAsync();
            }
        }

        public List<GoogleAttendees> AttendeeList(List<string> validEmails)
        {
            List<GoogleAttendees> attendeeList = new List<GoogleAttendees>();
            if (validEmails.Count != 0)
            {
                foreach (string email in validEmails)
                {
                    GoogleAttendees emailObject = new GoogleAttendees
                    {
                        Email = email
                    };
                    attendeeList.Add(emailObject);
                }
            }
            return attendeeList;
        }

        public void CreateEvent(int UserId, int LocalId, TimeBlockCreateRequest NewEvent)
        {
            string CalendarApi = "https://www.googleapis.com/calendar/v3/calendars/primary/events?sendNotifications=true&access_token=";
            string AccessToken = oAuthService.GetGoogleAccessToken(UserId);
            string gapiRespObject;

            if (AccessToken != null)
            {

                string gStartTime = NewEvent.StartDate + "T" + NewEvent.StartTime + ":00";
                string gEndTime = NewEvent.EndDate + "T" + NewEvent.EndTime + ":00";
                HttpWebRequest createRequest = (HttpWebRequest)WebRequest.Create(CalendarApi + AccessToken);
                createRequest.Method = "POST";
                createRequest.ContentType = "application/json";

                // set bg color
                int categoryId = 2;
                if (NewEvent.ActivityCategory.HasValue) { categoryId = NewEvent.ActivityCategory.Value; }
                string bgColor = exCalHelperService.GetGoogleColor(categoryId);

                // timezone failsafe logic
                string userTimezone = "America/Los_Angeles";
                if (NewEvent.TimeZone != null) { userTimezone = NewEvent.TimeZone; }

                // placeholders for the eventual start & end fields
                object startField;
                object endField;
                // regular events
                var startDateTime = new { dateTime = gStartTime, timeZone = userTimezone };
                var endDateTime = new { dateTime = gEndTime, timeZone = userTimezone };
                // all day events
                var startDate = new { date = NewEvent.StartDate };
                var endDate = new { date = NewEvent.EndDate };
                // select one
                if (NewEvent.AllDay) { startField = startDate; endField = endDate; }
                else { startField = startDateTime; endField = endDateTime; }

                List<string> validEmails = exCalHelperService.ReturnValidEmails(NewEvent.GuestEmail);

                string json = "";
                if (validEmails.Count != 0)
                {
                    List<GoogleAttendees> attendeeEmails = AttendeeList(validEmails);

                    var tempObj = new
                    {
                        summary = NewEvent.Title,
                        location = NewEvent.Location,
                        description = NewEvent.Description,
                        start = startField,
                        end = endField,
                        colorId = bgColor,
                        reminders = new { useDefault = true },
                        attendees = attendeeEmails
                    };
                    json = JsonConvert.SerializeObject(tempObj);

                }
                else
                {
                    json = new JavaScriptSerializer().Serialize(new
                    {
                        summary = NewEvent.Title,
                        location = NewEvent.Location,
                        description = NewEvent.Description,
                        start = startField,
                        end = endField,
                        colorId = bgColor,
                        reminders = new { useDefault = true }
                    });
                }

                try
                {
                    using (var streamWriter = new StreamWriter(createRequest.GetRequestStream()))
                    {
                        streamWriter.Write(json);

                    }

                    var httpResponse = (HttpWebResponse)createRequest.GetResponse();
                    using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                    {
                        gapiRespObject = streamReader.ReadToEnd();
                    }

                    var GoogleResp = (JObject)JsonConvert.DeserializeObject(gapiRespObject);
                    string ExternalId = GoogleResp["id"].Value<string>();
                    string ETag = GoogleResp["etag"].Value<string>();
                    exCalHelperService.StoreExternalEventId(UserId, LocalId, ExternalId, 1, ETag); // 1 = Google
                }
                catch (WebException ex)
                {
                    string resp;
                    using (var streamReader = new StreamReader(ex.Response.GetResponseStream()))
                    {
                        resp = streamReader.ReadToEnd();
                    }
                    System.Diagnostics.Debug.WriteLine("Failed to POST ERROR");
                    System.Diagnostics.Debug.WriteLine(resp);
                }
            }
        }

        public void UpdateEvent(int UserId, int LocalId, TimeBlockUpdateRequest updated)
        {
            string EventId = exCalHelperService.GetExternalEventId(LocalId, 1); // provider 1 = google
            string CalendarApi = "https://www.googleapis.com/calendar/v3/calendars/primary/events/" + EventId + "?sendNotifications=true&access_token=";
            string AccessToken = null;
            if (EventId != null) { AccessToken = oAuthService.GetGoogleAccessToken(UserId); }
            string gapiRespObject;

            if (EventId != null && AccessToken != null)
            {

                string gStartTime = updated.StartDate + "T" + updated.StartTime + ":00";
                string gEndTime = updated.EndDate + "T" + updated.EndTime + ":00";
                HttpWebRequest updateRequest = (HttpWebRequest)WebRequest.Create(CalendarApi + AccessToken);
                updateRequest.Method = "PUT";
                updateRequest.ContentType = "application/json";

                // set bg color
                int categoryId = 2;
                //categoryId = updated.ActivityCategory.Value;
                if (updated.ActivityCategory.HasValue) { categoryId = updated.ActivityCategory.Value; }
                string bgColor = exCalHelperService.GetGoogleColor(categoryId);

                // timezone failsafe logic
                string userTimezone = "America/Los_Angeles";
                if (updated.TimeZone != null) { userTimezone = updated.TimeZone; }

                // for events marked as canceled
                string eventStatus = "confirmed";
                if (updated.Canceled) { eventStatus = "cancelled"; }

                // placeholders for the eventual start & end fields
                object startField;
                object endField;
                // regular events
                var startDateTime = new { dateTime = gStartTime, timeZone = userTimezone };
                var endDateTime = new { dateTime = gEndTime, timeZone = userTimezone };
                // all day events
                var startDate = new { date = updated.StartDate };
                var endDate = new { date = updated.EndDate };
                // select one
                if (updated.AllDay) { startField = startDate; endField = endDate; }
                else { startField = startDateTime; endField = endDateTime; }

                List<string> validEmails = exCalHelperService.ReturnValidEmails(updated.GuestEmail);

                string json = "";
                if (validEmails.Count != 0)
                {
                    List<GoogleAttendees> attendeeEmails = AttendeeList(validEmails);
                    var tempObj = new
                    {
                        summary = updated.Title,
                        location = updated.Location,
                        description = updated.Description,
                        start = startField,
                        end = endField,
                        colorId = bgColor,
                        reminders = new { useDefault = true },
                        attendees = attendeeEmails,
                        status = eventStatus
                    };
                    json = JsonConvert.SerializeObject(tempObj);
                }
                else
                {
                    json = new JavaScriptSerializer().Serialize(new
                    {
                        summary = updated.Title,
                        location = updated.Location,
                        description = updated.Description,
                        start = startField,
                        end = endField,
                        colorId = bgColor,
                        status = eventStatus
                    });
                }

                try
                {
                    using (var streamWriter = new StreamWriter(updateRequest.GetRequestStream()))
                    {
                        streamWriter.Write(json);

                    }

                    System.Diagnostics.Debug.WriteLine(json);
                    var httpResponse = (HttpWebResponse)updateRequest.GetResponse();
                    using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                    {
                        gapiRespObject = streamReader.ReadToEnd();
                    }

                    var GoogleResp = (JObject)JsonConvert.DeserializeObject(gapiRespObject);

                    string ETag = GoogleResp["etag"].Value<string>();
                    exCalHelperService.UpdateETag(EventId, ETag);
                }
                catch (WebException ex)
                {
                    string resp;
                    using (var streamReader = new StreamReader(ex.Response.GetResponseStream()))
                    {
                        resp = streamReader.ReadToEnd();
                    }
                    System.Diagnostics.Debug.WriteLine("Failed to UPDATE");
                    System.Diagnostics.Debug.WriteLine(resp);
                }
            }
        }

        public void DeleteEvent(int UserId, int LocalId)
        {
            string EventId = exCalHelperService.GetExternalEventId(LocalId, 1); // provider 1 = google
            string CalendarApi = "https://www.googleapis.com/calendar/v3/calendars/primary/events/" + EventId + "?access_token=";
            string AccessToken = null;
            if (EventId != null) { AccessToken = oAuthService.GetGoogleAccessToken(UserId); }
            string gapiRespObject;

            if (EventId != null && AccessToken != null)
            {
                HttpWebRequest deleteRequest = (HttpWebRequest)WebRequest.Create(CalendarApi + AccessToken);
                deleteRequest.Method = "DELETE";
                deleteRequest.ContentType = "application/x-www-form-urlencoded";
                deleteRequest.ContentLength = 0;

                try
                {
                    using (HttpWebResponse webResp = (HttpWebResponse)deleteRequest.GetResponse())
                    {
                        using (Stream stream = webResp.GetResponseStream())
                        {
                            StreamReader reader = new StreamReader(stream, System.Text.Encoding.UTF8);
                            gapiRespObject = reader.ReadToEnd();
                        }
                        var gapiRespString = (JObject)JsonConvert.DeserializeObject(gapiRespObject);
                        System.Diagnostics.Debug.WriteLine(gapiRespString);
                    }

                    exCalHelperService.DeleteExternalEvent(EventId);
                }
                catch (WebException ex)
                {
                    string resp;
                    using (var streamReader = new StreamReader(ex.Response.GetResponseStream()))
                    {
                        resp = streamReader.ReadToEnd();
                    }
                    System.Diagnostics.Debug.WriteLine("Failed to DELETE");
                    System.Diagnostics.Debug.WriteLine(resp);
                }
            }
        }        
    }
}
