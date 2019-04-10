namespace Services
{
    public class MicrosoftCalendarService
    {
        readonly IDataProvider dataProvider;
        readonly MicrosoftOAuthService oAuthService;
        readonly ExternalCalendarHelperService exCalHelperService;
        public MicrosoftCalendarService(IDataProvider dataProvider, MicrosoftOAuthService oAuthService, ExternalCalendarHelperService exCalHelperService)
        {
            this.dataProvider = dataProvider;
            this.oAuthService = oAuthService;
            this.exCalHelperService = exCalHelperService;
        }      

        public void BatchInsertFromTemplate(int UserId, string Timezone, List<WeeklyTemplateEventData> fromTemplateEvents)
        {
            string CalendarApi = "https://graph.microsoft.com/v1.0/$batch";
            string AccessToken = oAuthService.GetMicrosoftAccessToken(UserId);
            string respObject;

            if (AccessToken != null && fromTemplateEvents.Count != 0)
            {
                HttpWebRequest createRequest = (HttpWebRequest)WebRequest.Create(CalendarApi);
                createRequest.Method = "POST";
                createRequest.ContentType = "application/json";
                createRequest.Headers["authorization"] = AccessToken;

                List<MsBatchRequest.PostContainer> BatchPostCollection = new List<MsBatchRequest.PostContainer>();
                foreach (WeeklyTemplateEventData template in fromTemplateEvents)
                {
                    MsBatchRequest.PostContainer batchEntry = new MsBatchRequest.PostContainer
                    {
                        Id = template.LocalId.ToString(),
                        Url = "/me/calendar/events",
                        Method = "POST",
                        Headers = new MsBatchRequest.MsHeaders
                        {
                            Content_Type = "application/json"
                        },
                        Body = new MsBatchRequest.PostBody
                        {
                            Subject = template.Title,
                            Start = new CalendarDateTimeZone
                            {
                                DateTime = template.StartDateTime,
                                TimeZone = Timezone
                            },
                            End = new CalendarDateTimeZone
                            {
                                DateTime = template.EndDateTime,
                                TimeZone = Timezone
                            },
                            Categories = new[] { exCalHelperService.GetMicrosoftColor(template.CategoryId) }
                        }
                };

                    BatchPostCollection.Add(batchEntry);
                }

                MsBatchRequest.PostRequestsContainer BatchPostRequests = new MsBatchRequest.PostRequestsContainer
                {
                    Requests = BatchPostCollection
                };
                                
                var json = JsonConvert.SerializeObject(BatchPostRequests);

                try
                {
                    using (var streamWriter = new StreamWriter(createRequest.GetRequestStream()))
                    {
                        streamWriter.Write(json);

                    }
                    
                    var httpResponse = (HttpWebResponse)createRequest.GetResponse();
                    using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                    {
                        respObject = streamReader.ReadToEnd();
                    }

                    var batchResponse = MsBatchResponse.FromJson(respObject);
                    foreach (Response postResp in batchResponse.Responses)
                    {
                        exCalHelperService.StoreExternalEventId(UserId, Convert.ToInt32(postResp.Id), postResp.Body.Id, 2, postResp.Body.OdataEtag);                        
                    }
                }
                catch (WebException ex)
                {
                    string resp;
                    using (var streamReader = new StreamReader(ex.Response.GetResponseStream()))
                    {
                        resp = streamReader.ReadToEnd();
                    }
                    System.Diagnostics.Debug.WriteLine("Failed to BATCH INSERT/POST");
                    System.Diagnostics.Debug.WriteLine(resp);
                }
            }            
        }

        public void BatchDeleteFromTemplateEvents(int UserId, List<ExternalEventIds> externalEventIds)
        {
            string CalendarApi = "https://graph.microsoft.com/v1.0/$batch";
            string AccessToken = oAuthService.GetMicrosoftAccessToken(UserId);
            string respObject;

            if (AccessToken != null && externalEventIds.Count != 0)
            {
                HttpWebRequest createRequest = (HttpWebRequest)WebRequest.Create(CalendarApi);
                createRequest.Method = "POST";
                createRequest.ContentType = "application/json";
                createRequest.Headers["authorization"] = AccessToken;

                List<MsBatchRequest.DeleteEventContainer> BatchDeleteCollection = new List<MsBatchRequest.DeleteEventContainer>();
                foreach (ExternalEventIds template in externalEventIds)
                {
                    if (template.MicrosoftId != null)
                    {
                        MsBatchRequest.DeleteEventContainer batchEntry = new MsBatchRequest.DeleteEventContainer
                        {
                            Id = template.LocalId.ToString(),
                            Url = "/me/calendar/events/" + template.MicrosoftId,
                            Method = "DELETE",
                            Headers = new MsBatchRequest.MsDeleteHeaders
                            {
                                Content_Type = "application/x-www-form-urlencoded",
                                Content_Length = "0"
                            }
                        };

                        BatchDeleteCollection.Add(batchEntry);
                        exCalHelperService.DeleteExternalEvent(template.MicrosoftId);
                    }
                }

                MsBatchRequest.DeleteRequestsContainer BatchPostRequests = new MsBatchRequest.DeleteRequestsContainer
                {
                    Requests = BatchDeleteCollection
                };

                var json = JsonConvert.SerializeObject(BatchPostRequests);

                try
                {
                    using (var streamWriter = new StreamWriter(createRequest.GetRequestStream()))
                    {
                        streamWriter.Write(json);

                    }

                    var httpResponse = (HttpWebResponse)createRequest.GetResponse();
                    using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                    {
                        respObject = streamReader.ReadToEnd();
                    }

                    var serverResp = (JObject)JsonConvert.DeserializeObject(respObject);
                    System.Diagnostics.Debug.WriteLine(serverResp);
                }
                catch (WebException ex)
                {
                    string resp;
                    using (var streamReader = new StreamReader(ex.Response.GetResponseStream()))
                    {
                        resp = streamReader.ReadToEnd();
                    }
                    System.Diagnostics.Debug.WriteLine("Failed to BATCH DELETE");
                    System.Diagnostics.Debug.WriteLine(resp);
                }
            }
        }

        public List<MicrosoftAttendees.Container> AttendeeList(List<string> validEmails)
        {
            List<MicrosoftAttendees.Container> attendeeList = new List<MicrosoftAttendees.Container>();
            if (validEmails.Count != 0)
            {
                foreach (string email in validEmails)
                {
                    MicrosoftAttendees.Container emailObject = new MicrosoftAttendees.Container
                    {
                        Type = "required",
                        EmailAddress = new MicrosoftAttendees.Email
                        {
                            Address = email
                        }
                    };
                    attendeeList.Add(emailObject);
                }
            }
            return attendeeList;
        }

        public void CreateEvent(int UserId, int LocalId, TimeBlockCreateRequest NewEvent)
        {
            string CalendarApi = "https://graph.microsoft.com/v1.0/me/calendar/events";
            string AccessToken = oAuthService.GetMicrosoftAccessToken(UserId);
            string respObject;

            if (AccessToken != null)
            {                                
                HttpWebRequest createRequest = (HttpWebRequest)WebRequest.Create(CalendarApi);
                createRequest.Method = "POST";
                createRequest.ContentType = "application/json";
                createRequest.Headers["authorization"] = AccessToken;

                string startTime = NewEvent.StartDate + "T" + NewEvent.StartTime + ":00";
                string endTime = NewEvent.EndDate + "T" + NewEvent.EndTime + ":00";
                // if all day, times need to be midnight
                if (NewEvent.AllDay)
                {
                    startTime = NewEvent.StartDate + "T00:00:00";
                    if (NewEvent.StartDate == NewEvent.EndDate)
                    { // a single all day event needs to be >24 hours otherwise MS will reject
                        string nextDay = DateTime.Parse(NewEvent.EndDate).AddDays(1).ToString("yyyy-MM-dd");
                        endTime = nextDay + "T00:00:00";
                    }
                    else { endTime = NewEvent.EndDate + "T00:00:00"; }

                }
                // grab activity category and turn it into a color
                int categoryId = 2; // failsafe = green
                if (NewEvent.ActivityCategory.HasValue) { categoryId = NewEvent.ActivityCategory.Value; }
                string[] activityColor = new[] { exCalHelperService.GetMicrosoftColor(categoryId) };

                // timezone failsafe logic
                string userTimezone = "America/Los_Angeles";
                if (NewEvent.TimeZone != null ) { userTimezone = NewEvent.TimeZone; }

                List<string> validEmails = exCalHelperService.ReturnValidEmails(NewEvent.GuestEmail);

                string json = "";
                if (validEmails.Count != 0)
                {
                    List<MicrosoftAttendees.Container> attendeeEmails = AttendeeList(validEmails);                    
                    var tempObj = new
                    {
                        subject = NewEvent.Title,
                        body = new { contentType = "Text", content = NewEvent.Description },
                        location = new { displayName = NewEvent.Location },
                        start = new { dateTime = startTime, timeZone = userTimezone },
                        end = new { dateTime = endTime, timeZone = userTimezone },
                        isAllDay = NewEvent.AllDay,
                        categories = activityColor,
                        attendees = attendeeEmails
                    };
                    json = JsonConvert.SerializeObject(tempObj);
                } else
                {
                    json = new JavaScriptSerializer().Serialize(new
                    {
                        subject = NewEvent.Title,
                        body = new { contentType = "Text", content = NewEvent.Description },
                        location = new { displayName = NewEvent.Location },
                        start = new { dateTime = startTime, timeZone = userTimezone },
                        end = new { dateTime = endTime, timeZone = userTimezone },
                        isAllDay = NewEvent.AllDay,
                        categories = activityColor
                    });
                }

                try
                {
                    using (var streamWriter = new StreamWriter(createRequest.GetRequestStream()))
                    {
                        streamWriter.Write(json);

                    }

                    System.Diagnostics.Debug.WriteLine(json);
                    var httpResponse = (HttpWebResponse)createRequest.GetResponse();
                    using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                    {
                        respObject = streamReader.ReadToEnd();
                    }

                    var serverResp = (JObject)JsonConvert.DeserializeObject(respObject);

                    string ExternalId = serverResp["id"].Value<string>();
                    string ETag = serverResp["@odata.etag"].Value<string>();
                    exCalHelperService.StoreExternalEventId(UserId, LocalId, ExternalId, 2, ETag); // 2 = microsoft
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
            string EventId = exCalHelperService.GetExternalEventId(LocalId, 2); // provider 2 = microsoft
            string CalendarApi = "https://graph.microsoft.com/v1.0/me/calendar/events/" + EventId;
            string AccessToken = null;
            if (EventId != null) { AccessToken = oAuthService.GetMicrosoftAccessToken(UserId); }
            string respObject;

            if (EventId != null && AccessToken != null)
            {                               
                HttpWebRequest updateRequest = (HttpWebRequest)WebRequest.Create(CalendarApi);
                updateRequest.Method = "PATCH";
                updateRequest.ContentType = "application/json";
                updateRequest.Headers["authorization"] = AccessToken;

                string startTime = updated.StartDate + "T" + updated.StartTime + ":00";
                string endTime = updated.EndDate + "T" + updated.EndTime + ":00";
                // if all day, times need to be set at midnight for Microsoft
                if (updated.AllDay) {
                    startTime = updated.StartDate + "T00:00:00";                    
                    if (updated.StartDate == updated.EndDate)
                    { // a single all day event needs to be >24 hours otherwise MS will reject
                        string nextDay = DateTime.Parse(updated.EndDate).AddDays(1).ToString("yyyy-MM-dd");
                        endTime = nextDay + "T00:00:00";
                    } else { endTime = updated.EndDate + "T00:00:00"; }

                }

                // grab activity category and turn it into a color
                int categoryId = 2; // failsafe = green
                if (updated.ActivityCategory.HasValue) { categoryId = updated.ActivityCategory.Value; }
                string[] activityColor = new[] { exCalHelperService.GetMicrosoftColor(categoryId) };

                // timezone failsafe logic
                string userTimezone = "America/Los_Angeles";
                if (updated.TimeZone != null) { userTimezone = updated.TimeZone; }

                List<string> validEmails = exCalHelperService.ReturnValidEmails(updated.GuestEmail);

                if (updated.Canceled && validEmails.Count != 0) // no need to trigger cancel API for guest-less events
                {
                    CancelEvent(EventId, AccessToken);
                }
                else
                {
                    string json = "";
                    if (validEmails.Count != 0)
                    {
                        List<MicrosoftAttendees.Container> attendeeEmails = AttendeeList(validEmails);
                        var tempObj = new
                        {
                            subject = updated.Title,
                            body = new { contentType = "Text", content = updated.Description },
                            location = new { displayName = updated.Location },
                            start = new { dateTime = startTime, timeZone = userTimezone },
                            end = new { dateTime = endTime, timeZone = userTimezone },
                            isAllDay = updated.AllDay,
                            categories = activityColor,
                            attendees = attendeeEmails
                        };
                        json = JsonConvert.SerializeObject(tempObj);
                    }
                    else
                    {
                        json = new JavaScriptSerializer().Serialize(new
                        {
                            subject = updated.Title,
                            body = new { contentType = "Text", content = updated.Description },
                            location = new { displayName = updated.Location },
                            start = new { dateTime = startTime, timeZone = userTimezone },
                            end = new { dateTime = endTime, timeZone = userTimezone },
                            isAllDay = updated.AllDay,
                            categories = activityColor
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
                            respObject = streamReader.ReadToEnd();
                        }

                        var serverResp = (JObject)JsonConvert.DeserializeObject(respObject);

                        string ETag = serverResp["@odata.etag"].Value<string>();
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
        }

        public void DeleteEvent(int UserId, int LocalId)
        {
            string EventId = exCalHelperService.GetExternalEventId(LocalId, 2); // provider 2 = microsoft
            string CalendarApi = "https://graph.microsoft.com/v1.0/me/calendar/events/" + EventId;
            string AccessToken = null;
            if (EventId != null) { AccessToken = oAuthService.GetMicrosoftAccessToken(UserId); }
            string respObject;

            if (EventId != null && AccessToken != null)
            {
                HttpWebRequest deleteRequest = (HttpWebRequest)WebRequest.Create(CalendarApi);
                deleteRequest.Method = "DELETE";
                deleteRequest.ContentType = "application/x-www-form-urlencoded";
                deleteRequest.ContentLength = 0;
                deleteRequest.Headers["authorization"] = AccessToken;

                try
                {
                    using (HttpWebResponse webResp = (HttpWebResponse)deleteRequest.GetResponse())
                    {
                        using (Stream stream = webResp.GetResponseStream())
                        {
                            StreamReader reader = new StreamReader(stream, System.Text.Encoding.UTF8);
                            respObject = reader.ReadToEnd();
                        }
                        var serverResp = (JObject)JsonConvert.DeserializeObject(respObject);
                        System.Diagnostics.Debug.WriteLine(serverResp);
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

        public void CancelEvent(string EventId, string AccessToken)
        {
            string CalendarApi = "https://graph.microsoft.com/beta/me/events/" + EventId + "/cancel";

            string respObject;

            HttpWebRequest cancelRequest = (HttpWebRequest)WebRequest.Create(CalendarApi);
            cancelRequest.Method = "POST";
            cancelRequest.ContentType = "application/x-www-form-urlencoded";
            cancelRequest.ContentLength = 0;
            cancelRequest.Headers["authorization"] = AccessToken;

            try
            {
                using (HttpWebResponse webResp = (HttpWebResponse)cancelRequest.GetResponse())
                {
                    using (Stream stream = webResp.GetResponseStream())
                    {
                        StreamReader reader = new StreamReader(stream, System.Text.Encoding.UTF8);
                        respObject = reader.ReadToEnd();
                    }
                    var serverResp = (JObject)JsonConvert.DeserializeObject(respObject);
                    System.Diagnostics.Debug.WriteLine(serverResp);
                    exCalHelperService.DeleteExternalEvent(EventId);
                }                
            }
            catch (WebException ex)
            {
                string resp;
                using (var streamReader = new StreamReader(ex.Response.GetResponseStream()))
                {
                    resp = streamReader.ReadToEnd();
                }
                System.Diagnostics.Debug.WriteLine("Failed Microsoft's beta cancellation");
                System.Diagnostics.Debug.WriteLine(resp);
            }
        }
    }
}
