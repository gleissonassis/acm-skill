using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace AlexaSkill.Controllers
{
    public class AlexaController : ApiController
    {
        [HttpPost, Route("api/alexa/process")]
        public dynamic Process(dynamic request)
        {
            switch(request.request.intent.name.ToString())
            {
                case "ScheduleIntent":
                    return ProcessScheduleIntent(request);
                case "CheckIntent":
                    return ProcessCheckIntent(request);
                default:
                    return GenerateTextResponse("You can schedule a meeting on Office 365");
            }
        }

        private dynamic ProcessCheckIntent(dynamic request)
        {
            string dateCheck = Convert.ToString(request.request.intent.slots.when.value);
            string timeCheck = Convert.ToString(request.request.intent.slots.time.value);
            string durationCheck = Convert.ToString(request.request.intent.slots.duration.value);

            if (string.IsNullOrEmpty(dateCheck) && string.IsNullOrEmpty(timeCheck) && string.IsNullOrEmpty(durationCheck))
            {
                return GenerateTextResponse("Sorry! I didn't understand when and what time you're trying to check. Could you repeat, please?");
            }
            else
            {
                if (string.IsNullOrEmpty(dateCheck))
                {
                    dateCheck = DateTime.Now.ToString("yyyy-MM-dd");
                }
                if (string.IsNullOrEmpty(timeCheck))
                {
                    timeCheck = DateTime.Now.ToString("HH");
                }
                if (string.IsNullOrEmpty(durationCheck))
                {
                    durationCheck = "PT1H";
                }

                if (timeCheck.Length <= 2)
                {
                    timeCheck += ":00";
                }

                DateTime parsedDateCheck = Convert.ToDateTime(string.Concat(dateCheck, " ", timeCheck));
                TimeSpan parsedDurationCheck = System.Xml.XmlConvert.ToTimeSpan(durationCheck);
                var parsedEndTimeCheck = parsedDateCheck.Add(parsedDurationCheck);

                

                try
                {
                    HttpWebResponse httpResponse = VerifyMeeting(parsedDateCheck, parsedEndTimeCheck);

                    switch (httpResponse.StatusCode)
                    {
                        case HttpStatusCode.OK:
                            return GenerateTextResponse("Great! The room is free for use!");
                        default:
                            return GenerateTextResponse("Sorry! The service is not availabe now. Try againg later!");
                    }
                }
                catch (WebException ex)
                {
                    if (ex.Message.Equals("The remote server returned an error: (409) Conflict."))
                    {
                        return GenerateTextResponse("There is a reservation for this room in the request time. Do you want to try another time?");
                    }
                    else
                    {
                        return GenerateTextResponse("Sorry! The service is not availabe now. Try againg later!");
                    }
                }
            }
        }

        private HttpWebResponse VerifyMeeting(DateTime parsedDateCheck, DateTime parsedEndTimeCheck)
        {
            return PostJson("https://services.sysmap.com.br/acm/v1/meetings/verify", JsonConvert.SerializeObject(new
            {
                startDate = parsedDateCheck.ToString("o"),
                endDate = parsedEndTimeCheck.ToString("o"),
            }));
        }

        private HttpWebResponse SchedulingMeeting(DateTime parsedDate, DateTime parsedEndTime, string responsible)
        {
            return PostJson("https://services.sysmap.com.br/acm/v1/meetings", JsonConvert.SerializeObject(new
            {
                startDate = parsedDate.ToString("o"),
                endDate = parsedEndTime.ToString("o"),
                responsible = responsible
            }));
        }

        private static HttpWebResponse PostJson(string url, string json)
        {
            var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = "POST";

            using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
            {
                streamWriter.Write(json);
            }
            return (HttpWebResponse)httpWebRequest.GetResponse();
        }

        private dynamic GenerateTextResponse(string response)
        {
            return new
            {
                version = "1.0",
                sessionAttributes = new { },
                response = new
                {
                    outputSpeech = new
                    {
                        type = "PlainText",
                        text = response
                    },
                    shouldEndSession = true
                }
            };
        }

        private dynamic ProcessScheduleIntent(dynamic request)
        {
            string date = Convert.ToString(request.request.intent.slots.when.value);
            string time = Convert.ToString(request.request.intent.slots.time.value);
            string duration = Convert.ToString(request.request.intent.slots.duration.value);
            string responsible = Convert.ToString(request.request.intent.slots.responsible.value);

            if (string.IsNullOrEmpty(date) && string.IsNullOrEmpty(time) && string.IsNullOrEmpty(responsible))
            {
                return GenerateTextResponse("Sorry! I didn't understand what you're saying. Could you repeat, please?");
            }
            else if (string.IsNullOrEmpty(responsible))
            {
                return GenerateTextResponse("Sorry! I couldn't understant who is the responsible for the meeting. Could you repeat your request, please?");
            }
            else
            {
                if (string.IsNullOrEmpty(date))
                {
                    date = DateTime.Now.ToString("yyyy-MM-dd");
                }
                if (string.IsNullOrEmpty(time))
                {
                    time = DateTime.Now.ToString("HH:mm");
                }
                if (string.IsNullOrEmpty(duration))
                {
                    duration = "PT1H";
                }

                if (time.Length <= 2)
                {
                    time += ":00";
                }

                DateTime parsedDate = Convert.ToDateTime(string.Concat(date, " ", time));
                TimeSpan parsedDuration = System.Xml.XmlConvert.ToTimeSpan(duration);
                var parsedEndTime = parsedDate.Add(parsedDuration);

                try
                {
                    var httpResponse = SchedulingMeeting(parsedDate, parsedEndTime, responsible);

                    switch (httpResponse.StatusCode)
                    {
                        case HttpStatusCode.OK:
                        case HttpStatusCode.Created:
                            return GenerateTextResponse("Great! Your meeting is scheduled. Don't forget it!");

                        default:
                            return GenerateTextResponse("Sorry! The service is not availabe now. Try againg later!");
                    }
                }
                catch (WebException ex)
                {
                    if (ex.Message.Equals("The remote server returned an error: (409) Conflict."))
                    {
                        return GenerateTextResponse("There is a reservation for this room in the request time. Do you want to try another time?");
                    }
                    else
                    {
                        return GenerateTextResponse("Sorry! The service is not availabe now. Try againg later!");
                    }
                }
            }
        }
    }
}
