using System;
using System.Configuration;
using ES.SendMail.Models;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using RestSharp;

namespace ES.SendMail
{
    class Program
    {
        private static readonly string Resourse = ConfigurationManager.AppSettings["resourse"];
        private static readonly string ClientId = ConfigurationManager.AppSettings["clientId"];
        private static readonly string RedirectUrl = ConfigurationManager.AppSettings["redirectUrl"];
        private static readonly string AuthorizationUri = ConfigurationManager.AppSettings["authorizationUri"];
        private static readonly string TenantId = ConfigurationManager.AppSettings["tenantId"];

        [STAThread]
        static void Main()
        {
            // Get an access token
            var authResult = GetAccessToken();

            // Get current user
            var user = GetCurrentUser(authResult.AccessToken);

            // Send an email
            SendMailToUser(authResult.AccessToken, user.mail);

            Console.Read();
        }

        private static AuthenticationResult GetAccessToken()
        {
            //  https://login.windows.net/<tenant-id>/oauth2/authorize
            var authority = AuthorizationUri.Replace("common", TenantId);
            var authenticationContext = new AuthenticationContext(authority);
            var authenticationResult = authenticationContext.AcquireToken(Resourse, ClientId, new Uri(RedirectUrl), PromptBehavior.Always);
            return authenticationResult;
        }

        private static User GetCurrentUser(string accessToken)
        {
            var client = new RestClient(Resourse);
            var request = new RestRequest("/v1.0/me", Method.GET);
            request.AddHeader("Authorization", "Bearer " + accessToken);
            request.AddHeader("Content-Type", "application/json");
            request.AddHeader("Accept", "application/json");

            var response = client.Execute(request);
            var content = response.Content;

            return JsonConvert.DeserializeObject<User>(content);
        }

        private static void SendMailToUser(string accessToken, string email)
        {
            var mailMarkup = @"<h1>A mail with an embedded image <img src='cid:thumbsUp' alt='Thumbs up' /> <img src='https://cdn2.iconfinder.com/data/icons/hawcons-gesture-stroke/32/icon_3_high_five-24.png' alt='High five' /></h1>";
            var contentBytes = "iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAADLUlEQVRIS7XVQWhcRRgH8P/3zW6zoQ3Rgl40FnWx0pps3nuz2wYRc1DMwYooVKj1IBrQXqSH4kEQPVgEvQgqLIhYKG3BU6kg2IILBhEy8/YlelCbmoIpVCpCgrqa3fk+edItS6HNardznfn+v3nvfW+G0OeI43hWVd9m5gDgLe/9e/2UUj+LqtXqWKfTWTbGNFS1BeBxAIe89+9uVN8v8IiInAbwkPd+zlp7QlWf7HQ69y4sLFy4HtIXYK19QVU/AlD23p+LomgbMy+r6qtpmr5zw0CSJEdFZKbZbN4OQPLAycnJc0T0TbPZfPaGgEqlcgczn2XmI977l7thURQtGmPOOuee/t/A9PR0aXV19XMANVXdmWXZ+W5YkiTLRDTnnHtuI4CttXtFZDszk6oGETljjAmq+iGAiIj2O+eO94SPAvgthPB1oVA4pap/A1hX1b+I6HcRWRkdHZ1vNBodSpLkAwAHrtqFAiARuUhEs2maftY7H8fxM0R04no7DyH8RER7KIqiP5n5mPf+xbygXC4PjYyMZMz8Sbvdfn9xcfGPjXodACdJYtrt9qbh4eEt6+vrDwCoq+qv+RMoEb3pnHujGxTHcSNN0+k+gq+5JEmSwyJy6KYAU1NTW1ut1kljzJ0DBay19bxZRGTMGHMPgNcHCnTf1+XvuMLMJ28KkCRJEcAlETk2MKBWq90tIttEhJl5v6o+z8yPDgyI43iGmXer6mYROWCMcc65hwcG9PartfZMp9PZkmXZ7oEDExMTm4vF4gUR+bTZbM5eAQAcV9U9pVKp3mq1Tv2XHy2Koh3GmL0AiiGEGQAVItqVpqm/AoiIBfAjEa2o6hNra2uPLS0t5YfYhmN8fPzWoaGhiojcR0R1IjrsnHstL+wFPDPnF/tBIvpYVceI6LSIfCUic71H9bVEa+12Vf2eiF5yztW7wA8hhIIx5ihRflprSUReYeafAdwG4JZ8oYj8QkTfqup5Zr5EROu9kIgUiOgpAHcxc3l+fv7iv4C1thZCOMLM918OCkT0hYjsy7JstVqt7gghPAhgl4jsJKL8Pt4KYNNVQGDm7wAc9N5/2Z37B3CL0ZrdDw2kAAAAAElFTkSuQmCC";
            var contentType = "image/png";

            // Create mail message
            var mail = new Mail
            {
                message = new Message
                {
                    toRecipients = new[]
                    {
                        new Torecipient
                        {
                            emailAddress = new Emailaddress
                            {
                                address = email,
                                name = email
                            }
                        }
                    },
                    body = new Body
                    {
                        content = mailMarkup,
                        contentType = "HTML"
                    },
                    subject = "Mail with an embedded image",
                    hasAttachments = true,
                    attachments = new[]
                    {
                        new Attachment
                        {
                            odatatype = "#microsoft.graph.fileAttachment",
                            contentBytes = contentBytes,
                            contentType = contentType,
                            contentId = "thumbsUp",
                            name = "thumbs-up.png"
                        }
                    }
                },
                saveToSentItems = true
            };

            // Send the mail
            var client = new RestClient(Resourse);
            var request = new RestRequest($"/v1.0/users/{email}/microsoft.graph.sendMail", Method.POST);
            request.AddHeader("Authorization", "Bearer " + accessToken);
            request.AddHeader("Accept", "application/json");
            // Remove null objects from the serialized JSON object
            var jsonBody = JsonConvert.SerializeObject(mail, Formatting.None, new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore
            });
            request.AddParameter("application/json", jsonBody, ParameterType.RequestBody);
            var response = client.Execute(request);
            Console.WriteLine("Statuscode: {0}", response.StatusCode);
        }
    }
}
