using Newtonsoft.Json;

namespace ES.SendMail.Models
{
    public class Mail
    {
        public Message message { get; set; }
        public bool saveToSentItems { get; set; }
    }

    public class Message
    {
        public string subject { get; set; }
        public Body body { get; set; }
        public Torecipient[] toRecipients { get; set; }
        public bool hasAttachments { get; set; }
        public Attachment[] attachments { get; set; }
    }

    public class Body
    {
        public string contentType { get; set; }
        public string content { get; set; }
    }

    public class Torecipient
    {
        public Emailaddress emailAddress { get; set; }
    }

    public class Emailaddress
    {
        public string name { get; set; }
        public string address { get; set; }
    }

    public class Attachment
    {
        [JsonProperty("@odata.type")]
        public string odatatype { get; set; }
        public string contentBytes { get; set; }
        public string contentId { get; set; }
        public string contentLocation { get; set; }
        public string contentType { get; set; }
        public string name { get; set; }
    }
}
