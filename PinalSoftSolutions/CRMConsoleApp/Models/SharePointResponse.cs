using System.Runtime.Serialization;

namespace zCRMConsoleApp.Models
{
    [DataContract]
    public class SharePointResponse
    {
        [DataMember]
        public string token_type { get; set; }
        [DataMember]
        public string expires_in { get; set; }
        [DataMember]
        public string not_before { get; set; }
        [DataMember]
        public string expires_on { get; set; }
        [DataMember]
        public string resource { get; set; }
        [DataMember]
        public string access_token { get; set; }
    }
}
