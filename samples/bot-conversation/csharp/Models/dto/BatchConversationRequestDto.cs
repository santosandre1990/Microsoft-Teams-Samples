using Microsoft.Bot.Schema;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace Microsoft.BotBuilderSamples.Models.dto
{
    public class BatchConversationRequestDto
    {
        [JsonProperty("Members")]
        public IEnumerable<Entry> Members { get; set; }
        [JsonProperty("Activity")]
        public JToken Activity { get; set; }
        [JsonProperty("TenantId")]
        public string TenantId { get; set; }
        [JsonProperty("TeamId")]
        public string TeamId { get; set; }

        public string Serialize()
        {
            return JsonConvert.SerializeObject(this);
        }
    }
}
