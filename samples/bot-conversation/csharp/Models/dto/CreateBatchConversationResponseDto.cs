using Newtonsoft.Json;

namespace Microsoft.BotBuilderSamples.Models.dto
{
    public class CreateBatchConversationResponseDto
    {
        [JsonProperty("OperationId")]
        public string OperationId { get; set; }

        public static CreateBatchConversationResponse Deserialize(string json)
        {
            return JsonConvert.DeserializeObject<CreateBatchConversationResponse>(json);
        }
    }
}
