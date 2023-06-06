using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace Microsoft.BotBuilderSamples.Models.dto
{
    public class GetFailedEntriesResponseDto
    {
        [JsonProperty("ContinuationToken")]
        public string ContinuationToken { get; set; }
        [JsonProperty("FailedEntryResponses")]
        public IEnumerable<OperationFailedEntryInfoDto> FailedEntryResponses { get; set; }

        public static GetFailedEntriesResponse Deserialize(string json)
        {
            return JsonConvert.DeserializeObject<GetFailedEntriesResponse>(json);
        }

    }
}
