using Microsoft.BotBuilderSamples.Models.dto;
using System.Collections.Generic;

namespace Microsoft.BotBuilderSamples.Models
{
    public class GetFailedEntriesResponse
    {
        public string ContinuationToken { get; set; }
        public IEnumerable<OperationFailedEntryInfoDto> FailedEntryResponses { get; set; }
    }
}
