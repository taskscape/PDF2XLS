using System.Text.Json.Serialization;

namespace PDF2XLS.Models;

public class PromptRequestModel
{
    [JsonInclude]
    [JsonPropertyName("role")]
    public string Role { get; set; }
    [JsonInclude]
    [JsonPropertyName("content")]
    public string Content { get; set; }
    [JsonInclude]
    [JsonPropertyName("attachments")]
    public List<AttachmentModel> Attachments { get; set; }
}