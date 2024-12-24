using System.Text.Json.Serialization;

namespace PDF2XLS;

public class AttachmentModel
{
    [JsonInclude]
    [JsonPropertyName("file_id")]
    public string FileId { get; set; }
    [JsonInclude]
    [JsonPropertyName("tools")]
    public List<ToolModel> Tools { get; set; }
}