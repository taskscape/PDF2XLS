using System.Text.Json.Serialization;

namespace PDF2XLS;

public class ToolModel
{
    [JsonInclude]
    [JsonPropertyName("type")]
    public string Type { get; set; }
}