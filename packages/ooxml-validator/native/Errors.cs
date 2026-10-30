namespace OoxmlValidator;

using System.Text.Json.Serialization;
using DocumentFormat.OpenXml.Validation;

public sealed class ErrorNodeDto
{
  [JsonPropertyName("localName")]
  public string? LocalName { get; set; }

  [JsonPropertyName("namespaceUri")]
  public string? NamespaceUri { get; set; }
}

public sealed class ValidationErrorDto
{
  [JsonPropertyName("description")]
  public string? Description { get; set; }

  [JsonPropertyName("path")]
  public string? Path { get; set; }

  [JsonPropertyName("xPath")]
  public string? XPath { get; set; }

  [JsonPropertyName("id")]
  public string? Id { get; set; }

  [JsonPropertyName("errorType")]
  public string? ErrorType { get; set; }

  [JsonPropertyName("part")]
  public string? Part { get; set; }

  [JsonPropertyName("class")]
  public string? Class { get; set; }

  [JsonPropertyName("node")]
  public ErrorNodeDto? Node { get; set; }

  [JsonPropertyName("relatedNode")]
  public ErrorNodeDto? RelatedNode { get; set; }

  public static ValidationErrorDto FromValidationErrorInfo(ValidationErrorInfo error)
  {
    var partUri = error.Part?.Uri.ToString();
    var dto = new ValidationErrorDto
    {
      Description = error.Description,
      Path = partUri,
      Part = partUri,
      XPath = error.Path?.XPath,
      Id = error.Id,
      ErrorType = error.ErrorType.ToString(),
      Class = MapClass(error.ErrorType)
    };

    if (error.Node != null)
    {
      dto.Node = new ErrorNodeDto
      {
        LocalName = error.Node.LocalName,
        NamespaceUri = error.Node.NamespaceUri
      };
    }

    if (error.RelatedNode != null)
    {
      dto.RelatedNode = new ErrorNodeDto
      {
        LocalName = error.RelatedNode.LocalName,
        NamespaceUri = error.RelatedNode.NamespaceUri
      };
    }

    return dto;
  }

  public static ValidationErrorDto FromException(string message)
  {
    return new ValidationErrorDto
    {
      Description = message,
      ErrorType = "Exception",
      Class = "exception"
    };
  }

  public static string MapClass(ValidationErrorType errorType)
  {
    return errorType switch
    {
      ValidationErrorType.Schema => "schema",
      ValidationErrorType.Semantic => "semantic",
      ValidationErrorType.MarkupCompatibility => "markup-compat",
      ValidationErrorType.Package => "package",
      _ => errorType.ToString().ToLowerInvariant()
    };
  }
}
