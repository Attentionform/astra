// Worker process for ASTra OOXML Validator. See LICENSE-MIT-xarsh.

using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace OoxmlValidator;

public static class Worker
{
  private const long DefaultMaxDocBytes = 64 * 1024 * 1024; // 64 MiB
  private const int DefaultMaxErrors = 200;

  private static readonly Dictionary<FileFormatVersions, OpenXmlValidator> Validators = new();
  private static readonly string SdkVersion = ResolveSdkVersion();

  private static readonly JsonSerializerOptions JsonOptions = new()
  {
    PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    WriteIndented = false
  };

  private static string ResolveSdkVersion()
  {
    try
    {
      var attr = typeof(OpenXmlValidator).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>();
      if (!string.IsNullOrEmpty(attr?.InformationalVersion))
      {
        var v = attr.InformationalVersion;
        var plus = v.IndexOf('+');
        return plus > 0 ? v[..plus] : v;
      }
      return typeof(OpenXmlValidator).Assembly.GetName().Version?.ToString() ?? "3.5.1";
    }
    catch
    {
      return "3.5.1";
    }
  }

  private static OpenXmlValidator GetValidator(FileFormatVersions version)
  {
    if (!Validators.TryGetValue(version, out var validator))
    {
      validator = new OpenXmlValidator(version);
      Validators[version] = validator;
    }
    return validator;
  }

  public static int Run(string[]? _ = null)
  {
    using var stdin = Console.OpenStandardInput();
    using var stdout = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false))
    {
      AutoFlush = true
    };

    while (true)
    {
      string? headerLine = ReadLineUtf8(stdin);
      if (headerLine == null)
      {
        // EOF on stdin - clean exit
        return 0;
      }

      headerLine = headerLine.Trim();
      if (string.IsNullOrEmpty(headerLine))
      {
        continue;
      }

      WorkerRequestHeader? header;
      try
      {
        header = JsonSerializer.Deserialize<WorkerRequestHeader>(headerLine, JsonOptions);
        if (header == null)
        {
          throw new FormatException("Null header object");
        }
      }
      catch (Exception ex)
      {
        var fatalResp = new
        {
          id = (long?)null,
          ok = false,
          fatal = true,
          error = $"Unparsable header JSON: {ex.Message}"
        };
        stdout.WriteLine(JsonSerializer.Serialize(fatalResp, JsonOptions));
        return 1;
      }

      var op = (header.Op ?? "validate").ToLowerInvariant();

      if (op == "shutdown")
      {
        var shutdownResp = new WorkerResponse
        {
          Id = header.Id,
          Ok = true
        };
        stdout.WriteLine(JsonSerializer.Serialize(shutdownResp, JsonOptions));
        return 0;
      }

      if (op == "ping")
      {
        var pingResp = new WorkerResponse
        {
          Id = header.Id,
          Ok = true,
          Pong = true
        };
        stdout.WriteLine(JsonSerializer.Serialize(pingResp, JsonOptions));
        continue;
      }

      if (op == "version")
      {
        var verResp = new WorkerResponse
        {
          Id = header.Id,
          Ok = true,
          SdkVersion = SdkVersion
        };
        stdout.WriteLine(JsonSerializer.Serialize(verResp, JsonOptions));
        continue;
      }

      if (op != "validate")
      {
        var unknownResp = new WorkerResponse
        {
          Id = header.Id,
          Ok = false,
          ErrorCount = 1,
          Errors = new List<ValidationErrorDto>
          {
            ValidationErrorDto.FromException($"Unknown operation '{header.Op}'")
          }
        };
        stdout.WriteLine(JsonSerializer.Serialize(unknownResp, JsonOptions));
        continue;
      }

      // op == "validate"
      long size = header.Size ?? 0;
      long maxBytes = header.MaxDocBytes ?? DefaultMaxDocBytes;
      int maxErrors = header.MaxErrors ?? DefaultMaxErrors;
      var officeStr = header.Office ?? "Office2019";
      var ffVersion = Program.ParseOfficeVersion(officeStr);

      if (ffVersion == null)
      {
        // If invalid office version, discard body if size > 0
        if (size > 0)
        {
          DiscardBytes(stdin, size);
        }
        var invalidResp = new WorkerResponse
        {
          Id = header.Id,
          Ok = false,
          SdkVersion = SdkVersion,
          Office = officeStr,
          ErrorCount = 1,
          Errors = new List<ValidationErrorDto>
          {
            ValidationErrorDto.FromException($"Invalid Office version '{officeStr}'.")
          }
        };
        stdout.WriteLine(JsonSerializer.Serialize(invalidResp, JsonOptions));
        continue;
      }

      if (size > maxBytes)
      {
        // Discard oversized body so stream stays in sync
        DiscardBytes(stdin, size);
        var oversizedResp = new WorkerResponse
        {
          Id = header.Id,
          Ok = false,
          SdkVersion = SdkVersion,
          Office = officeStr,
          ErrorCount = 1,
          Errors = new List<ValidationErrorDto>
          {
            ValidationErrorDto.FromException($"Document size {size} bytes exceeds maximum allowed {maxBytes} bytes.")
          }
        };
        stdout.WriteLine(JsonSerializer.Serialize(oversizedResp, JsonOptions));
        continue;
      }

      byte[] buffer;
      try
      {
        buffer = ReadExactBytes(stdin, size);
      }
      catch (Exception ex)
      {
        var fatalResp = new
        {
          id = header.Id,
          ok = false,
          fatal = true,
          error = $"Failed to read {size} bytes from stdin: {ex.Message}"
        };
        stdout.WriteLine(JsonSerializer.Serialize(fatalResp, JsonOptions));
        return 1;
      }

      var sw = Stopwatch.StartNew();
      var resp = ValidateMemory(header.Id, buffer, header.Kind ?? "docx", ffVersion.Value, officeStr, maxErrors);
      sw.Stop();
      resp.DurationMs = sw.ElapsedMilliseconds;

      stdout.WriteLine(JsonSerializer.Serialize(resp, JsonOptions));
    }
  }

  private static WorkerResponse ValidateMemory(
    long? id,
    byte[] bytes,
    string kind,
    FileFormatVersions version,
    string officeStr,
    int maxErrors)
  {
    var resp = new WorkerResponse
    {
      Id = id,
      SdkVersion = SdkVersion,
      Office = officeStr,
      Errors = new List<ValidationErrorDto>()
    };

    try
    {
      using var stream = new MemoryStream(bytes, false);
      using var doc = OpenPackageFromStream(stream, kind);
      var validator = GetValidator(version);

      int count = 0;
      bool truncated = false;

      foreach (var err in validator.Validate(doc))
      {
        if (count >= maxErrors)
        {
          truncated = true;
          break;
        }

        resp.Errors.Add(ValidationErrorDto.FromValidationErrorInfo(err));
        count++;
      }

      resp.Truncated = truncated;
      resp.ErrorCount = count;
      resp.Ok = count == 0;
    }
    catch (Exception ex)
    {
      resp.Ok = false;
      resp.ErrorCount = 1;
      resp.Errors.Add(ValidationErrorDto.FromException(ex.Message));
    }

    return resp;
  }

  private static OpenXmlPackage OpenPackageFromStream(Stream stream, string kind)
  {
    var k = (kind ?? "docx").ToLowerInvariant();
    return k switch
    {
      "docx" or ".docx" or "docm" or "dotx" or "dotm"
          => WordprocessingDocument.Open(stream, false),
      "pptx" or ".pptx" or "pptm" or "potx" or "potm" or "ppsx" or "ppsm"
          => PresentationDocument.Open(stream, false),
      "xlsx" or ".xlsx" or "xlsm" or "xltx" or "xltm" or "xlam"
          => SpreadsheetDocument.Open(stream, false),
      _ => throw new InvalidOperationException($"Unsupported document kind: {kind}")
    };
  }

  private static string? ReadLineUtf8(Stream stream)
  {
    var ms = new MemoryStream();
    while (true)
    {
      int b = stream.ReadByte();
      if (b == -1)
      {
        if (ms.Length == 0) return null;
        break;
      }
      if (b == '\n')
      {
        break;
      }
      if (b != '\r')
      {
        ms.WriteByte((byte)b);
      }
    }
    return Encoding.UTF8.GetString(ms.ToArray());
  }

  private static byte[] ReadExactBytes(Stream stream, long count)
  {
    if (count == 0) return Array.Empty<byte>();
    var buffer = new byte[count];
    long totalRead = 0;
    while (totalRead < count)
    {
      int read = stream.Read(buffer, (int)totalRead, (int)(count - totalRead));
      if (read == 0)
      {
        throw new EndOfStreamException($"Expected {count} bytes, but stream ended at {totalRead} bytes.");
      }
      totalRead += read;
    }
    return buffer;
  }

  private static void DiscardBytes(Stream stream, long count)
  {
    var scratch = new byte[Math.Min(count, 65536)];
    long remaining = count;
    while (remaining > 0)
    {
      int toRead = (int)Math.Min(remaining, scratch.Length);
      int read = stream.Read(scratch, 0, toRead);
      if (read == 0)
      {
        break;
      }
      remaining -= read;
    }
  }
}

public sealed class WorkerRequestHeader
{
  public long? Id { get; set; }
  public string? Op { get; set; }
  public string? Kind { get; set; }
  public string? Office { get; set; }
  public long? Size { get; set; }
  public int? MaxErrors { get; set; }
  public long? MaxDocBytes { get; set; }
}

public sealed class WorkerResponse
{
  public long? Id { get; set; }
  public bool Ok { get; set; }
  public bool? Pong { get; set; }
  public string? SdkVersion { get; set; }
  public string? Office { get; set; }
  public long? DurationMs { get; set; }
  public int? ErrorCount { get; set; }
  public bool? Truncated { get; set; }
  public List<ValidationErrorDto>? Errors { get; set; }
}
