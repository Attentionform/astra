using System.Diagnostics;
using System.Text;
using System.Text.Json;
using Xunit;

namespace OoxmlValidator.Tests;

public class WorkerTests
{
  private static readonly string FixturesDir = Path.GetFullPath(
    Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "tests", "fixtures")
  );

  private static readonly string BinaryPath = ResolveBinaryPath();

  private static string ResolveBinaryPath()
  {
    var nativeDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
    var candidates = new[]
    {
      Path.Combine(nativeDir, "out", "linux-x64", "ooxml-validator"),
      Path.Combine(nativeDir, "bin", "Release", "net10.0", "linux-x64", "ooxml-validator"),
      Path.Combine(nativeDir, "bin", "Debug", "net10.0", "ooxml-validator"),
    };

    foreach (var c in candidates)
    {
      if (File.Exists(c)) return c;
    }

    return candidates[0];
  }

  private static Process StartWorker()
  {
    var psi = new ProcessStartInfo
    {
      FileName = BinaryPath,
      Arguments = "--worker",
      RedirectStandardInput = true,
      RedirectStandardOutput = true,
      RedirectStandardError = true,
      UseShellExecute = false,
      StandardOutputEncoding = new UTF8Encoding(false),
    };
    return Process.Start(psi) ?? throw new InvalidOperationException("Failed to start worker process");
  }

  private static JsonElement SendRequest(Process p, object header, byte[]? body = null)
  {
    var jsonHeader = JsonSerializer.Serialize(header);
    p.StandardInput.WriteLine(jsonHeader);
    p.StandardInput.Flush();

    if (body is { Length: > 0 })
    {
      p.StandardInput.BaseStream.Write(body, 0, body.Length);
      p.StandardInput.BaseStream.Flush();
    }

    var line = p.StandardOutput.ReadLine();
    Assert.NotNull(line);
    using var doc = JsonDocument.Parse(line);
    return doc.RootElement.Clone();
  }

  [Fact]
  public void ValidDocx_ReturnsOkTrue()
  {
    using var p = StartWorker();
    try
    {
      var bytes = File.ReadAllBytes(Path.Combine(FixturesDir, "valid.docx"));
      var resp = SendRequest(p, new
      {
        id = 101,
        op = "validate",
        kind = "docx",
        office = "Office2019",
        size = bytes.Length
      }, bytes);

      Assert.True(resp.GetProperty("ok").GetBoolean());
      Assert.Equal(101, resp.GetProperty("id").GetInt64());
      Assert.Equal(0, resp.GetProperty("errorCount").GetInt32());
      Assert.Equal(0, resp.GetProperty("errors").GetArrayLength());
    }
    finally
    {
      if (!p.HasExited) p.Kill();
    }
  }

  [Fact]
  public void InvalidSchema_ReturnsOkFalseWithErrors()
  {
    using var p = StartWorker();
    try
    {
      var bytes = File.ReadAllBytes(Path.Combine(FixturesDir, "invalid-schema.docx"));
      var resp = SendRequest(p, new
      {
        id = 102,
        op = "validate",
        kind = "docx",
        office = "Office2019",
        size = bytes.Length
      }, bytes);

      Assert.False(resp.GetProperty("ok").GetBoolean());
      Assert.Equal(102, resp.GetProperty("id").GetInt64());
      Assert.True(resp.GetProperty("errorCount").GetInt32() > 0);
      var errors = resp.GetProperty("errors");
      Assert.True(errors.GetArrayLength() > 0);
      Assert.Equal("Schema", errors[0].GetProperty("errorType").GetString());
    }
    finally
    {
      if (!p.HasExited) p.Kill();
    }
  }

  [Fact]
  public void CorruptZip_ReturnsOkFalseWithExceptionError()
  {
    using var p = StartWorker();
    try
    {
      var bytes = File.ReadAllBytes(Path.Combine(FixturesDir, "corrupt-zip.docx"));
      var resp = SendRequest(p, new
      {
        id = 103,
        op = "validate",
        kind = "docx",
        office = "Office2019",
        size = bytes.Length
      }, bytes);

      Assert.False(resp.GetProperty("ok").GetBoolean());
      Assert.Equal(103, resp.GetProperty("id").GetInt64());
      Assert.Equal(1, resp.GetProperty("errorCount").GetInt32());
      var errors = resp.GetProperty("errors");
      Assert.Equal("Exception", errors[0].GetProperty("errorType").GetString());
    }
    finally
    {
      if (!p.HasExited) p.Kill();
    }
  }

  [Fact]
  public void OversizedDoc_RejectsAndKeepsWorkerAlive()
  {
    using var p = StartWorker();
    try
    {
      var fakeBody = new byte[100];
      var resp = SendRequest(p, new
      {
        id = 104,
        op = "validate",
        kind = "docx",
        office = "Office2019",
        size = fakeBody.Length,
        maxDocBytes = 50 // less than 100
      }, fakeBody);

      Assert.False(resp.GetProperty("ok").GetBoolean());
      Assert.Contains("exceeds maximum allowed", resp.GetProperty("errors")[0].GetProperty("description").GetString());

      // Worker should still be alive and accept next request
      var pingResp = SendRequest(p, new { id = 105, op = "ping" });
      Assert.True(pingResp.GetProperty("ok").GetBoolean());
      Assert.True(pingResp.GetProperty("pong").GetBoolean());
    }
    finally
    {
      if (!p.HasExited) p.Kill();
    }
  }

  [Fact]
  public void TruncatedBody_ReturnsFatalAndExits()
  {
    using var p = StartWorker();
    try
    {
      // Header claims 100 bytes, but we only write 10 bytes and close input stream
      p.StandardInput.WriteLine("{\"id\":106,\"op\":\"validate\",\"size\":100}");
      p.StandardInput.Flush();
      p.StandardInput.BaseStream.Write(new byte[10], 0, 10);
      p.StandardInput.BaseStream.Flush();
      p.StandardInput.BaseStream.Close(); // close stream prematurely

      var line = p.StandardOutput.ReadLine();
      Assert.NotNull(line);
      using var doc = JsonDocument.Parse(line);
      Assert.True(doc.RootElement.GetProperty("fatal").GetBoolean());
      Assert.False(doc.RootElement.GetProperty("ok").GetBoolean());

      p.WaitForExit(5000);
      Assert.True(p.HasExited);
      Assert.NotEqual(0, p.ExitCode);
    }
    finally
    {
      if (!p.HasExited) p.Kill();
    }
  }

  [Fact]
  public void MaxErrors_TruncatesErrorList()
  {
    using var p = StartWorker();
    try
    {
      var bytes = File.ReadAllBytes(Path.Combine(FixturesDir, "many-errors.docx"));
      var resp = SendRequest(p, new
      {
        id = 107,
        op = "validate",
        kind = "docx",
        office = "Office2019",
        size = bytes.Length,
        maxErrors = 2
      }, bytes);

      Assert.Equal(107, resp.GetProperty("id").GetInt64());
      Assert.True(resp.GetProperty("errors").GetArrayLength() <= 2);
      if (resp.GetProperty("errorCount").GetInt32() >= 2)
      {
        Assert.True(resp.GetProperty("truncated").GetBoolean());
      }
    }
    finally
    {
      if (!p.HasExited) p.Kill();
    }
  }

  [Fact]
  public void UnknownOp_ReturnsErrorWithoutCrashing()
  {
    using var p = StartWorker();
    try
    {
      var resp = SendRequest(p, new
      {
        id = 108,
        op = "nonExistentOperation"
      });

      Assert.False(resp.GetProperty("ok").GetBoolean());
      Assert.Contains("Unknown operation", resp.GetProperty("errors")[0].GetProperty("description").GetString());

      // Worker still alive
      var ping = SendRequest(p, new { id = 109, op = "ping" });
      Assert.True(ping.GetProperty("pong").GetBoolean());
    }
    finally
    {
      if (!p.HasExited) p.Kill();
    }
  }

  [Fact]
  public void RicherErrors_SchemaErrorContainsNodePartAndClass()
  {
    using var p = StartWorker();
    try
    {
      var bytes = File.ReadAllBytes(Path.Combine(FixturesDir, "invalid-schema.docx"));
      var resp = SendRequest(p, new
      {
        id = 110,
        op = "validate",
        kind = "docx",
        office = "Office2019",
        size = bytes.Length
      }, bytes);

      Assert.False(resp.GetProperty("ok").GetBoolean());
      var errors = resp.GetProperty("errors");
      Assert.True(errors.GetArrayLength() > 0);

      var err = errors[0];
      Assert.Equal("schema", err.GetProperty("class").GetString());
      Assert.Equal("Schema", err.GetProperty("errorType").GetString());
      Assert.Equal("/word/document.xml", err.GetProperty("part").GetString());
      Assert.Equal("/word/document.xml", err.GetProperty("path").GetString());

      Assert.True(err.TryGetProperty("node", out var node));
      Assert.Equal("body", node.GetProperty("localName").GetString());
      Assert.Equal("http://schemas.openxmlformats.org/wordprocessingml/2006/main", node.GetProperty("namespaceUri").GetString());

      Assert.True(err.TryGetProperty("relatedNode", out var relNode));
      Assert.Equal("bogusElementWithInvalidSchema", relNode.GetProperty("localName").GetString());
    }
    finally
    {
      if (!p.HasExited) p.Kill();
    }
  }

  [Fact]
  public void RicherErrors_MultiErrorDocumentHasClassAndPartOnEveryError()
  {
    using var p = StartWorker();
    try
    {
      var bytes = File.ReadAllBytes(Path.Combine(FixturesDir, "many-errors.docx"));
      var resp = SendRequest(p, new
      {
        id = 111,
        op = "validate",
        kind = "docx",
        office = "Office2007",
        size = bytes.Length,
        maxErrors = 10
      }, bytes);

      Assert.False(resp.GetProperty("ok").GetBoolean());
      var errors = resp.GetProperty("errors");
      Assert.True(errors.GetArrayLength() > 0);

      foreach (var err in errors.EnumerateArray())
      {
        Assert.NotNull(err.GetProperty("class").GetString());
        Assert.NotNull(err.GetProperty("part").GetString());
        Assert.True(err.TryGetProperty("node", out _));
      }
    }
    finally
    {
      if (!p.HasExited) p.Kill();
    }
  }
}
