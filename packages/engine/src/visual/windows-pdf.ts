/**
 * Windows built-in PDF renderer (Windows.Data.Pdf) as a PDF -> page-image step.
 *
 * Runs one embedded Windows PowerShell 5.1 script per call that loads the PDF through WinRT and renders every
 * selected page, so there is no per-page start cost. Needs Windows 10 or later. It is independent of how the
 * PDF was made (Word, LibreOffice or anything else). Poppler stays the fallback on every platform and the
 * explicit choice for users who want identical pixels on every OS: output is NOT pixel-identical to Poppler.
 */

import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
  POWERSHELL_PRELUDE,
  RunQueue,
  powershellPath,
  runPowerShell,
  sanitiseDetail,
} from './ps-runner';
import type { FailureClass } from './renderer-preference';

/** Widest image the Windows renderer produces (a very large page at 300 dpi is capped, Poppler is unchanged). */
export const WINDOWS_PDF_MAX_WIDTH = 8000;

export interface WindowsRendererProbe {
  platform?: NodeJS.Platform;
  release?: string;
  exists?: (p: string) => boolean;
}

/** Can the Windows renderer be attempted? Cheap: nothing is started. Runtime failures still fall through to Poppler. */
export function windowsRendererAvailable(o: WindowsRendererProbe = {}): boolean {
  if ((o.platform ?? process.platform) !== 'win32') return false;
  const major = Number.parseInt((o.release ?? os.release()).split('.')[0] ?? '', 10);
  if (!Number.isFinite(major) || major < 10) return false;
  return (o.exists ?? ((p: string) => fs.existsSync(p)))(powershellPath());
}

/**
 * Exit codes: 0 ok, 12 constrained language, 41 Windows.Data.Pdf not present, 42 PDF could not be loaded
 * (password-protected or damaged), 43 a page failed to render.
 */
export function buildWindowsPdfScript(): string {
  return `${POWERSHELL_PRELUDE}
$pdf = $env:ASTRA_PDF_IN; $dir = $env:ASTRA_IMG_DIR; $fmt = $env:ASTRA_IMG_FORMAT; $dpi = [double]$env:ASTRA_IMG_DPI
$q = [int]$env:ASTRA_IMG_QUALITY; $max = [int]$env:ASTRA_IMG_MAX; $maxW = [double]$env:ASTRA_IMG_MAXWIDTH
$scale = 1.0
$raw = @($env:ASTRA_IMG_PAGES -split ',' | Where-Object { $_ } | ForEach-Object { [int]$_ })
try {
  Add-Type -AssemblyName System.Runtime.WindowsRuntime
  $null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
  $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
  $null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
  $op = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation\`1' })[0]
  $act = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' })[0]
} catch { Die 41 $_ }
function Await($o, $t) { $x = $op.MakeGenericMethod($t).Invoke($null, @($o)); $x.Wait(-1) | Out-Null; $x.Result }
try {
  $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($pdf)) ([Windows.Storage.StorageFile])
  $doc = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
} catch { Die 42 $_ }
$total = [int]$doc.PageCount
Say "TOTAL $total"
if ($raw.Count -eq 0) { $sel = @(1..[Math]::Min($total, $max)) }
else { $sel = @($raw | ForEach-Object { [Math]::Max(1, [Math]::Min($total, $_)) } | Sort-Object -Unique | Select-Object -First $max) }
foreach ($n in $sel) {
  try {
    $page = $doc.GetPage($n - 1)
    $o = New-Object Windows.Data.Pdf.PdfPageRenderOptions
    # The pixel width that Poppler would give: the page width in points times dpi / 72 (Size is in 1/96 inch).
    $want = [Math]::Max(1, [Math]::Min($maxW, [Math]::Round($page.Size.Width * $dpi / 96)))
    # Windows.Data.Pdf multiplies the requested width by the system display scale (125 percent on a 120 dpi
    # display), so the first render is measured and, when it is off, requested again with the factor removed.
    # The factor is the same for every page of the run.
    for ($try = 0; $try -lt 3; $try++) {
      $o.DestinationWidth = [uint32][Math]::Max(1, [Math]::Round($want / $scale))
      $ms = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
      $t = $act.Invoke($null, @($page.RenderToStreamAsync($ms, $o))); $t.Wait(-1) | Out-Null
      $bytes = New-Object byte[] $ms.Size; $ms.Seek(0)
      $r = New-Object Windows.Storage.Streams.DataReader($ms); $null = Await ($r.LoadAsync([uint32]$ms.Size)) ([uint32]); $r.ReadBytes($bytes)
      $got = ([int]$bytes[16] -shl 24) -bor ([int]$bytes[17] -shl 16) -bor ([int]$bytes[18] -shl 8) -bor [int]$bytes[19]
      if ([Math]::Abs($got - $want) -le 1 -or $got -lt 1) { break }
      $measured = $got / [double]$o.DestinationWidth
      if ($measured -lt 0.2 -or $measured -gt 8) { break }
      $scale = $measured
    }
    $out = Join-Path $dir ("page-{0}.{1}" -f $n, $(if ($fmt -eq 'jpeg') { 'jpg' } else { 'png' }))
    if ($fmt -eq 'jpeg') {
      Add-Type -AssemblyName System.Drawing
      $img = [System.Drawing.Image]::FromStream((New-Object IO.MemoryStream (, $bytes)))
      $enc = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object MimeType -eq 'image/jpeg'
      $ep = New-Object System.Drawing.Imaging.EncoderParameters(1)
      $ep.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, [long]$q)
      $img.Save($out, $enc, $ep); $img.Dispose()
    } else { [IO.File]::WriteAllBytes($out, $bytes) }
    $page.Dispose()
    Say "PAGE $n"
  } catch { Die 43 $_ }
}
Say 'OK'
`;
}

export interface WindowsPdfProtocol {
  total?: number;
  pages: number[];
  ok: boolean;
  err?: { code: number; hresult: string; message: string };
}

export function parseWindowsPdfProtocol(lines: string[]): WindowsPdfProtocol {
  const p: WindowsPdfProtocol = { pages: [], ok: false };
  for (const line of lines) {
    let m: RegExpExecArray | null;
    if ((m = /^TOTAL (\d+)$/.exec(line))) p.total = Number(m[1]);
    else if ((m = /^PAGE (\d+)$/.exec(line))) p.pages.push(Number(m[1]));
    else if (line === 'OK') p.ok = true;
    else if ((m = /^ERR (\d+) (0x[0-9A-Fa-f]{8}) ?(.*)$/.exec(line))) p.err = { code: Number(m[1]), hresult: `0x${m[2].slice(2).toUpperCase()}`, message: sanitiseDetail(m[3]) };
  }
  return p;
}

export class WindowsPdfError extends Error {
  readonly code: 'E_DEPENDENCY' | 'E_VALIDATION';
  readonly remedy: string;
  readonly failureClass: FailureClass;
  readonly exitCode?: number;
  constructor(code: WindowsPdfError['code'], failureClass: FailureClass, message: string, remedy: string, exitCode?: number) {
    super(message);
    this.name = 'WindowsPdfError';
    this.code = code;
    this.failureClass = failureClass;
    this.remedy = remedy;
    this.exitCode = exitCode;
  }
}

export interface WindowsPdfOptions {
  format: 'png' | 'jpeg';
  dpi: number;
  jpegQuality: number;
  /** Raw 1-based page request; clamped to the real page count by the script, like Poppler's selection. */
  pages?: number[];
  maxPages: number;
  timeoutMs?: number;
  /** Tests only. */
  powershellBin?: string;
  powershellArgs?: string[];
}

export interface WindowsPdfOutcome {
  totalPages: number;
  images: Array<{ page: number; file: string }>;
}

const windowsPdfQueue = new RunQueue();

export function windowsPdfToImages(pdfPath: string, outDir: string, o: WindowsPdfOptions): Promise<WindowsPdfOutcome> {
  return windowsPdfQueue.run(async () => {
    const ext = o.format === 'jpeg' ? 'jpg' : 'png';
    // Same normalisation the Poppler path applies (finite numbers only, floored, at least 1); the script
    // clamps to the real page count, which only it knows.
    const pages = (o.pages ?? [])
      .map((n) => Number(n))
      .filter((n) => Number.isFinite(n))
      .map((n) => Math.min(1_000_000, Math.max(1, Math.floor(n))));
    const res = await runPowerShell({
      script: buildWindowsPdfScript(),
      env: {
        ASTRA_PDF_IN: pdfPath,
        ASTRA_IMG_DIR: outDir,
        ASTRA_IMG_FORMAT: o.format,
        ASTRA_IMG_DPI: String(o.dpi),
        ASTRA_IMG_QUALITY: String(o.jpegQuality),
        ASTRA_IMG_MAX: String(o.maxPages),
        ASTRA_IMG_MAXWIDTH: String(WINDOWS_PDF_MAX_WIDTH),
        ASTRA_IMG_PAGES: pages.join(','),
      },
      timeoutMs: o.timeoutMs ?? 60000,
      powershellBin: o.powershellBin,
      powershellArgs: o.powershellArgs,
    });
    const proto = parseWindowsPdfProtocol(res.lines);
    if (res.spawnError) {
      throw new WindowsPdfError(
        'E_DEPENDENCY',
        'infrastructure',
        `E_DEPENDENCY: Windows PowerShell could not be started (${sanitiseDetail(res.spawnError.code ?? res.spawnError.message, 80)}).`,
        'PowerShell is blocked or missing; use --pdf-to-image poppler.'
      );
    }
    if (res.timedOut) {
      throw new WindowsPdfError('E_DEPENDENCY', 'infrastructure', 'E_DEPENDENCY: the Windows PDF renderer did not respond in time.', 'Use --pdf-to-image poppler.');
    }
    if (res.exitCode !== 0 || !proto.ok) {
      const e = proto.err;
      if (e) {
        if (e.code === 12) {
          throw new WindowsPdfError('E_DEPENDENCY', 'infrastructure', 'E_DEPENDENCY: PowerShell is in constrained language mode (policy), so the Windows PDF renderer cannot run.', 'Use --pdf-to-image poppler.', 12);
        }
        if (e.code === 41) {
          throw new WindowsPdfError('E_DEPENDENCY', 'infrastructure', 'E_DEPENDENCY: the Windows PDF renderer (Windows.Data.Pdf) is not available on this Windows.', 'Use --pdf-to-image poppler.', 41);
        }
        if (e.code === 42) {
          throw new WindowsPdfError('E_VALIDATION', 'document', `E_VALIDATION: the Windows PDF renderer could not load the PDF (${e.hresult}); it may be password-protected or damaged.`, 'Verify the converted PDF is readable and retry.', 42);
        }
        throw new WindowsPdfError('E_VALIDATION', 'document', `E_VALIDATION: the Windows PDF renderer could not render a page (${e.hresult}: ${e.message}).`, 'Verify the converted PDF is readable and retry.', e.code);
      }
      throw new WindowsPdfError('E_DEPENDENCY', 'infrastructure', `E_DEPENDENCY: the Windows PDF renderer ended without a result (exit code ${res.exitCode ?? 'none'}).`, 'Use --pdf-to-image poppler.');
    }
    const total = proto.total;
    if (total === undefined || total < 1 || proto.pages.length === 0) {
      throw new WindowsPdfError('E_VALIDATION', 'document', 'E_VALIDATION: the Windows PDF renderer produced no page images.', 'Verify the document has at least one page and retry.');
    }
    const images = proto.pages
      .map((page) => ({ page, file: path.join(outDir, `page-${page}.${ext}`) }))
      .filter((i) => fs.existsSync(i.file) && fs.statSync(i.file).size > 0);
    if (images.length !== proto.pages.length) {
      throw new WindowsPdfError('E_DEPENDENCY', 'infrastructure', 'E_DEPENDENCY: the Windows PDF renderer reported pages that were not written.', 'Use --pdf-to-image poppler.');
    }
    return { totalPages: total, images };
  });
}
