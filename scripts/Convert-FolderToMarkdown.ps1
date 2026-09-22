<#
.SYNOPSIS
    Batch-converts every file in a folder to Markdown by invoking the markitdown CLI once per file.

.DESCRIPTION
    Enumerates the files in -InputFolder, runs the compiled `markitdown` command for each one, and
    writes the results into a separate folder created next to the original (<InputFolder>-md by
    default). The source folder is never modified.

    The markitdown executable is located automatically: -MarkItDownPath if given, then a .venv next
    to this script or its parent folder, then `markitdown` on PATH, then `python -m markitdown`.

.PARAMETER InputFolder
    Folder holding the files to convert.

.PARAMETER OutputFolder
    Destination folder. Defaults to a sibling of the input folder named <InputFolder><Suffix>.

.PARAMETER Suffix
    Suffix used to build the default output folder name. Default: '-md'.

.PARAMETER Recurse
    Also convert files in subfolders, mirroring the folder tree in the output.

.PARAMETER Include
    Only convert these extensions, e.g. -Include .pdf,.docx. Default: every file.

.PARAMETER Flat
    With -Recurse, write every result into the top-level output folder instead of mirroring the tree.

.PARAMETER Force
    Overwrite existing .md files. Without it, files that already exist are left alone and reported
    as skipped, so an interrupted run can simply be started again.

.PARAMETER MarkItDownPath
    Explicit path to markitdown.exe (or a command name on PATH) to use instead of auto-detection.

.PARAMETER UsePlugins
    Pass -p to markitdown so installed third-party plugins are used.

.PARAMETER KeepDataUris
    Pass --keep-data-uris to markitdown so base64 image data is kept instead of truncated.

.PARAMETER ExtractImages
    Write embedded images out as real image files and rewrite the Markdown links to point at
    them. Implies -KeepDataUris, since there is nothing to extract without it. By default each
    document gets its own <name>.assets folder beside its .md file.

.PARAMETER AssetsFolder
    With -ExtractImages, collect every image into this one folder under the output root
    (e.g. -AssetsFolder assets) instead of a per-document .assets folder.

.PARAMETER MinImagePixels
    With -ExtractImages, ignore images smaller than this many pixels (width x height).
    Default 0 keeps everything; try 10000 (100x100) to drop icons and rules from PDFs.

.PARAMETER NoImageStitching
    PDFs often store one picture as several image XObjects laid edge to edge, which would
    otherwise be extracted as separate fragments. They are stitched back together by default;
    this switch turns that off and writes every fragment as its own file.

.EXAMPLE
    .\Convert-FolderToMarkdown.ps1 -InputFolder C:\data\docs
    Converts C:\data\docs\* into C:\data\docs-md\.

.EXAMPLE
    .\Convert-FolderToMarkdown.ps1 C:\data\docs -Recurse -Include .pdf,.docx -Force

.EXAMPLE
    .\Convert-FolderToMarkdown.ps1 C:\data\docs -Recurse -ExtractImages
    Writes notes.md alongside notes.assets\notes-001.png, with links pointing at the file.

.EXAMPLE
    .\Convert-FolderToMarkdown.ps1 C:\data\docs -Recurse -ExtractImages -AssetsFolder assets
    Same, but every image from every document lands in C:\data\docs-md\assets\.

.EXAMPLE
    .\Convert-FolderToMarkdown.ps1 C:\data\docs -WhatIf
    Lists what would be converted, without running anything.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $InputFolder,

    [Parameter(Position = 1)]
    [string] $OutputFolder,

    [string] $Suffix = '-md',

    [switch] $Recurse,

    [string[]] $Include,

    [switch] $Flat,

    [switch] $Force,

    [string] $MarkItDownPath,

    [switch] $UsePlugins,

    [switch] $KeepDataUris,

    [switch] $ExtractImages,

    [string] $AssetsFolder,

    [int] $MinImagePixels = 0,

    [switch] $NoImageStitching
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# markitdown truncates data URIs unless asked not to, so there would be nothing to extract.
if ($ExtractImages -and -not $KeepDataUris) {
    $KeepDataUris = $true
    Write-Verbose '-ExtractImages implies --keep-data-uris'
}

if ($AssetsFolder -and -not $ExtractImages) {
    Write-Warning '-AssetsFolder has no effect without -ExtractImages.'
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Resolve-PythonForExecutable {
    <#
        Finds the python.exe that belongs to a markitdown executable, so PDF image
        extraction runs against the same environment that did the conversion.
    #>
    param([string] $ExecutablePath)

    if ($ExecutablePath) {
        $exeDir = Split-Path -Path $ExecutablePath -Parent
        $candidates = @(
            (Join-Path -Path $exeDir -ChildPath 'python.exe'),                              # venv\Scripts\
            (Join-Path -Path (Split-Path -Path $exeDir -Parent) -ChildPath 'python.exe')    # ...\python.exe
        )
        foreach ($candidate in $candidates) {
            if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
        }
    }

    $onPath = Get-Command -Name 'python' -CommandType Application -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($onPath) { return $onPath.Source }
    return $null
}

function Resolve-MarkItDownCommand {
    <#
        Returns the executable to run plus any leading arguments it needs
        (empty for markitdown.exe, '-m markitdown' for the python fallback),
        along with the python.exe of the same environment.
    #>
    param([string] $Explicit)

    if ($Explicit) {
        if (Test-Path -LiteralPath $Explicit -PathType Leaf) {
            $resolved = (Resolve-Path -LiteralPath $Explicit).Path
            return [pscustomobject]@{
                File   = $resolved
                Prefix = @()
                Label  = $Explicit
                Python = Resolve-PythonForExecutable -ExecutablePath $resolved
            }
        }
        $cmd = Get-Command -Name $Explicit -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($cmd) {
            return [pscustomobject]@{
                File   = $cmd.Source
                Prefix = @()
                Label  = $cmd.Source
                Python = Resolve-PythonForExecutable -ExecutablePath $cmd.Source
            }
        }
        throw "markitdown was not found at -MarkItDownPath '$Explicit'."
    }

    # A virtual environment next to this script, or next to its parent (repo checkout layout).
    if ($PSScriptRoot) {
        $roots = @($PSScriptRoot, (Split-Path -Path $PSScriptRoot -Parent))
        foreach ($root in $roots) {
            if (-not $root) { continue }
            $candidate = Join-Path -Path $root -ChildPath '.venv\Scripts\markitdown.exe'
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                return [pscustomobject]@{
                    File   = $candidate
                    Prefix = @()
                    Label  = $candidate
                    Python = Resolve-PythonForExecutable -ExecutablePath $candidate
                }
            }
        }
    }

    $onPath = Get-Command -Name 'markitdown' -CommandType Application -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($onPath) {
        return [pscustomobject]@{
            File   = $onPath.Source
            Prefix = @()
            Label  = $onPath.Source
            Python = Resolve-PythonForExecutable -ExecutablePath $onPath.Source
        }
    }

    foreach ($py in @('python', 'py')) {
        $found = Get-Command -Name $py -CommandType Application -ErrorAction SilentlyContinue |
            Select-Object -First 1
        if ($found) {
            return [pscustomobject]@{
                File   = $found.Source
                Prefix = @('-m', 'markitdown')
                Label  = "$($found.Source) -m markitdown"
                Python = $found.Source
            }
        }
    }

    throw 'Could not find markitdown. Install it with: pip install "markitdown[all]", or pass -MarkItDownPath.'
}

function ConvertTo-CommandLine {
    <# Quotes arguments for the Windows command line (paths may contain spaces). #>
    param([string[]] $Arguments)

    $parts = foreach ($argument in $Arguments) {
        # Backslashes immediately before the closing quote must be doubled, or they escape it.
        '"' + ($argument -replace '(\\+)$', '$1$1') + '"'
    }
    return ($parts -join ' ')
}

function Invoke-MarkItDown {
    <# Runs one conversion and returns the exit code plus captured output. #>
    param(
        [string]   $File,
        [string[]] $Arguments
    )

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = $File
    $psi.Arguments              = ConvertTo-CommandLine -Arguments $Arguments
    $psi.UseShellExecute        = $false
    $psi.CreateNoWindow         = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $psi.StandardErrorEncoding  = [System.Text.Encoding]::UTF8

    $proc = [System.Diagnostics.Process]::Start($psi)
    # Drain stderr asynchronously so neither pipe can fill up and deadlock the child.
    $errTask = $proc.StandardError.ReadToEndAsync()
    $stdout = $proc.StandardOutput.ReadToEnd()
    $proc.WaitForExit()

    return [pscustomobject]@{
        ExitCode = $proc.ExitCode
        StdOut   = $stdout
        StdErr   = $errTask.Result
    }
}

function Get-ImageExtension {
    <# Maps a data-URI MIME type onto a sensible file extension. #>
    param([string] $MimeType)

    $known = @{
        'image/png'     = '.png'
        'image/jpeg'    = '.jpg'
        'image/jpg'     = '.jpg'
        'image/gif'     = '.gif'
        'image/bmp'     = '.bmp'
        'image/webp'    = '.webp'
        'image/tiff'    = '.tif'
        'image/svg+xml' = '.svg'
        'image/x-emf'   = '.emf'
        'image/emf'     = '.emf'
        'image/x-wmf'   = '.wmf'
        'image/wmf'     = '.wmf'
        'image/x-icon'  = '.ico'
    }

    $key = $MimeType.ToLowerInvariant()
    if ($known.ContainsKey($key)) { return $known[$key] }

    $subtype = ($key -split '/')[-1] -replace '^x-', '' -replace '[^a-z0-9]', ''
    if ($subtype) { return ".$subtype" }
    return '.bin'
}

function New-AssetState {
    <# Per-document bookkeeping for the assets folder: naming, numbering and dedupe. #>
    param([string] $Dir, [string] $NamePrefix, [hashtable] $HashIndex)

    return [pscustomobject]@{
        Dir       = $Dir
        Prefix    = $NamePrefix
        HashIndex = $HashIndex
        Written   = 0
        Sha       = [System.Security.Cryptography.SHA256]::Create()
    }
}

function Clear-DocumentAsset {
    <# Removes assets a previous run of this document wrote, so -Force leaves no orphans. #>
    param($State)

    if (-not (Test-Path -LiteralPath $State.Dir -PathType Container)) { return }

    $stalePattern = '^' + [regex]::Escape($State.Prefix) + '-\d{3}\.'
    Get-ChildItem -LiteralPath $State.Dir -File |
        Where-Object { $_.Name -match $stalePattern } |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

function Add-Asset {
    <#
        Writes one image into the assets folder and returns its path. Bytes that were
        already written (same SHA-256) are reused instead of duplicated.
    #>
    param($State, [byte[]] $Bytes, [string] $Extension)

    $hash = [System.BitConverter]::ToString($State.Sha.ComputeHash($Bytes)).Replace('-', '')
    if ($State.HashIndex.ContainsKey($hash)) {
        return $State.HashIndex[$hash]
    }

    if (-not (Test-Path -LiteralPath $State.Dir -PathType Container)) {
        New-Item -ItemType Directory -Path $State.Dir -Force | Out-Null
    }

    $State.Written++
    $fileName = '{0}-{1:d3}{2}' -f $State.Prefix, $State.Written, $Extension
    $imagePath = Join-Path -Path $State.Dir -ChildPath $fileName
    [System.IO.File]::WriteAllBytes($imagePath, $Bytes)
    $State.HashIndex[$hash] = $imagePath

    return $imagePath
}

function Get-MarkdownLink {
    <# Relative, forward-slashed and %20-escaped, so any Markdown viewer can follow it. #>
    param([string] $FromFile, [string] $ToFile)

    $from = New-Object System.Uri($FromFile)
    return $from.MakeRelativeUri((New-Object System.Uri($ToFile))).ToString()
}

function Export-EmbeddedImage {
    <#
        Decodes every base64 image in a Markdown file, writes it into the assets folder,
        and replaces the data URI with a relative link. Returns how many files were written.
    #>
    param([string] $MarkdownPath, $State)

    $text = [System.IO.File]::ReadAllText($MarkdownPath, [System.Text.Encoding]::UTF8)
    $pattern = 'data:(?<mime>image/[A-Za-z0-9.+-]+);base64,(?<data>[A-Za-z0-9+/]+={0,2})'
    $found = [regex]::Matches($text, $pattern)
    if ($found.Count -eq 0) { return 0 }

    $before = $State.Written
    $builder = New-Object System.Text.StringBuilder
    $position = 0

    foreach ($match in $found) {
        [void] $builder.Append($text.Substring($position, $match.Index - $position))
        $position = $match.Index + $match.Length

        try {
            $bytes = [System.Convert]::FromBase64String($match.Groups['data'].Value)
        } catch {
            # Not valid base64 after all - leave the original text alone.
            [void] $builder.Append($match.Value)
            continue
        }

        $imagePath = Add-Asset -State $State -Bytes $bytes `
                               -Extension (Get-ImageExtension -MimeType $match.Groups['mime'].Value)
        [void] $builder.Append((Get-MarkdownLink -FromFile $MarkdownPath -ToFile $imagePath))
    }

    [void] $builder.Append($text.Substring($position))
    [System.IO.File]::WriteAllText($MarkdownPath, $builder.ToString(), (New-Object System.Text.UTF8Encoding($false)))

    return ($State.Written - $before)
}

# Run inside markitdown's own Python: pdfminer.six ships with markitdown[pdf], and Pillow comes
# with pdfplumber. pdfminer's own ImageWriter is not used - it writes uncompressed BMPs and gets
# indexed and CMYK images wrong - so the image streams are decoded here and saved as PNG/JPEG.
$PdfImageHelperSource = @'
import io
import json
import os
import sys

from pdfminer.high_level import extract_pages
from pdfminer.layout import LTFigure, LTImage
from pdfminer.pdftypes import PDFStream, resolve1
from pdfminer.psparser import PSLiteral

try:
    from PIL import Image, ImageChops
except ImportError:
    Image = None

DCT = ("DCTDecode", "DCT")
JPX = ("JPXDecode",)
JBIG2 = ("JBIG2Decode",)

GRAY_NAMES = ("DeviceGray", "CalGray", "G")
RGB_NAMES = ("DeviceRGB", "CalRGB", "RGB", "Lab")
CMYK_NAMES = ("DeviceCMYK", "CMYK")

# How close two image edges must be, in PDF units, to count as one tiled picture.
TILE_TOLERANCE = 1.0
# Refuse to build a stitched canvas larger than this, to stay out of memory trouble.
MAX_CANVAS_PIXELS = 80_000_000
# Above this many images on one page, skip the quadratic tile grouping entirely.
MAX_TILE_RECORDS = 1500


def walk(container):
    """PDF images can be nested inside figures, so recurse."""
    for item in container:
        if isinstance(item, LTImage):
            yield item
        elif isinstance(item, LTFigure):
            for nested in walk(item):
                yield nested


def literal_name(obj):
    obj = resolve1(obj)
    if isinstance(obj, PSLiteral):
        return obj.name
    if isinstance(obj, bytes):
        return obj.decode("latin-1")
    return str(obj)


def colorspace_info(colorspace):
    """Return (components_per_pixel, palette) where palette is (base_components, bytes)."""
    colorspace = resolve1(colorspace)
    if isinstance(colorspace, list) and len(colorspace) == 1:
        colorspace = resolve1(colorspace[0])

    if not isinstance(colorspace, list):
        name = literal_name(colorspace)
        if name in RGB_NAMES:
            return 3, None
        if name in CMYK_NAMES:
            return 4, None
        return 1, None

    family = literal_name(colorspace[0])

    if family in ("Indexed", "I"):
        base_components, _ = colorspace_info(colorspace[1])
        lookup = resolve1(colorspace[3])
        if isinstance(lookup, PDFStream):
            lookup = lookup.get_data()
        elif isinstance(lookup, str):
            lookup = lookup.encode("latin-1")
        return 1, (base_components, bytes(lookup))

    if family == "ICCBased":
        stream = resolve1(colorspace[1])
        try:
            return int(resolve1(stream["N"])), None
        except Exception:
            return 3, None

    if family == "DeviceN":
        try:
            return len(resolve1(colorspace[1])), None
        except Exception:
            return 1, None

    if family == "Separation":
        return 1, None
    if family in ("CalRGB", "Lab"):
        return 3, None
    if family == "CalGray":
        return 1, None

    name = literal_name(colorspace[0])
    if name in RGB_NAMES:
        return 3, None
    if name in CMYK_NAMES:
        return 4, None
    return 1, None


def pad_rows(data, stride, height):
    rows = [data[y * stride:(y + 1) * stride].ljust(stride, b"\x00") for y in range(height)]
    return b"".join(rows)


def unpack_samples(data, width, height, components, bits, scale_values):
    """Normalise any bit depth to one byte per sample, honouring PDF row padding."""
    if bits == 8:
        return pad_rows(data, width * components, height)

    if bits == 16:
        stride = width * components * 2
        out = bytearray()
        for y in range(height):
            row = data[y * stride:(y + 1) * stride].ljust(stride, b"\x00")
            out += row[0::2]  # keep the high byte of each sample
        return bytes(out)

    # 1, 2 or 4 bits per component: rows are padded to a byte boundary.
    stride = (width * components * bits + 7) // 8
    max_value = (1 << bits) - 1
    factor = 255 // max_value
    out = bytearray()
    for y in range(height):
        row = data[y * stride:(y + 1) * stride].ljust(stride, b"\x00")
        packed = int.from_bytes(row, "big")
        total_bits = stride * 8
        for i in range(width * components):
            value = (packed >> (total_bits - (i + 1) * bits)) & max_value
            out.append(value * factor if scale_values else value)
    return bytes(out)


def samples_to_image(data, width, height, components, bits, palette):
    if palette is not None:
        indices = unpack_samples(data, width, height, 1, bits, scale_values=False)
        image = Image.frombytes("P", (width, height), indices)
        base_components, lookup = palette
        if base_components == 3:
            image.putpalette(lookup[: 256 * 3].ljust(768, b"\x00"))
        elif base_components == 1:
            expanded = bytearray()
            for value in lookup[:256]:
                expanded += bytes((value, value, value))
            image.putpalette(bytes(expanded).ljust(768, b"\x00"))
        elif base_components == 4:
            expanded = bytearray()
            for i in range(0, min(len(lookup), 256 * 4), 4):
                c, m, y, k = lookup[i:i + 4]
                expanded += bytes((
                    int(255 - min(255, c + k)),
                    int(255 - min(255, m + k)),
                    int(255 - min(255, y + k)),
                ))
            image.putpalette(bytes(expanded).ljust(768, b"\x00"))
        return image.convert("RGB")

    if bits == 1 and components == 1:
        packed = pad_rows(data, (width + 7) // 8, height)
        return Image.frombytes("1", (width, height), packed).convert("L")

    samples = unpack_samples(data, width, height, components, bits, scale_values=True)
    mode = {1: "L", 3: "RGB", 4: "CMYK"}.get(components)
    if mode is None:
        return None
    image = Image.frombytes(mode, (width, height), samples)
    return image.convert("RGB") if mode == "CMYK" else image


def decode_array(stream):
    try:
        return resolve1(stream.get_any(("D", "Decode")))
    except Exception:
        return None


def stencil_to_image(stream, width, height):
    """An /ImageMask paints a single colour through a 1-bit stencil.

    Sample 0 paints and 1 leaves the page alone (reversed by a Decode of [1 0]).
    The paint colour lives in the content stream's graphics state, which is not
    available here, so the shape is rendered black on a transparent background -
    that keeps the glyph readable instead of a solid black rectangle.
    """
    packed = pad_rows(stream.get_data(), (width + 7) // 8, height)
    bitmap = Image.frombytes("1", (width, height), packed).convert("L")

    # In PIL, 0 is black; those are exactly the painted samples under Decode [0 1].
    alpha = ImageChops.invert(bitmap)

    decode = decode_array(stream)
    if decode and len(decode) >= 2 and decode[0] == 1:
        alpha = ImageChops.invert(alpha)

    picture = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    picture.putalpha(alpha)
    return picture


def apply_decode_array(image, stream):
    """A Decode of [1 0] means the samples are stored inverted."""
    decode = decode_array(stream)
    if decode and len(decode) >= 2 and decode[0] == 1 and decode[1] == 0:
        if image.mode in ("L", "RGB"):
            return ImageChops.invert(image)
    return image


def get_soft_mask(stream):
    """SMask holds the alpha channel; without it transparent logos come out black."""
    try:
        smask = resolve1(stream.get_any(("SMask",)))
    except Exception:
        return None
    if not isinstance(smask, PDFStream):
        return None
    try:
        width = int(resolve1(smask.get_any(("W", "Width"))))
        height = int(resolve1(smask.get_any(("H", "Height"))))
        bits = int(resolve1(smask.get_any(("BPC", "BitsPerComponent"))) or 8)
        filters = smask.get_filters()
        if filters and literal_name(filters[-1][0]) in DCT:
            mask = Image.open(io.BytesIO(smask.get_data())).convert("L")
        else:
            mask = samples_to_image(smask.get_data(), width, height, 1, bits, None)
        return mask.convert("L") if mask is not None else None
    except Exception:
        return None


def is_stencil(stream):
    try:
        return bool(resolve1(stream.get_any(("IM", "ImageMask"))))
    except Exception:
        return False


def decode_image(image):
    """Return (PIL image or None, raw_bytes or None, extension)."""
    stream = image.stream
    filters = stream.get_filters()
    last = literal_name(filters[-1][0]) if filters else ""

    if last in JBIG2:
        return None, None, None  # nothing in the stack can decode JBIG2

    width, height = image.srcsize

    if is_stencil(stream):
        return stencil_to_image(stream, width, height), None, ".png"

    bits = image.bits or 8
    components, palette = colorspace_info(stream.get_any(("CS", "ColorSpace")))
    mask = get_soft_mask(stream)

    if last in DCT:
        data = stream.get_data()
        if components == 4 or mask is not None:
            # Pillow already applies the Adobe APP14 transform, so no manual inversion here.
            picture = Image.open(io.BytesIO(data)).convert("RGB")
            if mask is not None:
                picture.putalpha(mask.resize(picture.size))
            return picture, None, ".png"
        return None, data, ".jpg"  # already a valid JPEG file - pass it straight through

    if last in JPX:
        data = stream.get_data()
        try:
            picture = Image.open(io.BytesIO(data))
            picture.load()
            return picture.convert("RGB"), None, ".png"
        except Exception:
            return None, data, ".jp2"

    picture = samples_to_image(stream.get_data(), width, height, components, bits, palette)
    if picture is None:
        return None, None, None

    # pdfminer's CCITT decoder already normalises BlackIs1, so no polarity fix belongs here.
    picture = apply_decode_array(picture, stream)
    if mask is not None:
        picture = picture.convert("RGB")
        picture.putalpha(mask.resize(picture.size))
    return picture, None, ".png"


def edges_match(a_low, a_high, b_low, b_high):
    return abs(a_low - b_low) <= TILE_TOLERANCE and abs(a_high - b_high) <= TILE_TOLERANCE


def are_tiles(a, b):
    """True when two placements sit edge to edge and line up on the other axis."""
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b

    if edges_match(ay0, ay1, by0, by1):  # side by side
        if abs(ax1 - bx0) <= TILE_TOLERANCE or abs(bx1 - ax0) <= TILE_TOLERANCE:
            return True

    if edges_match(ax0, ax1, bx0, bx1):  # stacked
        if abs(ay1 - by0) <= TILE_TOLERANCE or abs(by1 - ay0) <= TILE_TOLERANCE:
            return True

    return False


def group_tiles(records):
    """Union-find over placements, so a whole grid of tiles ends up in one group."""
    # The pairwise scan below is quadratic; a page carrying thousands of fragments is
    # pathological, so leave those alone rather than spending minutes on them.
    if len(records) > MAX_TILE_RECORDS:
        return [[record] for record in records]

    parent = list(range(len(records)))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def union(i, j):
        ri, rj = find(i), find(j)
        if ri != rj:
            parent[rj] = ri

    for i in range(len(records)):
        for j in range(i + 1, len(records)):
            if are_tiles(records[i]["bbox"], records[j]["bbox"]):
                union(i, j)

    groups = {}
    for index in range(len(records)):
        groups.setdefault(find(index), []).append(records[index])
    return list(groups.values())


def ensure_picture(record):
    """Tiles have to be decoded before they can be pasted onto a canvas."""
    if record["picture"] is not None:
        return record["picture"]
    if record["raw"] is None:
        return None
    try:
        picture = Image.open(io.BytesIO(record["raw"]))
        picture.load()
        record["picture"] = picture
        return picture
    except Exception:
        return None


def stitch(members):
    """Paste tiles onto one canvas using their positions on the page."""
    for member in members:
        if ensure_picture(member) is None:
            return None

    x0 = min(m["bbox"][0] for m in members)
    y0 = min(m["bbox"][1] for m in members)
    x1 = max(m["bbox"][2] for m in members)
    y1 = max(m["bbox"][3] for m in members)
    if x1 - x0 <= 0 or y1 - y0 <= 0:
        return None

    # Keep the sharpest tile's resolution for the whole picture.
    scale_x = max(m["picture"].width / max(m["bbox"][2] - m["bbox"][0], 1e-6) for m in members)
    scale_y = max(m["picture"].height / max(m["bbox"][3] - m["bbox"][1], 1e-6) for m in members)

    canvas_width = max(1, int(round((x1 - x0) * scale_x)))
    canvas_height = max(1, int(round((y1 - y0) * scale_y)))
    if canvas_width * canvas_height > MAX_CANVAS_PIXELS:
        return None

    transparent = any(m["picture"].mode in ("RGBA", "LA", "P") for m in members)
    canvas = Image.new("RGBA" if transparent else "RGB",
                       (canvas_width, canvas_height),
                       (0, 0, 0, 0) if transparent else (255, 255, 255))

    for member in members:
        mx0, my0, mx1, my1 = member["bbox"]
        target_width = max(1, int(round((mx1 - mx0) * scale_x)))
        target_height = max(1, int(round((my1 - my0) * scale_y)))

        piece = member["picture"]
        if piece.size != (target_width, target_height):
            piece = piece.resize((target_width, target_height), Image.LANCZOS)

        left = int(round((mx0 - x0) * scale_x))
        top = int(round((y1 - my1) * scale_y))  # PDF y grows upwards, images downwards

        if piece.mode in ("RGBA", "LA"):
            canvas.paste(piece, (left, top), piece)
        else:
            canvas.paste(piece, (left, top))

    return canvas


def main():
    pdf_path, out_dir, min_pixels = sys.argv[1], sys.argv[2], int(sys.argv[3])
    stitch_tiles = len(sys.argv) < 5 or sys.argv[4] != "0"
    os.makedirs(out_dir, exist_ok=True)

    if Image is None:
        json.dump({"images": [], "skipped": 0, "error": "Pillow is not installed"}, sys.stdout)
        return

    exported = []
    skipped = 0
    stitched = 0
    counter = 0

    for page_number, page in enumerate(extract_pages(pdf_path), start=1):
        records = []
        for image in walk(page):
            try:
                picture, raw, extension = decode_image(image)
            except Exception:
                picture, raw, extension = None, None, None

            if extension is None:
                skipped += 1
                continue

            records.append({
                "bbox": tuple(float(v) for v in image.bbox),
                "pixels": image.srcsize[0] * image.srcsize[1],
                "picture": picture,
                "raw": raw,
                "ext": extension,
            })

        groups = group_tiles(records) if stitch_tiles else [[r] for r in records]

        for members in groups:
            picture = None
            raw = None
            extension = ".png"

            if len(members) > 1:
                picture = stitch(members)
                if picture is not None:
                    stitched += len(members)
                else:
                    # Could not combine them; fall back to writing each tile.
                    for member in members:
                        counter = save_record(member, out_dir, counter, exported,
                                              page_number, min_pixels)
                    continue
            else:
                member = members[0]
                if min_pixels and member["pixels"] < min_pixels:
                    skipped += 1
                    continue
                picture, raw, extension = member["picture"], member["raw"], member["ext"]

            if picture is not None and min_pixels:
                if picture.width * picture.height < min_pixels:
                    skipped += 1
                    continue

            counter += 1
            path = os.path.join(out_dir, "img-%04d%s" % (counter, extension))
            try:
                if raw is not None:
                    with open(path, "wb") as handle:
                        handle.write(raw)
                else:
                    picture.save(path)
            except Exception:
                skipped += 1
                counter -= 1
                continue

            exported.append({"page": page_number, "path": path})

    json.dump({"images": exported, "skipped": skipped, "stitched": stitched}, sys.stdout)


def save_record(member, out_dir, counter, exported, page_number, min_pixels):
    if min_pixels and member["pixels"] < min_pixels:
        return counter
    counter += 1
    path = os.path.join(out_dir, "img-%04d%s" % (counter, member["ext"]))
    try:
        if member["raw"] is not None:
            with open(path, "wb") as handle:
                handle.write(member["raw"])
        else:
            member["picture"].save(path)
    except Exception:
        return counter - 1
    exported.append({"page": page_number, "path": path})
    return counter


main()
'@

function Export-PdfImage {
    <#
        markitdown converts PDFs with pdfminer/pdfplumber text extraction only, so a PDF
        never produces data URIs. This pulls the image XObjects straight out of the PDF and
        appends them to the Markdown, grouped by page.
    #>
    param(
        [string] $PdfPath,
        [string] $MarkdownPath,
        $State,
        [string] $PythonExe,
        [string] $HelperScript,
        [int]    $MinPixels,
        [bool]   $Stitch
    )

    $tempDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) `
                         -ChildPath ('markitdown-pdfimg-' + [System.Guid]::NewGuid().ToString('N'))

    try {
        $stitchFlag = if ($Stitch) { '1' } else { '0' }
        $run = Invoke-MarkItDown -File $PythonExe `
                                 -Arguments @($HelperScript, $PdfPath, $tempDir, "$MinPixels", $stitchFlag)

        if ($run.ExitCode -ne 0) {
            Write-Warning ('image extraction failed for {0}: {1}' -f
                           (Split-Path -Path $PdfPath -Leaf),
                           (Get-FailureReason -StdErr $run.StdErr -StdOut ''))
            return 0
        }

        $report = $run.StdOut | ConvertFrom-Json

        if ($report.PSObject.Properties.Name -contains 'error') {
            Write-Warning ('image extraction unavailable: {0}' -f $report.error)
            return 0
        }

        $images = @($report.images)
        if ($images.Count -eq 0) { return 0 }

        $before = $State.Written
        $sections = New-Object System.Collections.Generic.List[string]
        $currentPage = -1
        $indexOnPage = 0

        foreach ($image in $images) {
            if (-not (Test-Path -LiteralPath $image.path -PathType Leaf)) { continue }

            $bytes = [System.IO.File]::ReadAllBytes($image.path)
            $extension = [System.IO.Path]::GetExtension($image.path)
            if (-not $extension) { $extension = '.img' }

            $imagePath = Add-Asset -State $State -Bytes $bytes -Extension $extension

            if ($image.page -ne $currentPage) {
                $currentPage = $image.page
                $indexOnPage = 0
                $sections.Add('')
                $sections.Add("### Page $currentPage")
                $sections.Add('')
            }
            $indexOnPage++

            $link = Get-MarkdownLink -FromFile $MarkdownPath -ToFile $imagePath
            $sections.Add("![Page $currentPage image $indexOnPage]($link)")
            $sections.Add('')
        }

        if ($sections.Count -eq 0) { return 0 }

        # The extracted text carries no page markers, so the images go in their own section
        # rather than being guessed into place mid-document.
        $appendix = "`n`n## Extracted images`n" + ($sections -join "`n") + "`n"
        [System.IO.File]::AppendAllText($MarkdownPath, $appendix, (New-Object System.Text.UTF8Encoding($false)))

        return ($State.Written - $before)
    } finally {
        if (Test-Path -LiteralPath $tempDir -PathType Container) {
            Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-FailureReason {
    <# Boils markitdown's output down to one actionable line. #>
    param([string] $StdErr, [string] $StdOut)

    $text = ("$StdErr`n$StdOut").Trim()
    if (-not $text) { return 'markitdown exited with a non-zero status' }

    $missing = [regex]::Match($text, "pip install 'markitdown\[([a-z0-9_-]+)\]'")
    if ($missing.Success) {
        return ('missing dependency - run: pip install "markitdown[{0}]"' -f $missing.Groups[1].Value)
    }

    $lines = @($text -split "`r?`n" | Where-Object { $_.Trim() })
    if ($lines.Count -eq 0) { return 'markitdown exited with a non-zero status' }
    return $lines[-1].Trim()
}

# ---------------------------------------------------------------------------
# Resolve folders
# ---------------------------------------------------------------------------

if (-not (Test-Path -LiteralPath $InputFolder -PathType Container)) {
    throw "Input folder not found: $InputFolder"
}
$srcRoot = (Resolve-Path -LiteralPath $InputFolder).Path.TrimEnd('\')

if ($OutputFolder) {
    $outRoot = [System.IO.Path]::GetFullPath($OutputFolder).TrimEnd('\')
} else {
    # Sibling of the original: C:\data\docs -> C:\data\docs-md
    $outRoot = Join-Path -Path (Split-Path -Path $srcRoot -Parent) `
                         -ChildPath ((Split-Path -Path $srcRoot -Leaf) + $Suffix)
    $outRoot = $outRoot.TrimEnd('\')
}

if ($outRoot -eq $srcRoot) {
    throw 'The output folder must be different from the input folder.'
}

$markitdown = Resolve-MarkItDownCommand -Explicit $MarkItDownPath
Write-Verbose "Using markitdown: $($markitdown.Label)"

# ---------------------------------------------------------------------------
# Enumerate the files to convert
# ---------------------------------------------------------------------------

$gci = @{ LiteralPath = $srcRoot; File = $true }
if ($Recurse) { $gci['Recurse'] = $true }

$outPrefix = $outRoot + '\'
$files = @(
    Get-ChildItem @gci |
        Where-Object {
            # Never re-convert our own output if it sits inside the source tree.
            -not $_.FullName.StartsWith($outPrefix, [System.StringComparison]::OrdinalIgnoreCase)
        }
)

if ($Include) {
    $wanted = @($Include | ForEach-Object {
        $ext = $_.Trim().ToLowerInvariant()
        if ($ext -and -not $ext.StartsWith('.')) { $ext = ".$ext" }
        $ext
    })
    $files = @($files | Where-Object { $wanted -contains $_.Extension.ToLowerInvariant() })
}

if ($files.Count -eq 0) {
    Write-Warning "No matching files found in $srcRoot"
    return
}

# ---------------------------------------------------------------------------
# Plan output names up front, so re-runs always land on the same file names
# ---------------------------------------------------------------------------

$plan = @(
    foreach ($file in $files) {
        if ($Recurse -and -not $Flat) {
            $relativeDir = $file.DirectoryName.Substring($srcRoot.Length).TrimStart('\')
            if ($relativeDir) {
                $targetDir = Join-Path -Path $outRoot -ChildPath $relativeDir
            } else {
                $targetDir = $outRoot
            }
        } else {
            $targetDir = $outRoot
        }

        [pscustomobject]@{
            Source    = $file
            TargetDir = $targetDir
            Output    = $null
        }
    }
)

# report.pdf and report.docx in the same target folder would both want report.md,
# so when a base name is not unique every one of them keeps its extension.
$stemCounts = @{}
foreach ($item in $plan) {
    $key = "$($item.TargetDir)|$($item.Source.BaseName)".ToLowerInvariant()
    if ($stemCounts.ContainsKey($key)) { $stemCounts[$key]++ } else { $stemCounts[$key] = 1 }
}

$usedNames = @{}
foreach ($item in $plan) {
    $key = "$($item.TargetDir)|$($item.Source.BaseName)".ToLowerInvariant()
    if ($stemCounts[$key] -gt 1) {
        $name = "$($item.Source.Name).md"
    } else {
        $name = "$($item.Source.BaseName).md"
    }

    # -Flat can pull identically named files from different subfolders into one folder.
    $nameKey = "$($item.TargetDir)|$name".ToLowerInvariant()
    if ($usedNames.ContainsKey($nameKey)) {
        $seen = $usedNames[$nameKey]
        $usedNames[$nameKey] = $seen + 1
        $name = "$([System.IO.Path]::GetFileNameWithoutExtension($name)).$seen.md"
    } else {
        $usedNames[$nameKey] = 1
    }

    $item.Output = Join-Path -Path $item.TargetDir -ChildPath $name
}

# ---------------------------------------------------------------------------
# Convert
# ---------------------------------------------------------------------------

if (-not (Test-Path -LiteralPath $outRoot)) {
    if ($PSCmdlet.ShouldProcess($outRoot, 'Create output folder')) {
        New-Item -ItemType Directory -Path $outRoot -Force | Out-Null
    }
}

Write-Host "markitdown : $($markitdown.Label)"
Write-Host "source     : $srcRoot"
Write-Host "output     : $outRoot"
Write-Host "files      : $($plan.Count)`n"

$converted   = 0
$skipped     = 0
$imagesTotal = 0
$failures    = New-Object System.Collections.Generic.List[object]
$index       = 0

# Shared -AssetsFolder mode dedupes images across every document; per-document mode
# starts a fresh index for each file.
$sharedHashIndex = @{}

# PDF image extraction needs a python that can import pdfminer, which markitdown[pdf] provides.
$pdfHelperScript = $null
$pdfHelperDir = $null
if ($ExtractImages -and ($plan | Where-Object { $_.Source.Extension -ieq '.pdf' })) {
    if ($markitdown.Python) {
        # Python puts the script's own folder first on sys.path, so the helper gets a private
        # directory: dropping it straight into %TEMP% lets any stray file there (a leftover
        # inspect.py, say) shadow a standard library module and break the import.
        $pdfHelperDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) `
                                  -ChildPath "markitdown-pdfimages-$PID"
        New-Item -ItemType Directory -Path $pdfHelperDir -Force | Out-Null
        $pdfHelperScript = Join-Path -Path $pdfHelperDir -ChildPath 'extract_pdf_images.py'
        Set-Content -LiteralPath $pdfHelperScript -Value $PdfImageHelperSource -Encoding UTF8
    } else {
        Write-Warning 'No python.exe found for PDF image extraction; PDFs will be text only.'
    }
}

foreach ($item in $plan) {
    $index++
    $source = $item.Source
    $output = $item.Output

    Write-Progress -Activity 'Converting to Markdown' `
                   -Status "$index of $($plan.Count): $($source.Name)" `
                   -PercentComplete (($index / $plan.Count) * 100)

    if ((Test-Path -LiteralPath $output -PathType Leaf) -and -not $Force) {
        Write-Host ('exists  {0}' -f $source.Name) -ForegroundColor DarkGray
        $skipped++
        continue
    }

    if (-not $PSCmdlet.ShouldProcess($source.FullName, "convert to $output")) { continue }

    $targetDir = Split-Path -Path $output -Parent
    if (-not (Test-Path -LiteralPath $targetDir -PathType Container)) {
        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    }

    $arguments = @($markitdown.Prefix) + @($source.FullName, '-o', $output)
    if ($UsePlugins)   { $arguments += '-p' }
    if ($KeepDataUris) { $arguments += '--keep-data-uris' }

    $result = Invoke-MarkItDown -File $markitdown.File -Arguments $arguments

    if ($result.ExitCode -eq 0 -and (Test-Path -LiteralPath $output -PathType Leaf)) {
        $imageNote = ''
        if ($ExtractImages) {
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($output)
            if ($AssetsFolder) {
                $assetsDir = Join-Path -Path $outRoot -ChildPath $AssetsFolder
                $hashIndex = $sharedHashIndex
            } else {
                $assetsDir = Join-Path -Path $targetDir -ChildPath "$baseName.assets"
                $hashIndex = @{}
            }

            $assetState = New-AssetState -Dir $assetsDir -NamePrefix $baseName -HashIndex $hashIndex
            try {
                Clear-DocumentAsset -State $assetState

                # Office/HTML output carries images inline as data URIs...
                [void] (Export-EmbeddedImage -MarkdownPath $output -State $assetState)

                # ...but PDF output is text only, so those images come from the PDF itself.
                if ($source.Extension -ieq '.pdf' -and $pdfHelperScript) {
                    [void] (Export-PdfImage -PdfPath $source.FullName `
                                            -MarkdownPath $output `
                                            -State $assetState `
                                            -PythonExe $markitdown.Python `
                                            -HelperScript $pdfHelperScript `
                                            -MinPixels $MinImagePixels `
                                            -Stitch (-not $NoImageStitching))
                }

                $imageCount = $assetState.Written
            } finally {
                $assetState.Sha.Dispose()
            }

            $imagesTotal += $imageCount
            if ($imageCount -gt 0) {
                $imageNote = ' (+{0} image{1})' -f $imageCount, $(if ($imageCount -eq 1) { '' } else { 's' })
            }
        }

        Write-Host ('ok      {0} -> {1}{2}' -f $source.Name, (Split-Path -Path $output -Leaf), $imageNote) -ForegroundColor Green
        $converted++
    } else {
        $reason = Get-FailureReason -StdErr $result.StdErr -StdOut $result.StdOut
        Write-Host ('FAILED  {0}: {1}' -f $source.Name, $reason) -ForegroundColor Red
        $failures.Add([pscustomobject]@{
            File   = $source.FullName
            Reason = $reason
            Detail = ("$($result.StdErr)`n$($result.StdOut)").Trim()
        })
        # A partial file from a crashed run would look like a real conversion on the next pass.
        if (Test-Path -LiteralPath $output -PathType Leaf) {
            Remove-Item -LiteralPath $output -Force -ErrorAction SilentlyContinue
        }
    }
}

Write-Progress -Activity 'Converting to Markdown' -Completed

if ($pdfHelperDir -and (Test-Path -LiteralPath $pdfHelperDir -PathType Container)) {
    Remove-Item -LiteralPath $pdfHelperDir -Recurse -Force -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

$logPath = $null
if ($failures.Count -gt 0 -and (Test-Path -LiteralPath $outRoot -PathType Container)) {
    $logPath = Join-Path -Path $outRoot -ChildPath '_markitdown-errors.log'
    $lines = foreach ($failure in $failures) {
        "=== $($failure.File)"
        $failure.Reason
        $failure.Detail
        ''
    }
    Set-Content -LiteralPath $logPath -Value $lines -Encoding UTF8
}

Write-Host ''
$summary = '{0} converted, {1} skipped, {2} failed' -f $converted, $skipped, $failures.Count
if ($ExtractImages) {
    $summary += ', {0} image{1} extracted' -f $imagesTotal, $(if ($imagesTotal -eq 1) { '' } else { 's' })
}
Write-Host $summary
Write-Host ("output: $outRoot")
if ($logPath) { Write-Host "errors: $logPath" -ForegroundColor Yellow }

[pscustomobject]@{
    InputFolder     = $srcRoot
    OutputFolder    = $outRoot
    Converted       = $converted
    Skipped         = $skipped
    Failed          = $failures.Count
    ImagesExtracted = $imagesTotal
    Failures        = $failures.ToArray()
    ErrorLog        = $logPath
}
