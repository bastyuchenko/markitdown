# Convert-FolderToMarkdown.ps1

Batch-converts every file in a folder to Markdown by running the compiled `markitdown`
command once per file, and writes the results into a **new folder created next to the
original**. The source folder is never touched.

```
C:\data\docs\           ->   C:\data\docs-md\
   report.pdf                   report.md
   notes.docx                   notes.md
   sheet.xlsx                   sheet.md
```

---

## 1. Prerequisites

You need `markitdown` installed with the extras for the formats you care about
(PDF, Office, audio, etc. each pull in their own dependencies):

```powershell
pip install "markitdown[all]"
```

If you are working inside this repository checkout, install into its virtual environment instead:

```powershell
C:\Projects_PoC\markitdown\.venv\Scripts\python.exe -m pip install -e "C:\Projects_PoC\markitdown\packages\markitdown[all]"
```

> `markitdown[pdf]` is now installed in this repo's `.venv`, so PDFs convert and their images
> can be extracted. PPTX, XLSX and audio still need `[all]` (or their own extra) before those
> formats will convert.

`-ExtractImages` needs no extra package: PDF image extraction uses the **pdfminer.six** that
`markitdown[pdf]` already pulls in, and everything else comes out of the Markdown itself.

The script finds the executable itself, in this order:

1. `-MarkItDownPath` if you pass one
2. `.venv\Scripts\markitdown.exe` next to the script or next to its parent folder
3. `markitdown` on `PATH`
4. `python -m markitdown`

## 2. Quick start

```powershell
cd C:\Projects_PoC\markitdown\scripts
.\Convert-FolderToMarkdown.ps1 -InputFolder C:\data\docs
```

That converts the files directly inside `C:\data\docs` and writes them to `C:\data\docs-md`.

If PowerShell blocks the script (`running scripts is disabled on this system`), either
unblock it once:

```powershell
Unblock-File .\Convert-FolderToMarkdown.ps1
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

or run it without changing any policy:

```powershell
powershell -ExecutionPolicy Bypass -File .\Convert-FolderToMarkdown.ps1 -InputFolder C:\data\docs
```

## 3. Parameters

| Parameter | Description |
| --- | --- |
| `-InputFolder` | **Required.** Folder holding the files to convert. Positional, so you can drop the name. |
| `-OutputFolder` | Destination folder. Default: a sibling of the input named `<InputFolder><Suffix>`. |
| `-Suffix` | Suffix for the default output folder name. Default `-md`. |
| `-Recurse` | Also convert subfolders, mirroring the tree in the output. |
| `-IncludeBrowserAssets` | With `-Recurse`, also convert the `page_files\` folder a browser saves next to a page. Skipped by default. See section 8. |
| `-Include` | Only these extensions, e.g. `-Include .pdf,.docx`. Default: the types listed below. |
| `-AllFileTypes` | Convert every file, not only the types markitdown has a converter for. |
| `-Flat` | With `-Recurse`, put every result in the top-level output folder instead of mirroring. |
| `-Force` | Overwrite existing `.md` files. Without it they are skipped. |
| `-MarkItDownPath` | Explicit path to `markitdown.exe` instead of auto-detection. |
| `-UsePlugins` | Passes `-p` to markitdown so third-party plugins are used. |
| `-KeepDataUris` | Passes `--keep-data-uris` so base64 images are kept inline rather than truncated. |
| `-ExtractImages` | Writes embedded images out as real files, copies in the ones a saved page keeps beside it, and repoints the links at them. Implies `-KeepDataUris`. See section 8. |
| `-AssetsFolder` | With `-ExtractImages`, collect every image into this one folder under the output root instead of per-document `<name>.assets`. |
| `-MinImagePixels` | With `-ExtractImages`, skip images smaller than this many pixels (width × height). Default `0` keeps everything; `10000` ≈ 100×100. |
| `-NoImageStitching` | Write each PDF image fragment as its own file instead of reassembling tiled pictures. See section 8. |
| `-WhatIf` | Dry run: lists what would be converted and where, converts nothing. |
| `-Verbose` | Also prints which markitdown executable was selected. |

### Which file types are converted

By default only the formats markitdown has a converter for, as the converters under
`packages/markitdown` declare them:

| Group | Extensions |
| --- | --- |
| Documents | `.pdf` `.docx` `.pptx` `.xlsx` `.xls` |
| Web pages | `.html` `.htm` `.xhtml` |
| Feeds | `.rss` `.atom` `.xml` |
| Text and data | `.csv` `.json` `.jsonl` `.txt` `.text` `.md` `.markdown` |
| Books, notebooks, mail, archives | `.epub` `.ipynb` `.msg` `.zip` |
| Images | `.jpg` `.jpeg` `.png` — EXIF metadata, plus an LLM caption if one is configured |
| Audio | `.wav` `.mp3` `.m4a` `.mp4` — tags, plus a transcript if one is configured |

Anything else is counted and reported rather than converted:

```
files      : 83
ignored    : 1191 files of other types (-Include or -AllFileTypes changes that)
```

This is not a cosmetic filter. markitdown's `PlainTextConverter` accepts **any** file it can
read as text — `.js`, `.css`, `.br`, `.gz`, even a file with no extension at all — and converting
those just hands the source code back as Markdown. Worse, web bundles inline their icons as data
URIs, so one `bundle-client_commons.js` came out as a 4.4 MB `.md` of base64.

`-Include` narrows the list further or picks types outside it (`-Include .pdf,.docx`), and
`-AllFileTypes` goes back to converting every file regardless of type.

## 4. Examples

```powershell
# Everything under C:\data\docs, including subfolders, tree mirrored into C:\data\docs-md
.\Convert-FolderToMarkdown.ps1 C:\data\docs -Recurse

# Only PDFs and Word files, into a folder you name yourself, overwriting previous results
.\Convert-FolderToMarkdown.ps1 C:\data\docs D:\markdown\docs -Include .pdf,.docx -Force

# All subfolders, but collect every .md in one flat folder
.\Convert-FolderToMarkdown.ps1 C:\data\docs -Recurse -Flat

# Pull the pictures out as real files into <name>.assets next to each .md
.\Convert-FolderToMarkdown.ps1 C:\data\docs -Recurse -ExtractImages

# Pages saved with Ctrl+S: the page_files folders are skipped, their pictures copied in
.\Convert-FolderToMarkdown.ps1 C:\data\pages -Recurse -ExtractImages

# ...or gather every image into one shared C:\data\docs-md\assets folder
.\Convert-FolderToMarkdown.ps1 C:\data\docs -Recurse -ExtractImages -AssetsFolder assets

# See the plan before committing to it
.\Convert-FolderToMarkdown.ps1 C:\data\docs -Recurse -WhatIf

# Use a specific markitdown install
.\Convert-FolderToMarkdown.ps1 C:\data\docs -MarkItDownPath C:\tools\venv\Scripts\markitdown.exe

# Capture the result object for scripting
$result = .\Convert-FolderToMarkdown.ps1 C:\data\docs -Recurse
$result.Failed
$result.Failures | Format-Table File, Reason
```

## 5. What the output looks like

```
markitdown : C:\Projects_PoC\markitdown\.venv\Scripts\markitdown.exe
source     : C:\data\docs
output     : C:\data\docs-md
files      : 5

ok      notes.docx -> notes.md
ok      data.json -> data.json.md
FAILED  sheet.xlsx: missing dependency - run: pip install "markitdown[xlsx]"
ok      my report.html -> my report.md
exists  archive.pdf

3 converted, 1 skipped, 1 failed
output: C:\data\docs-md
errors: C:\data\docs-md\_markitdown-errors.log
```

- `ok` — converted, `.md` written
- `exists` — a `.md` was already there; re-run with `-Force` to redo it
- `FAILED` — markitdown returned a non-zero exit code; the full traceback goes to
  `_markitdown-errors.log` in the output folder

The script also returns an object with `Converted`, `Skipped`, `Failed`, `Failures`
and `ErrorLog`, so it composes with other scripts.

## 6. How output files are named

- Normally `report.pdf` becomes `report.md`.
- If two source files in the same target folder share a base name — `report.pdf` and
  `report.docx` — **both** keep their extension: `report.pdf.md` and `report.docx.md`.
  Names are planned before anything runs, so a re-run always produces the same names.
- With `-Flat`, identically named files pulled in from different subfolders get a
  numeric suffix (`test.md`, `test.1.md`).

## 7. Re-running and interruptions

Files that already have a `.md` in the output folder are skipped, so if a long run is
interrupted you can just start it again and it picks up where it stopped. Use `-Force`
when you actually want everything redone.

If a conversion fails, any partial output file is deleted, so a failed file is never
mistaken for a finished one on the next pass.

## 8. What happens to images

markitdown **never writes image files**. Nothing is exported to disk next to the Markdown.
What you get depends on the source format:

| Source | Images in the Markdown |
| --- | --- |
| `.docx` | Inline base64 data URI — **kept only with `-KeepDataUris`**, otherwise truncated to a stub `![alt](data:image/png;base64...)` |
| `.pptx` | Same with `-KeepDataUris`; without it you get `![alt](Picture1.jpg)`, a **placeholder that points at no real file** |
| `.html` | `<img src>` is passed through unchanged, so relative paths and URLs still point at the original location. With `-ExtractImages` the local files are copied in and the links repointed |
| `.pdf` | markitdown itself extracts **text only**. `-ExtractImages` pulls the images out of the PDF directly (see below) |
| `.xlsx`, `.csv`, `.json` | No image handling |
| `.jpg`, `.png` as input | Converted to EXIF metadata (plus an LLM caption if one is configured), not embedded |

### Getting real image files: `-ExtractImages`

`-ExtractImages` writes the embedded pictures out as actual image files and rewrites the
Markdown links to point at them. It turns `--keep-data-uris` on for you, so this is all you need:

```powershell
.\Convert-FolderToMarkdown.ps1 C:\data\docs -Recurse -ExtractImages
```

Each document gets its own assets folder beside its `.md`:

```
docs-md\
   notes.md                       ![...](notes.assets/notes-001.png)
   notes.assets\
      notes-001.png               a real 117 KB PNG
   deck\
      slides.md
      slides.assets\
         slides-001.png
```

Use `-AssetsFolder` to collect everything into one folder under the output root instead:

```powershell
.\Convert-FolderToMarkdown.ps1 C:\data\docs -Recurse -ExtractImages -AssetsFolder assets
```

Details:

- **File type comes from the data URI**, so PNG stays `.png`, JPEG becomes `.jpg`, SVG stays `.svg`.
- **Identical images are written once.** Files are matched by SHA-256, so a logo repeated on
  40 slides produces one file and 40 links to it. In `-AssetsFolder` mode the dedupe spans
  every document, so a shared image is named after whichever document was converted first.
- **Links are URL-escaped**, so folders and files with spaces (`my report.assets/my%20report-001.png`)
  work in any Markdown viewer.
- **Re-runs stay clean.** Before extracting, the script deletes the assets that the *same*
  document wrote previously (files matching `<name>-###.*`), so `-Force` never leaves orphans.
- Size goes back to normal: the same Word file that was 161 KB with inline base64 is **4.7 KB**
  plus a 117 KB PNG next to it.

### Pages saved from a browser

"Save page as - Webpage, complete" (Ctrl+S) writes two things: `page.html`, and a sidecar
`page_files\` folder holding every script, stylesheet and picture the page used. That shape
needs no special flags, but it is worth knowing what the script does with it.

**The sidecar folder is not converted.** With `-Recurse` it would otherwise be walked into, and
its bundles convert into megabytes of noise. The type allowlist already keeps the scripts and
stylesheets out, but a sidecar folder also holds `.html` fragments and `.png` images that *are*
convertible types, so the folder itself is skipped as well. The folders are
found by their pairing with a page of the same stem, so the localised suffixes (`_files`,
`_fichiers`, `_archivos`, `_Dateien`) all work, and the run says how many it left alone:

```
files      : 83
skipped    : 83 browser asset folders (-IncludeBrowserAssets converts them too)
```

Pass `-IncludeBrowserAssets` if you do want the scripts and stylesheets converted as well.

**The pictures are copied in by `-ExtractImages`.** A saved page keeps them as real files, so
there is no data URI to decode. The `<img src>` is followed on disk instead, relative to the
original page, and the file is copied into the assets folder like any other image:

```
pages-md\
   Sample Page.md              ![alt](Sample%20Page.assets/Sample%20Page-002.png)
   Sample Page.assets\
      Sample Page-001.png      was inline as a data URI
      Sample Page-002.png      was .\Sample Page_files\pic one.png
      Sample Page-003.svg      was .\Sample Page_files\logo.svg?v=3
```

Both spellings of a path are followed, so `pic%20one.png` and `pic one.png` reach the same file
and it is written once; a `?v=3` cache buster is ignored; and anything that is not a local image
file — an `https://` URL, a site-relative `/img/logo.svg`, a file that is simply missing — is
left exactly as it was.

Without `-ExtractImages` the links stay as the page wrote them, `./Sample Page_files/pic one.png`,
which resolves next to the original `.html` and not next to the converted `.md` — and the raw
spaces break the link in stricter Markdown viewers.

### PDFs work differently

markitdown converts PDFs with plain text extraction, so a PDF never produces a data URI — there
is nothing inline to pull out. `-ExtractImages` therefore opens the PDF itself and lifts the
image XObjects straight out of it, using the **pdfminer.six** and **Pillow** that
`markitdown[pdf]` already installs. No extra package to install.

Images are decoded and written as **PNG**, or passed straight through as **JPEG** when the PDF
already stores them that way. Nothing is written as BMP.

> This deliberately does *not* use pdfminer's own `ImageWriter`. That writer produces
> uncompressed BMPs, and it decides the format from `LITERAL_DEVICE_RGB in image.colorspace` —
> which is also true for an *indexed* image whose palette happens to be RGB-based. It then writes
> one-byte palette indices as if they were three-byte RGB triples, so the file opens but shows
> noise. In testing, a 320×200 indexed image came out as a 192 KB BMP of garbage; decoded
> properly it is a 31 KB PNG that matches the original exactly. CMYK images were wrong too.

Because the extracted text carries no page markers, the images cannot be placed accurately
mid-document. They are appended in their own section instead, grouped by page:

```markdown
## Extracted images

### Page 1

![Page 1 image 1](photos.assets/photos-001.jpg)

### Page 2

![Page 2 image 1](photos.assets/photos-002.jpg)
```

#### Pictures split into fragments

A PDF very often stores one visible picture as **several image XObjects laid edge to edge** —
a logo cut into pieces, a photo sliced into horizontal strips. Extracting the XObjects one by one
gives you those fragments instead of the picture.

Fragments are reassembled automatically. Images on the same page are grouped when their
placements touch (within 1 PDF unit) *and* line up on the other axis, then pasted onto a single
canvas at the sharpest tile's resolution. A logo split into three pieces comes out as one file.

The alignment requirement is what stops unrelated pictures being glued together: two photos side
by side in a collage have different heights, or a gap between them, so they stay separate. If the
grouping ever gets it wrong, `-NoImageStitching` writes every fragment as its own file.

#### Stencils (`/ImageMask`)

Logos and line art are frequently stored as a **1-bit stencil mask** — the shape is in the image,
the colour comes from the page's graphics state. Treated as a normal image, a stencil decodes to
a solid black rectangle, which is where "black bars instead of a logo" comes from.

Stencils are rendered as **black on a transparent background**, honouring the `Decode` array. The
true paint colour lives in the content stream and is not available to the extractor, so black is
assumed — correct for the overwhelming majority of logos and diagrams, though a stencil that was
painted a light colour on a dark page will look inverted against a white background.

#### Other PDF notes

- **Only raster images come out.** Charts and diagrams drawn as vector graphics — most figures in
  LaTeX papers, for example — are drawing instructions, not images, and there is nothing to
  export. A PDF can legitimately yield zero images.
- **Use `-MinImagePixels` if you get noise.** Some PDFs are built from hundreds of small image
  fragments. `-MinImagePixels 10000` drops anything smaller than roughly 100x100.

```powershell
.\Convert-FolderToMarkdown.ps1 C:\data\pdfs -ExtractImages -MinImagePixels 10000
```

### What extraction cannot recover

- **Vector graphics in PDFs**, as above — they are not images in the file.
- **Remote or site-relative images in HTML**, e.g. `![](/autogen/img/ag.svg)`. Those links point
  at a web server, not at data inside the file, and are passed through unchanged.
- **PPTX placeholders without `-ExtractImages`.** `![alt](Picture1.jpg)` refers to a file that
  never existed; only the data-URI path produces real files.
- **JBIG2-compressed images** are skipped rather than written as files nothing can open — no
  decoder for that format exists in the markitdown dependency stack.
- **JPEG 2000** images are saved as `.jp2` if Pillow was built without OpenJPEG support. Most
  Markdown viewers cannot display those; everything else becomes PNG or JPEG.

Formats handled and verified against the source image: DeviceRGB, DeviceGray, DeviceCMYK
(including the Adobe inversion), Indexed/palette at any bit depth, 1-bit CCITT fax, ICCBased,
Separation and DeviceN, plus `SMask` transparency and inverted `Decode` arrays.

If you would rather keep everything in a single self-contained file, use `-KeepDataUris` on its
own — but expect one Word document to grow from **267 bytes to 161 KB** for a single screenshot,
all on one very long line that editors and Git diffs handle badly. Alt text survives in every
mode, so even with no images at all you still get a description of what the picture was.

## 9. Troubleshooting

**`missing dependency - run: pip install "markitdown[pdf]"`**
The format needs an optional dependency. Install the extra it names, or `markitdown[all]`
for all of them — into the same Python environment the script is using (`-Verbose` shows which).

**`Could not find markitdown.`**
markitdown is not on `PATH` and no `.venv` was found. Install it (`pip install "markitdown[all]"`)
or point the script at it with `-MarkItDownPath`.

**A specific file fails with something else**
Open `_markitdown-errors.log` in the output folder — it holds the full stderr for each
failure. To reproduce a single file by hand:

```powershell
markitdown "C:\data\docs\problem.pdf" -o "C:\temp\problem.md"
```

**Megabytes of base64 in the output, or `.md` files named after `.js` bundles**
That is a saved page's `page_files\` folder being converted. Those folders are skipped now, so
convert again; output from an earlier run can be thrown away. `-ExtractImages` is not the cause —
the base64 sits inside the page's own scripts and stylesheets, which inline their icons that way.

**Nothing was found**
Without `-Recurse` only the top level of the folder is read. Add `-Recurse` for subfolders,
and check that `-Include` extensions match (they are compared case-insensitively, with or
without the leading dot). If the run reports an `ignored` count, the files are there but are
not of a type markitdown converts — name the extension with `-Include`, or pass
`-AllFileTypes`.
