import struct
import zipfile
from functools import partial
from io import BytesIO
from typing import Any, BinaryIO, Dict, List
from xml.etree import ElementTree as ET

from bs4 import BeautifulSoup, Tag

from .math.omml import OMML_NS, oMath2Latex

# How an inlined comment is rendered. The commented words are delimited so that
# it is clear which text a comment refers to, and replies are chained onto the
# comment they answer.
COMMENT_SPAN_OPEN = "⟦"
COMMENT_SPAN_CLOSE = "⟧"
COMMENT_TEMPLATE = " [comment: {0}]"
COMMENT_REPLY_PREFIX = " ↳ reply: "

MATH_ROOT_TEMPLATE = "".join(
    (
        "<w:document ",
        'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ',
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ',
        'xmlns:o="urn:schemas-microsoft-com:office:office" ',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ',
        'xmlns:v="urn:schemas-microsoft-com:vml" ',
        'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" ',
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ',
        'xmlns:w10="urn:schemas-microsoft-com:office:word" ',
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ',
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ',
        'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ',
        'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" ',
        'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" ',
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14">',
        "{0}</w:document>",
    )
)


def _convert_omath_to_latex(tag: Tag) -> str:
    """
    Converts an OMML (Office Math Markup Language) tag to LaTeX format.

    Args:
        tag (Tag): A BeautifulSoup Tag object representing the OMML element.

    Returns:
        str: The LaTeX representation of the OMML element.
    """
    # Format the tag into a complete XML document string
    math_root = ET.fromstring(MATH_ROOT_TEMPLATE.format(str(tag)))
    # Find the 'oMath' element within the XML document
    math_element = math_root.find(OMML_NS + "oMath")
    if math_element is None:
        return ""
    # Convert the 'oMath' element to LaTeX using the oMath2Latex function
    latex = oMath2Latex(math_element).latex
    return latex


def _get_omath_tag_replacement(tag: Tag, block: bool = False) -> Tag:
    """
    Creates a replacement tag for an OMML (Office Math Markup Language) element.

    Args:
        tag (Tag): A BeautifulSoup Tag object representing the "oMath" element.
        block (bool, optional): If True, the LaTeX will be wrapped in double dollar signs for block mode. Defaults to False.

    Returns:
        Tag: A BeautifulSoup Tag object representing the replacement element.
    """
    t_tag = Tag(name="w:t")
    t_tag.string = (
        f"$${_convert_omath_to_latex(tag)}$$"
        if block
        else f"${_convert_omath_to_latex(tag)}$"
    )
    r_tag = Tag(name="w:r")
    r_tag.append(t_tag)
    return r_tag


def _replace_equations(tag: Tag):
    """
    Replaces OMML (Office Math Markup Language) elements with their LaTeX equivalents.

    Args:
        tag (Tag): A BeautifulSoup Tag object representing the OMML element. Could be either "oMathPara" or "oMath".

    Raises:
        ValueError: If the tag is not supported.
    """
    if tag.name == "oMathPara":
        # Create a new paragraph tag
        p_tag = Tag(name="w:p")
        # Replace each 'oMath' child tag with its LaTeX equivalent as block equations
        for child_tag in tag.find_all("oMath"):
            p_tag.append(_get_omath_tag_replacement(child_tag, block=True))
        # Replace the original 'oMathPara' tag with the new paragraph tag
        tag.replace_with(p_tag)
    elif tag.name == "oMath":
        # Replace the 'oMath' tag with its LaTeX equivalent as inline equation
        tag.replace_with(_get_omath_tag_replacement(tag, block=False))
    else:
        raise ValueError(f"Not supported tag: {tag.name}")


def _pre_process_strike(content: bytes) -> bytes:
    """
    Pre-processes the strikethrough content in a DOCX -> XML file by normalizing double
    strikethrough runs to single strikethrough runs.

    Word marks double strikethrough with "w:dstrike", which downstream converters do not
    recognize, causing those runs to lose their strikethrough entirely. Renaming the tag to
    "w:strike" preserves the strikethrough semantics.

    Args:
        content (bytes): The XML content of the DOCX file as bytes.

    Returns:
        bytes: The processed content with "dstrike" elements renamed to "strike", encoded as bytes.
    """
    # Double strikethrough is rare, and parsing/reserializing the XML is expensive
    # on large documents, so skip the round-trip when there is nothing to rename.
    if b"dstrike" not in content:
        return content

    soup = BeautifulSoup(content.decode(), features="xml")
    for tag in soup.find_all("dstrike"):
        tag.name = "strike"
    return str(soup).encode()


def _pre_process_math(content: bytes) -> bytes:
    """
    Pre-processes the math content in a DOCX -> XML file by converting OMML (Office Math Markup Language) elements to LaTeX.
    This preprocessed content can be directly replaced in the DOCX file -> XMLs.

    Args:
        content (bytes): The XML content of the DOCX file as bytes.

    Returns:
        bytes: The processed content with OMML elements replaced by their LaTeX equivalents, encoded as bytes.
    """
    soup = BeautifulSoup(content.decode(), features="xml")
    for tag in soup.find_all("oMathPara"):
        _replace_equations(tag)
    for tag in soup.find_all("oMath"):
        _replace_equations(tag)
    return str(soup).encode()


def _fix_zip_filename_casing(input_docx: BinaryIO) -> BinaryIO:
    """
    Fix ZIP files where local file header filenames differ in casing
    from the central directory filenames.

    Some document generators (e.g. certain Microsoft Word versions,
    legal document systems) produce .docx/.pptx files where the central
    directory records one casing (e.g. 'customXml/item2.xml') but
    the local file headers record another (e.g. 'customXML/item2.xml').
    Python's zipfile module raises BadZipFile when reading such files.

    This function patches local file header filenames to match the
    central directory, which is the authoritative source used by
    zipfile.ZipFile.
    """
    input_docx.seek(0)
    raw = bytearray(input_docx.read())

    # Read the central directory to get authoritative filenames
    try:
        with zipfile.ZipFile(BytesIO(raw), "r") as zf:
            cd_entries = {
                zi.header_offset: (zi.orig_filename, zi.flag_bits)
                for zi in zf.infolist()
            }
    except zipfile.BadZipFile:
        # Can't even read central directory — return as-is, let it fail later
        input_docx.seek(0)
        return input_docx

    patched = False
    for offset, (cd_name, flag_bits) in cd_entries.items():
        # Verify local file header signature
        if offset + 30 > len(raw) or raw[offset : offset + 4] != b"PK\x03\x04":
            continue

        local_fname_len = struct.unpack_from("<H", raw, offset + 26)[0]
        if offset + 30 + local_fname_len > len(raw):
            continue

        local_name = bytes(raw[offset + 30 : offset + 30 + local_fname_len])
        # ZIP filenames are cp437 unless flag bit 11 marks them as UTF-8, which is
        # how zipfile decoded orig_filename in the first place.
        central_name = cd_name.encode("utf-8" if flag_bits & 0x800 else "cp437")

        # Only patch if lengths match but content differs (casing mismatch)
        if (
            local_name != central_name
            and len(local_name) == len(central_name)
            and local_name.lower() == central_name.lower()
        ):
            raw[offset + 30 : offset + 30 + local_fname_len] = central_name
            patched = True

    if patched:
        return BytesIO(bytes(raw))
    input_docx.seek(0)
    return input_docx


def _pre_process_styles(content: bytes) -> bytes:
    """
    Repairs DOCX style definitions that Mammoth cannot read.

    Mammoth indexes ``w:type`` and ``w:styleId`` directly, so a ``w:style``
    element missing either attribute causes conversion to fail with a
    ``KeyError`` before any document text can be extracted.

    ``w:type`` is optional in OOXML: when it is absent the style type defaults
    to ``paragraph``, so the attribute is filled in rather than dropping the
    style, which would discard its formatting (a heading would be emitted as
    plain body text). A style with no ``w:styleId`` cannot be referenced by the
    document body, so it is removed. Double strikethrough is normalized to
    single strikethrough in the same namespace-aware pass.

    Match elements and attributes by namespace URI, preserving their qualified
    names when repairing the XML. Return the original bytes if no repair is needed.
    """
    from lxml import etree

    namespace = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    root = etree.fromstring(content, parser=parser)
    changed = False

    for style in root.findall(namespace + "style"):
        if namespace + "styleId" not in style.attrib:
            root.remove(style)
            changed = True
        elif namespace + "type" not in style.attrib:
            style.set(namespace + "type", "paragraph")
            changed = True

    for strike in root.iter(namespace + "dstrike"):
        strike.tag = namespace + "strike"
        changed = True

    if not changed:
        return content

    return etree.tostring(root.getroottree(), encoding="utf-8", xml_declaration=True)


def _index_by_comment_id(tags: List[Tag]) -> Dict[str, Tag]:
    """
    Indexes comment-related tags by the comment id they carry.

    Args:
        tags (List[Tag]): BeautifulSoup Tag objects carrying a "w:id" attribute.

    Returns:
        Dict[str, Tag]: Mapping of comment id to the tag that carries it.
    """
    indexed = {}
    for tag in tags:
        tag_id = tag.get("w:id")
        if tag_id is not None:
            indexed[str(tag_id)] = tag
    return indexed


def _get_comment_text(element: Tag) -> str:
    """
    Extracts the plain text of a comment, flattening a multi-paragraph comment
    into a single line.

    Args:
        element (Tag): A BeautifulSoup Tag object representing a "comment" element.

    Returns:
        str: The comment text, with paragraphs separated by spaces.
    """
    paragraphs = []
    for para in element.find_all("p"):
        runs = "".join(t.get_text() for t in para.find_all("t")).strip()
        if runs:
            paragraphs.append(runs)
    if not paragraphs:
        # Comments authored by some tools hold their runs outside of a paragraph
        fallback = "".join(t.get_text() for t in element.find_all("t")).strip()
        if fallback:
            paragraphs.append(fallback)
    return " ".join(paragraphs)


def _load_comments(comments_xml: bytes) -> Dict[str, Dict[str, Any]]:
    """
    Parses word/comments.xml into a mapping of comment id to its text and
    paragraph ids.

    The paragraph ids (w14:paraId) are retained because Word links a reply to
    the comment it answers by paragraph id rather than by comment id.

    Args:
        comments_xml (bytes): The XML content of word/comments.xml.

    Returns:
        Dict[str, Dict[str, Any]]: Mapping of comment id to {"text", "para_ids"}.
    """
    soup = BeautifulSoup(comments_xml.decode(), features="xml")
    comments: Dict[str, Dict[str, Any]] = {}
    for comment in soup.find_all("comment"):
        raw_id = comment.get("w:id")
        if raw_id is None:
            continue
        comments[str(raw_id)] = {
            "text": _get_comment_text(comment),
            "para_ids": [
                p.get("w14:paraId")
                for p in comment.find_all("p")
                if p.get("w14:paraId")
            ],
        }
    return comments


def _load_comment_parents(
    comments_extended_xml: bytes, comments: Dict[str, Dict[str, Any]]
) -> Dict[str, str]:
    """
    Parses word/commentsExtended.xml into a mapping of reply id to parent id.

    Args:
        comments_extended_xml (bytes): The XML content of word/commentsExtended.xml.
        comments (Dict[str, Dict[str, Any]]): The comments, as returned by _load_comments.

    Returns:
        Dict[str, str]: Mapping of a reply's comment id to its parent's comment id.
    """
    para_to_comment = {}
    for comment_id, data in comments.items():
        for para_id in data["para_ids"]:
            para_to_comment[para_id] = comment_id

    soup = BeautifulSoup(comments_extended_xml.decode(), features="xml")
    parents: Dict[str, str] = {}
    for comment_ex in soup.find_all("commentEx"):
        para_id = comment_ex.get("w15:paraId")
        parent_para_id = comment_ex.get("w15:paraIdParent")
        if not para_id or not parent_para_id:
            continue
        child = para_to_comment.get(para_id)
        parent = para_to_comment.get(parent_para_id)
        if child and parent and child != parent:
            parents[child] = parent
    return parents


def _build_comment_threads(
    comments: Dict[str, Dict[str, Any]], parents: Dict[str, str]
) -> Dict[str, str]:
    """
    Groups replies with the comment they answer and renders each thread.

    A thread is rendered once, at the place its first comment is anchored, so
    that a chain of replies stays next to the text it discusses.

    Args:
        comments (Dict[str, Dict[str, Any]]): The comments, as returned by _load_comments.
        parents (Dict[str, str]): Reply-to-parent mapping, as returned by _load_comment_parents.

    Returns:
        Dict[str, str]: Mapping of the thread's root comment id to its rendered text.
    """

    def root_of(comment_id: str) -> str:
        seen = {comment_id}
        while comment_id in parents:
            comment_id = parents[comment_id]
            if comment_id in seen:  # guard against malformed, cyclic parent data
                break
            seen.add(comment_id)
        return comment_id

    grouped: Dict[str, List[str]] = {}
    for comment_id in sorted(
        comments, key=lambda c: int(c) if c.lstrip("-").isdigit() else 0
    ):
        grouped.setdefault(root_of(comment_id), []).append(comment_id)

    threads = {}
    for root_id, thread in grouped.items():
        # The root opens the thread even if a reply was given a lower id
        ordered = [root_id] + [c for c in thread if c != root_id]
        texts = [comments[c]["text"] for c in ordered if comments[c]["text"]]
        if texts:
            threads[root_id] = COMMENT_TEMPLATE.format(
                texts[0] + "".join(COMMENT_REPLY_PREFIX + t for t in texts[1:])
            )
    return threads


def _load_comment_threads(files: Dict[str, bytes]) -> Dict[str, str]:
    """
    Reads the comment-related parts of a DOCX file and renders each thread.

    Args:
        files (Dict[str, bytes]): The contents of the DOCX file, keyed by name.

    Returns:
        Dict[str, str]: Mapping of the thread's root comment id to its rendered text.
    """
    comments_xml = files.get("word/comments.xml")
    if not comments_xml:
        return {}

    comments = _load_comments(comments_xml)
    if not comments:
        return {}

    # commentsExtended.xml is absent unless the document has threaded replies
    comments_extended_xml = files.get("word/commentsExtended.xml")
    parents = (
        _load_comment_parents(comments_extended_xml, comments)
        if comments_extended_xml
        else {}
    )
    return _build_comment_threads(comments, parents)


def _get_text_tag_replacement(text: str) -> Tag:
    """
    Creates a run holding a literal piece of text.

    Args:
        text (str): The text to wrap.

    Returns:
        Tag: A BeautifulSoup Tag object representing a "w:r" element.
    """
    t_tag = Tag(name="w:t")
    t_tag["xml:space"] = "preserve"
    t_tag.string = text
    r_tag = Tag(name="w:r")
    r_tag.append(t_tag)
    return r_tag


def _pre_process_comments(content: bytes, threads: Dict[str, str]) -> bytes:
    """
    Pre-processes a DOCX -> XML file by inlining comment text into the body.

    The commented words are delimited so that it is clear what a comment refers
    to, and the comment itself is placed immediately after them. Comments are
    injected as ordinary runs, so they survive conversion to HTML and Markdown.

    Args:
        content (bytes): The XML content of word/document.xml as bytes.
        threads (Dict[str, str]): Rendered threads keyed by root comment id.

    Returns:
        bytes: The processed content with comments inlined, encoded as bytes.
    """
    soup = BeautifulSoup(content.decode(), features="xml")

    range_starts = _index_by_comment_id(soup.find_all("commentRangeStart"))
    range_ends = _index_by_comment_id(soup.find_all("commentRangeEnd"))
    # Anchor on the enclosing run so the comment lands outside of it
    references = {
        comment_id: tag.find_parent("r") or tag
        for comment_id, tag in _index_by_comment_id(
            soup.find_all("commentReference")
        ).items()
    }

    for comment_id, rendered in threads.items():
        range_start = range_starts.get(comment_id)
        range_end = range_ends.get(comment_id)
        if range_start is not None and range_end is not None:
            # The comment covers a span of text: delimit it, then append the comment
            range_start.insert_after(_get_text_tag_replacement(COMMENT_SPAN_OPEN))
            range_end.insert_before(_get_text_tag_replacement(COMMENT_SPAN_CLOSE))
            range_end.insert_after(_get_text_tag_replacement(rendered))
        else:
            # No span was recorded, so fall back to the comment's anchor point
            anchor = references.get(comment_id)
            if anchor is not None:
                anchor.insert_after(_get_text_tag_replacement(rendered))

    return str(soup).encode()


def pre_process_docx(input_docx: BinaryIO, inline_comments: bool = True) -> BinaryIO:
    """
    Pre-processes a DOCX file with provided steps.

    The process works by unzipping the DOCX file in memory, transforming specific XML files
    (such as converting OMML elements to LaTeX), and then zipping everything back into a
    DOCX file without writing to disk.

    Args:
        input_docx (BinaryIO): A binary input stream representing the DOCX file.
        inline_comments (bool, optional): If True, review comments are inlined next to
            the text they annotate, with replies chained onto the comment they answer.
            Defaults to True.

    Returns:
        BinaryIO: A binary output stream representing the processed DOCX file.
    """
    # Fix ZIP filename casing mismatch before any processing
    input_docx = _fix_zip_filename_casing(input_docx)

    output_docx = BytesIO()
    # The pre-processing steps to apply to each file in the .docx
    pre_process_enable_files = {
        "word/document.xml": (_pre_process_strike, _pre_process_math),
        "word/footnotes.xml": (_pre_process_strike, _pre_process_math),
        "word/endnotes.xml": (_pre_process_strike, _pre_process_math),
        "word/styles.xml": (_pre_process_styles,),
    }
    with zipfile.ZipFile(input_docx, mode="r") as zip_input:
        files = {name: zip_input.read(name) for name in zip_input.namelist()}

        if inline_comments:
            try:
                threads = _load_comment_threads(files)
            except Exception:
                # If the comments cannot be read, convert the document without them
                threads = {}
            if threads:
                # Runs added here must not be re-processed by the earlier steps
                pre_process_enable_files["word/document.xml"] += (
                    partial(_pre_process_comments, threads=threads),
                )

        with zipfile.ZipFile(output_docx, mode="w") as zip_output:
            zip_output.comment = zip_input.comment
            for name, content in files.items():
                updated_content = content
                # Each step is applied independently, so one failing step does
                # not discard the results of the others.
                for pre_process_step in pre_process_enable_files.get(name, ()):
                    try:
                        updated_content = pre_process_step(updated_content)
                    except Exception:
                        pass
                zip_output.writestr(name, updated_content)
    output_docx.seek(0)
    return output_docx
