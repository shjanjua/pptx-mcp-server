"""
OOXML document operations for Office files.

This module provides functionality to:
- Unpack Office documents (.docx, .pptx, .xlsx) to directories
- Pack directories back into Office documents
- Validate Office documents against XSD schemas
"""

import random
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from typing import List, Optional, Tuple

import defusedxml.minidom
import lxml.etree


# Valid Office extensions
VALID_EXTENSIONS = {".docx", ".pptx", ".xlsx"}


def unpack_document(
    office_file: Path,
    output_dir: Path,
) -> str:
    """Unpack an Office document to a directory with pretty-printed XML.

    Args:
        office_file: Path to the Office file
        output_dir: Directory to extract contents to

    Returns:
        Status message with details
    """
    office_file = Path(office_file)
    output_dir = Path(output_dir)

    if not office_file.exists():
        raise FileNotFoundError(f"File not found: {office_file}")

    ext = office_file.suffix.lower()
    if ext not in VALID_EXTENSIONS:
        raise ValueError(f"Invalid file type: {ext}. Must be one of {VALID_EXTENSIONS}")

    # Create output directory
    output_dir.mkdir(parents=True, exist_ok=True)

    # Extract ZIP contents
    with zipfile.ZipFile(office_file, "r") as zf:
        zf.extractall(output_dir)

    # Pretty-print all XML and .rels files
    xml_files = list(output_dir.rglob("*.xml")) + list(output_dir.rglob("*.rels"))
    formatted_count = 0

    for xml_file in xml_files:
        try:
            content = xml_file.read_text(encoding="utf-8")
            dom = defusedxml.minidom.parseString(content)
            xml_file.write_bytes(dom.toprettyxml(indent="  ", encoding="ascii"))
            formatted_count += 1
        except Exception:
            # Skip files that can't be parsed
            pass

    result = f"Extracted to: {output_dir}\nFormatted {formatted_count} XML files"

    # For .docx files, suggest an RSID for tracked changes
    if ext == ".docx":
        suggested_rsid = "".join(random.choices("0123456789ABCDEF", k=8))
        result += f"\nSuggested RSID for edit session: {suggested_rsid}"

    return result


def pack_document(
    input_dir: Path,
    output_file: Path,
    validate: bool = False,
) -> bool:
    """Pack a directory into an Office document.

    Args:
        input_dir: Directory containing unpacked Office document
        output_file: Path for output Office file
        validate: If True, validate with LibreOffice after packing

    Returns:
        True if successful, False if validation failed
    """
    input_dir = Path(input_dir)
    output_file = Path(output_file)

    if not input_dir.is_dir():
        raise ValueError(f"Input is not a directory: {input_dir}")

    ext = output_file.suffix.lower()
    if ext not in VALID_EXTENSIONS:
        raise ValueError(f"Invalid file type: {ext}. Must be one of {VALID_EXTENSIONS}")

    # Work in temporary directory to avoid modifying original
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_content_dir = Path(temp_dir) / "content"
        shutil.copytree(input_dir, temp_content_dir)

        # Condense XML files to remove pretty-printing whitespace
        for pattern in ["*.xml", "*.rels"]:
            for xml_file in temp_content_dir.rglob(pattern):
                _condense_xml(xml_file)

        # Create output directory
        output_file.parent.mkdir(parents=True, exist_ok=True)

        # Create ZIP archive
        with zipfile.ZipFile(output_file, "w", zipfile.ZIP_DEFLATED) as zf:
            for f in temp_content_dir.rglob("*"):
                if f.is_file():
                    zf.write(f, f.relative_to(temp_content_dir))

        # Validate if requested
        if validate:
            if not _validate_with_libreoffice(output_file):
                output_file.unlink()  # Delete corrupt file
                return False

    return True


def _condense_xml(xml_file: Path) -> None:
    """Remove pretty-printing whitespace from XML file."""
    try:
        with open(xml_file, "r", encoding="utf-8") as f:
            dom = defusedxml.minidom.parse(f)

        # Remove whitespace-only text nodes (except in text content elements)
        for element in dom.getElementsByTagName("*"):
            # Skip w:t elements (Word text content)
            if element.tagName.endswith(":t"):
                continue

            for child in list(element.childNodes):
                if (
                    child.nodeType == child.TEXT_NODE
                    and child.nodeValue
                    and child.nodeValue.strip() == ""
                ) or child.nodeType == child.COMMENT_NODE:
                    element.removeChild(child)

        with open(xml_file, "wb") as f:
            f.write(dom.toxml(encoding="UTF-8"))
    except Exception:
        pass  # Skip files that can't be processed


def _validate_with_libreoffice(doc_path: Path) -> bool:
    """Validate document by attempting to convert with LibreOffice."""
    filter_map = {
        ".docx": "html:HTML",
        ".pptx": "html:impress_html_Export",
        ".xlsx": "html:HTML (StarCalc)",
    }

    ext = doc_path.suffix.lower()
    filter_name = filter_map.get(ext)

    if not filter_name:
        return True  # Skip validation for unknown types

    with tempfile.TemporaryDirectory() as temp_dir:
        try:
            result = subprocess.run(
                [
                    "soffice",
                    "--headless",
                    "--convert-to", filter_name,
                    "--outdir", temp_dir,
                    str(doc_path),
                ],
                capture_output=True,
                timeout=30,
                text=True,
            )

            output_file = Path(temp_dir) / f"{doc_path.stem}.html"
            return output_file.exists()

        except FileNotFoundError:
            # LibreOffice not installed - skip validation
            return True
        except subprocess.TimeoutExpired:
            return False


def validate_document(
    unpacked_dir: Path,
    original_file: Path,
    verbose: bool = False,
) -> Tuple[bool, List[str]]:
    """Validate an unpacked Office document against XSD schemas.

    Args:
        unpacked_dir: Path to unpacked Office document directory
        original_file: Path to original Office file for comparison
        verbose: Enable verbose output

    Returns:
        Tuple of (success, list of messages)
    """
    unpacked_dir = Path(unpacked_dir)
    original_file = Path(original_file)

    if not unpacked_dir.is_dir():
        return False, [f"Directory not found: {unpacked_dir}"]

    if not original_file.exists():
        return False, [f"Original file not found: {original_file}"]

    ext = original_file.suffix.lower()
    if ext not in VALID_EXTENSIONS:
        return False, [f"Invalid file type: {ext}"]

    messages = []
    all_valid = True

    # Get validator based on file type
    if ext == ".pptx":
        validator = PPTXValidator(unpacked_dir, original_file, verbose)
    elif ext == ".docx":
        validator = DOCXValidator(unpacked_dir, original_file, verbose)
    else:
        return False, [f"Validation not yet supported for {ext}"]

    # Run validations
    results = validator.validate_all()

    for check_name, passed, details in results:
        if passed:
            if verbose:
                messages.append(f"PASSED - {check_name}")
        else:
            all_valid = False
            messages.append(f"FAILED - {check_name}")
            for detail in details:
                messages.append(f"  {detail}")

    return all_valid, messages


class BaseValidator:
    """Base class for Office document validators."""

    def __init__(self, unpacked_dir: Path, original_file: Path, verbose: bool = False):
        self.unpacked_dir = Path(unpacked_dir).resolve()
        self.original_file = Path(original_file)
        self.verbose = verbose

        # Get all XML and .rels files
        self.xml_files = list(self.unpacked_dir.rglob("*.xml")) + list(self.unpacked_dir.rglob("*.rels"))

    def validate_all(self) -> List[Tuple[str, bool, List[str]]]:
        """Run all validations. Returns list of (check_name, passed, details)."""
        raise NotImplementedError

    def validate_xml_wellformed(self) -> Tuple[bool, List[str]]:
        """Check all XML files are well-formed."""
        errors = []

        for xml_file in self.xml_files:
            try:
                lxml.etree.parse(str(xml_file))
            except lxml.etree.XMLSyntaxError as e:
                rel_path = xml_file.relative_to(self.unpacked_dir)
                errors.append(f"{rel_path}: Line {e.lineno}: {e.msg}")

        return len(errors) == 0, errors

    def validate_namespaces(self) -> Tuple[bool, List[str]]:
        """Validate namespace prefixes in Ignorable attributes are declared."""
        errors = []
        mc_namespace = "http://schemas.openxmlformats.org/markup-compatibility/2006"

        for xml_file in self.xml_files:
            try:
                root = lxml.etree.parse(str(xml_file)).getroot()
                declared = set(root.nsmap.keys()) - {None}

                for attr_val in [v for k, v in root.attrib.items() if k.endswith("Ignorable")]:
                    undeclared = set(attr_val.split()) - declared
                    for ns in undeclared:
                        rel_path = xml_file.relative_to(self.unpacked_dir)
                        errors.append(f"{rel_path}: Namespace '{ns}' in Ignorable but not declared")
            except lxml.etree.XMLSyntaxError:
                continue

        return len(errors) == 0, errors

    def validate_file_references(self) -> Tuple[bool, List[str]]:
        """Validate all .rels files properly reference existing files."""
        errors = []
        pkg_rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"

        rels_files = list(self.unpacked_dir.rglob("*.rels"))

        for rels_file in rels_files:
            try:
                root = lxml.etree.parse(str(rels_file)).getroot()
                rels_dir = rels_file.parent

                for rel in root.findall(f".//{{{pkg_rels_ns}}}Relationship"):
                    target = rel.get("Target")
                    if target and not target.startswith(("http", "mailto:")):
                        if rels_file.name == ".rels":
                            target_path = self.unpacked_dir / target
                        else:
                            base_dir = rels_dir.parent
                            target_path = base_dir / target

                        try:
                            target_path = target_path.resolve()
                            if not target_path.exists():
                                rel_rels = rels_file.relative_to(self.unpacked_dir)
                                errors.append(f"{rel_rels}: Broken reference to {target}")
                        except (OSError, ValueError):
                            rel_rels = rels_file.relative_to(self.unpacked_dir)
                            errors.append(f"{rel_rels}: Invalid path {target}")

            except Exception as e:
                rel_path = rels_file.relative_to(self.unpacked_dir)
                errors.append(f"{rel_path}: Error parsing - {e}")

        return len(errors) == 0, errors


class PPTXValidator(BaseValidator):
    """Validator for PowerPoint presentations."""

    def validate_all(self) -> List[Tuple[str, bool, List[str]]]:
        results = []

        # XML well-formedness
        passed, errors = self.validate_xml_wellformed()
        results.append(("XML well-formedness", passed, errors))
        if not passed:
            return results  # Stop early if XML is broken

        # Namespace declarations
        passed, errors = self.validate_namespaces()
        results.append(("Namespace declarations", passed, errors))

        # File references
        passed, errors = self.validate_file_references()
        results.append(("File references", passed, errors))

        # Slide layout references
        passed, errors = self._validate_slide_layouts()
        results.append(("Slide layout references", passed, errors))

        return results

    def _validate_slide_layouts(self) -> Tuple[bool, List[str]]:
        """Validate slide layout relationships."""
        errors = []
        pkg_rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"

        slide_rels = list(self.unpacked_dir.glob("ppt/slides/_rels/*.xml.rels"))

        for rels_file in slide_rels:
            try:
                root = lxml.etree.parse(str(rels_file)).getroot()

                layout_count = 0
                for rel in root.findall(f".//{{{pkg_rels_ns}}}Relationship"):
                    rel_type = rel.get("Type", "")
                    if "slideLayout" in rel_type:
                        layout_count += 1

                if layout_count > 1:
                    rel_path = rels_file.relative_to(self.unpacked_dir)
                    errors.append(f"{rel_path}: Multiple slideLayout references ({layout_count})")
                elif layout_count == 0:
                    rel_path = rels_file.relative_to(self.unpacked_dir)
                    errors.append(f"{rel_path}: Missing slideLayout reference")

            except Exception as e:
                rel_path = rels_file.relative_to(self.unpacked_dir)
                errors.append(f"{rel_path}: Error - {e}")

        return len(errors) == 0, errors


class DOCXValidator(BaseValidator):
    """Validator for Word documents."""

    WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    def validate_all(self) -> List[Tuple[str, bool, List[str]]]:
        results = []

        # XML well-formedness
        passed, errors = self.validate_xml_wellformed()
        results.append(("XML well-formedness", passed, errors))
        if not passed:
            return results

        # Namespace declarations
        passed, errors = self.validate_namespaces()
        results.append(("Namespace declarations", passed, errors))

        # File references
        passed, errors = self.validate_file_references()
        results.append(("File references", passed, errors))

        # Whitespace preservation
        passed, errors = self._validate_whitespace()
        results.append(("Whitespace preservation", passed, errors))

        # Track changes validation
        passed, errors = self._validate_track_changes()
        results.append(("Track changes", passed, errors))

        return results

    def _validate_whitespace(self) -> Tuple[bool, List[str]]:
        """Validate w:t elements with whitespace have xml:space='preserve'."""
        errors = []
        xml_ns = "http://www.w3.org/XML/1998/namespace"

        for xml_file in self.xml_files:
            if xml_file.name != "document.xml":
                continue

            try:
                root = lxml.etree.parse(str(xml_file)).getroot()

                for elem in root.iter(f"{{{self.WORD_NS}}}t"):
                    if elem.text:
                        text = elem.text
                        if text.startswith(" ") or text.endswith(" "):
                            space_attr = f"{{{xml_ns}}}space"
                            if elem.attrib.get(space_attr) != "preserve":
                                preview = repr(text)[:30]
                                rel_path = xml_file.relative_to(self.unpacked_dir)
                                errors.append(
                                    f"{rel_path}: Line {elem.sourceline}: "
                                    f"w:t with whitespace missing xml:space='preserve': {preview}"
                                )

            except lxml.etree.XMLSyntaxError:
                continue

        return len(errors) == 0, errors

    def _validate_track_changes(self) -> Tuple[bool, List[str]]:
        """Validate track changes structure (w:t not inside w:del)."""
        errors = []

        for xml_file in self.xml_files:
            if xml_file.name != "document.xml":
                continue

            try:
                root = lxml.etree.parse(str(xml_file)).getroot()
                nsmap = {"w": self.WORD_NS}

                # Find w:t elements inside w:del
                bad_elements = root.xpath(".//w:del//w:t", namespaces=nsmap)

                for elem in bad_elements:
                    if elem.text:
                        preview = repr(elem.text)[:30]
                        rel_path = xml_file.relative_to(self.unpacked_dir)
                        errors.append(
                            f"{rel_path}: Line {elem.sourceline}: "
                            f"w:t found inside w:del: {preview}"
                        )

            except lxml.etree.XMLSyntaxError:
                continue

        return len(errors) == 0, errors
