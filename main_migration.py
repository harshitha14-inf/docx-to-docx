import argparse
from pathlib import Path
import json

from builder import Builder
from branding import BrandingEngine
from document_types import DocumentType
from extractor import Extractor
from style_engine.style_mapper import StyleMapper
from template_manager import TemplateManager
from validator import validate


DEFAULT_SOURCE = "samples/sample2.docx"
DEFAULT_OUTPUT = "samples/roundtrip_output.docx"


def _load_document_type_rules(project_root):

    config_path = Path(project_root) / "config" / "document_type_rules.json"

    with open(config_path, "r", encoding="utf-8") as handle:
        return json.load(handle)


def _resolve_document_type(args, rules):

    selected = args.document_type

    if not selected:
        selected = input("Select document type (datasheet/appnote/specification/user_manual/release_note): ").strip()

    if not selected:
        raise ValueError("Document type is mandatory by policy and cannot be empty.")

    doc_type = DocumentType.from_value(selected)

    allowed = set(rules.get("allowed_document_types", []))
    if doc_type.value not in allowed:
        allowed_text = ", ".join(sorted(allowed))
        raise ValueError(f"Document type '{doc_type.value}' is not allowed. Allowed values: {allowed_text}.")

    return doc_type


def run_migration(source, output, document_type, cover_metadata=None):

    project_root = Path(__file__).resolve().parent
    rules = _load_document_type_rules(project_root)

    doc_type = _resolve_document_type(
        argparse.Namespace(document_type=document_type),
        rules,
    )

    template_path = TemplateManager(project_root).get_template(doc_type)

    model = Extractor(source).extract()
    model = BrandingEngine().transform_document(model)

    Builder(project_root=project_root).build(
        model,
        output,
        document_type=doc_type,
        selected_template=template_path,
        preserve_template_cover=True,
        apply_style_mapping=True,
        cover_metadata=cover_metadata,
    )

    report = validate(
        source,
        output,
        template_used=template_path,
        style_mapper=StyleMapper.from_file(project_root / "config" / "style_map.json"),
        strict_roundtrip=False,
    )

    return report


def main():

    parser = argparse.ArgumentParser()
    parser.add_argument("--source", default=DEFAULT_SOURCE)
    parser.add_argument("--output", default=DEFAULT_OUTPUT)
    parser.add_argument("--document-type", dest="document_type", default=None)

    args = parser.parse_args()

    project_root = Path(__file__).resolve().parent
    rules = _load_document_type_rules(project_root)
    doc_type = _resolve_document_type(args, rules)

    report = run_migration(
        source=args.source,
        output=args.output,
        document_type=doc_type.value,
    )

    print(report["summary"])


if __name__ == "__main__":
    main()
