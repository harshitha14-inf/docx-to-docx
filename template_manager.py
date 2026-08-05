from pathlib import Path
import json

from document_types import DocumentType


class TemplateNotFoundError(FileNotFoundError):
    pass


class TemplateManager:

    def __init__(self, project_root=None):
        self.project_root = Path(project_root or Path(__file__).resolve().parent)
        self.rules = self._load_rules()

    def get_template(self, document_type):
        doc_type = DocumentType.from_value(document_type)
        allowed_types = set(self.rules.get("allowed_document_types", []))

        if doc_type.value not in allowed_types:
            allowed_text = ", ".join(sorted(allowed_types))
            raise ValueError(
                f"Document type '{doc_type.value}' is not allowed by policy."
                f" Allowed values: {allowed_text}."
            )

        template_name = self.rules.get("required_templates", {}).get(doc_type.value)
        if template_name is None:
            raise ValueError(f"No template mapping configured for document type '{doc_type.value}'.")

        primary_path = self.project_root / "templates" / template_name
        if primary_path.exists():
            return str(primary_path)

        raise TemplateNotFoundError(
            "Template file is missing for "
            f"'{doc_type.value}'. Expected at '{primary_path}'"
            "."
        )

    def _load_rules(self):

        config_path = self.project_root / "config" / "document_type_rules.json"

        if not config_path.exists():
            raise FileNotFoundError(
                f"Missing required document type rules file: {config_path}"
            )

        with open(config_path, "r", encoding="utf-8") as handle:
            return json.load(handle)


def get_template(document_type, project_root=None):
    return TemplateManager(project_root=project_root).get_template(document_type)
