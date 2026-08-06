import json
from dataclasses import dataclass
from pathlib import Path


DEFAULT_STYLE_MAP = {
    "heading_1": "Heading 1",
    "heading_2": "Heading 2",
    "heading_3": "Heading 3",
    "heading_4": "Heading 4",
    "heading_5": "Heading 4",
    "heading_6": "Heading 4",
    "body_paragraph": "Normal",
    "figure_caption": "Caption",
    "table_caption": "Caption",
    "table_style": {
        "type": "alias",
        "maps_to": [
            "TableHead",
            "TableHead-l",
            "TableHead-c",
            "TableCell",
            "TableCell-l",
            "TableCell-c",
            "TableCellBold",
            "TableCellBold-I",
        ],
    },
    "header_paragraph": "Header",
    "footer_paragraph": "Footer",
    "list_bullet": {
        "type": "semantic",
        "maps_to": "A bullet",
    },
    "list_number": {
        "type": "fallback",
        "maps_to": "Normal",
    },
    "glossary_term": "Normal",
    "reference_entry": "Normal",
    "revision_history": "Default Table",
    "code_listing": "Code",
}


@dataclass
class StyleMapper:
    mapping: dict

    @classmethod
    def from_file(cls, config_path=None):
        config_file = Path(config_path or "config/style_map.json")

        if not config_file.exists():
            return cls(mapping=dict(DEFAULT_STYLE_MAP))

        with open(config_file, "r", encoding="utf-8") as handle:
            data = json.load(handle)

        merged = dict(DEFAULT_STYLE_MAP)
        merged.update(data)
        return cls(mapping=merged)

    def style_for_heading(self, level):
        return self._resolve_mapping_value(
            f"heading_{level}",
            self.mapping.get("body_paragraph", self.mapping.get("paragraph")),
        )

    def style_for_caption(self, caption_type):
        if caption_type == "table":
            return self._resolve_mapping_value(
                "table_caption",
                self.mapping.get("body_paragraph", self.mapping.get("paragraph")),
            )
        return self._resolve_mapping_value(
            "figure_caption",
            self.mapping.get("body_paragraph", self.mapping.get("paragraph")),
        )

    def style_for_list(self, list_kind):
        if list_kind == "number":
            return self._resolve_mapping_value(
                "list_number",
                self.mapping.get("body_paragraph", self.mapping.get("paragraph")),
            )

        bullet_key = "list_bullet"

        if "bullet_list" in self.mapping:
            bullet_key = "bullet_list"

        return self._resolve_mapping_value(
            bullet_key,
            self.mapping.get("body_paragraph", self.mapping.get("paragraph")),
        )

    def style_for_item(self, item):
        name = item.__class__.__name__.lower()

        if name == "heading":
            level = getattr(item, "level", 1)
            return self.style_for_heading(level)

        if name == "caption":
            caption_type = getattr(item, "caption_type", "image")
            return self.style_for_caption(caption_type)

        if name == "table":
            return self._resolve_mapping_value(
                "table_style",
                self.mapping.get("table"),
            )

        if name == "paragraph":
            return self._resolve_mapping_value(
                "body_paragraph",
                self.mapping.get("paragraph"),
            )

        if name == "formula":
            return self._resolve_mapping_value(
                "code_listing",
                self.mapping.get("body_paragraph", self.mapping.get("paragraph")),
            )

        return None

    def style_for_table_cell(self, is_header, alignment="left", bold=False, italic=False):

        if is_header:
            if alignment == "center":
                return self._resolve_mapping_value(
                    "table_header_center",
                    self._resolve_mapping_value("table_header"),
                )
            return self._resolve_mapping_value(
                "table_header_left",
                self._resolve_mapping_value("table_header"),
            )

        if bold and italic:
            resolved = self._resolve_mapping_value("table_cell_bold_italic")
            if resolved:
                return resolved

        if bold:
            resolved = self._resolve_mapping_value("table_cell_bold")
            if resolved:
                return resolved

        if alignment == "center":
            return self._resolve_mapping_value(
                "table_cell_center",
                self._resolve_mapping_value("table_cell"),
            )

        return self._resolve_mapping_value(
            "table_cell_left",
            self._resolve_mapping_value("table_cell"),
        )

    def _resolve_mapping_value(self, key, default=None):

        if key not in self.mapping:
            return default

        value = self.mapping[key]

        if isinstance(value, str):
            return value

        if isinstance(value, dict):
            mapped = value.get("maps_to")

            if isinstance(mapped, list):
                return mapped[0] if mapped else default

            if isinstance(mapped, str):
                return mapped

            return default

        return default
