from enum import Enum


class DocumentType(str, Enum):
    DATASHEET = "datasheet"
    APPNOTE = "appnote"
    SPECIFICATION = "specification"
    USER_MANUAL = "user_manual"
    RELEASE_NOTE = "release_note"

    @classmethod
    def from_value(cls, value):
        if isinstance(value, cls):
            return value

        normalized = str(value).strip().lower()

        for member in cls:
            if member.value == normalized:
                return member

        valid_values = ", ".join(member.value for member in cls)
        raise ValueError(
            f"Unsupported document type '{value}'. Expected one of: {valid_values}."
        )
