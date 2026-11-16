"""Exception hierarchy.  All grammar errors are ``OoxValidationError``."""


class OoxError(Exception):
    """Base class for all oox errors."""


class OoxValidationError(OoxError, ValueError):
    """The operation would produce markup that violates the schema."""


class InvalidChild(OoxValidationError):
    pass


class DuplicateChild(OoxValidationError):
    pass


class InvalidAttribute(OoxValidationError, AttributeError):
    pass


class InvalidAttributeValue(OoxValidationError):
    pass


class MissingAttribute(OoxValidationError):
    pass


class TargetError(OoxValidationError):
    """``save(target=...)``: the document uses markup that Office version
    can't read.  ``findings`` maps part names to messages."""

    def __init__(self, message: str, findings: dict[str, list[str]] | None = None):
        super().__init__(message)
        self.findings = findings or {}


class PackageError(OoxError):
    """Malformed or unsupported OPC package."""
