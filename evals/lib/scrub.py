"""Shared leakage/metadata scrub helpers for eval fixture builders."""

import openpyxl


def scrub_metadata(wb: openpyxl.Workbook) -> None:
    """Strip personal/source metadata from a derived seed workbook."""
    props = wb.properties
    props.creator = "eval-fixture"
    props.lastModifiedBy = "eval-fixture"
    props.title = None
    props.subject = None
    props.description = None
    props.keywords = None
    props.category = None
    props.company = None
    props.manager = None
