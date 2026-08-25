"""Shared data structures for core document processing and template merging."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class ExcelData:
    headers: list[str]
    rows: list[dict[str, str]]
    excel_rows: list[int]


@dataclass
class OutputDirectoryRule:
    """Represents an output directory rule parsed into a fixed base root and relative variable rule."""

    raw_rule: str
    base_folder: str
    relative_rule: str = ""


@dataclass
class TemplateMergeItem:
    """Represents a Word template file in multi-template batch generation."""

    template_id: str
    file_path: str
    display_name: str
    fields: list[str] = field(default_factory=list)
    enabled: bool = True
    filename_rule: str = "{{模板名}}-{{数据序号}}"
    status: str = "待扫描"
    error: str = ""


@dataclass
class MergeJob:
    """Represents a planned, immutable merge task for 1 template × 1 data row."""

    template_id: str
    template_path: str
    template_name: str
    excel_row: int
    data_index: int
    values: dict[str, str]
    relative_folder: str
    filename: str
    destination: str
    filename_rule: str = ""
    requested_filename: str = ""
    collision_reason: str = ""


@dataclass
class MergePlanIssue:
    """Issue or validation warning discovered during task planning."""

    severity: str  # "error", "warning", "info"
    code: str
    message: str
    template_id: str = ""
    template_name: str = ""
    excel_row: int = 0
    rule: str = ""


@dataclass
class MergePlanResult:
    """Structured outcome of multi-template task planning."""

    jobs: list[MergeJob] = field(default_factory=list)
    issues: list[MergePlanIssue] = field(default_factory=list)
    batch_conflicts: list[dict] = field(default_factory=list)
    existing_file_conflicts: list[dict] = field(default_factory=list)


@dataclass
class MergeResult:
    excel_row: int
    filename: str
    success: bool
    template_id: str = ""
    template_name: str = ""
    output_path: str = ""
    relative_folder: str = ""
    replacements: int = 0
    error: str = ""


@dataclass
class FileProcessDetail:
    filename: str
    path: str = ""
    body_count: int = 0
    table_count: int = 0
    total: int = 0
    details: list[str] = field(default_factory=list)
    contexts: list[str] = field(default_factory=list)
    error: str = ""
    backup_path: str = ""


FileProcessResult = FileProcessDetail


@dataclass
class BatchProcessResult:
    total_files: int = 0
    files_processed: int = 0
    total_count: int = 0
    files_with_matches: int = 0
    successful_files: int = 0
    details: list[FileProcessDetail] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    backup_files: list[str] = field(default_factory=list)


@dataclass
class MultiDocItem:
    """Represents a single document and its mapped variable replacements."""

    file_path: str
    filename: str
    detected_variables: list[str] = field(default_factory=list)
    replacements: dict[str, str] = field(default_factory=dict)
    status: str = "就绪"


@dataclass
class MultiDocBatchResult:
    """Outcome of a multi-document batch replacement operation."""

    total_docs: int
    success_docs: int
    failed_docs: int
    errors: list[str] = field(default_factory=list)
    details: list[dict] = field(default_factory=list)
    backup_files: list[str] = field(default_factory=list)
