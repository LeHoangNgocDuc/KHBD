from __future__ import annotations

from pydantic import BaseModel, Field


class LessonMeta(BaseModel):
    subject_code: str = "DS"
    grade: str = "9"
    week: int = 1
    lesson_number: int = 1
    lesson_number_end: int | None = None
    lesson_index: int = 1
    chapter: str = "CHƯƠNG I."
    chapter_title: str = ""
    lesson_title: str = ""
    lesson_periods: int = 2
    date_start: str = ""
    prepare_date: str = ""
    classes: str = ""
    enable_nls: bool = True
    enable_ai: bool = True
    enable_gdhn: bool = False
    disability: str = "vận động/thị giác"
    previous_week: int | None = None
    previous_periods: str = ""
    header_only: bool = False
    lesson_content: dict = Field(default_factory=dict)


class MergeRequest(BaseModel):
    file_ids: list[str]

