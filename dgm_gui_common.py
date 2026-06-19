from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import List

import dgm_database

try:
	import openpyxl
except ImportError:
	openpyxl = None  # type: ignore[assignment]


@dataclass
class GuiMissingElement:
	FilePath: Path
	SheetName: str
	Row: int
	Name: str


@dataclass
class GuiConflictRow:
	FilePath: Path
	SheetName: str
	Row: int
	Name: str
	Record: dgm_database.ElementRecord
	SheetValues: dgm_database.DgmValues
	DatabaseValues: dgm_database.DgmValues
	Details: str


@dataclass
class GuiProcessResult:
	FilePath: Path
	ProcessedRows: int
	IgnoredRows: int
	Missing: List[GuiMissingElement]
	Conflicts: List[GuiConflictRow]


@dataclass
class GuiAddElementResult:
	Mode: str
	Values: dgm_database.DgmValues
	PathParts: List[str]
	Pattern: str = ""
	DisplayText: str = ""


WINDOW_TITLE = "DGM Database Editor"
