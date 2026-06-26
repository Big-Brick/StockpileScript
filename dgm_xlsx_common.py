from __future__ import annotations

import decimal
from pathlib import Path
from typing import List

import dgm_database

try:
	import openpyxl
except ImportError:
	openpyxl = None  # type: ignore[assignment]

STOP_AFTER_CONSECUTIVE_IGNORED_ROWS = 20
GRAM_NUMBER_FORMAT = "0.#########;-0.#########;0"


def CellHasUsableText(Value: object) -> bool:
	return isinstance(Value, str) and bool(Value.strip())


def DecimalToExcelValue(Value: decimal.Decimal) -> str:
	return dgm_database.DecimalToText(Value)


def ClearDgmCells(Sheet: openpyxl.worksheet.worksheet.Worksheet, Row: int, ColumnsInfo: dgm_database.Columns) -> None:
	for MetalKey, _ in dgm_database.METALS:
		Sheet[f"{ColumnsInfo.PerElement[MetalKey]}{Row}"].value = None
		Sheet[f"{ColumnsInfo.Total[MetalKey]}{Row}"].value = None


def WriteEntryToRow(Sheet: openpyxl.worksheet.worksheet.Worksheet, Row: int, ColumnsInfo: dgm_database.Columns, Entry: dgm_database.ElementRecord) -> None:
	for MetalKey, _ in dgm_database.METALS:
		PerElementCell = Sheet[f"{ColumnsInfo.PerElement[MetalKey]}{Row}"]
		TotalCell = Sheet[f"{ColumnsInfo.Total[MetalKey]}{Row}"]
		PerElementCell.value = DecimalToExcelValue(Entry.GetMetalValue(MetalKey))
		PerElementCell.number_format = GRAM_NUMBER_FORMAT
		TotalCell.value = f"={ColumnsInfo.PerElement[MetalKey]}{Row}*{ColumnsInfo.Quantity}{Row}"
		TotalCell.number_format = GRAM_NUMBER_FORMAT


def FindXlsxFiles(Folder: Path, IncludeSubfolders: bool = False) -> List[Path]:
	Items = Folder.rglob("*") if IncludeSubfolders else Folder.iterdir()
	return sorted(
		PathItem
		for PathItem in Items
		if PathItem.is_file()
		and PathItem.suffix.casefold() == ".xlsx"
		and not PathItem.name.startswith("~$")
	)
