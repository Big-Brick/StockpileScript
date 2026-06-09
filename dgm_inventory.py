#!/usr/bin/env python3
"""
Fill DGM metal content columns in XLSX inventory files using an XML database.

Required dependency:
    pip install openpyxl

Usage:
    python dgm_inventory.py database.xml [folder_with_xlsx_files]

If the folder argument is omitted, a folder selection dialog is shown.
"""

from __future__ import annotations

import decimal
import shutil
import sys
from pathlib import Path
from typing import Iterable, List, Optional, Set, Tuple

import dgm_database

try:
	import openpyxl
except ImportError:
	print("Missing dependency: openpyxl", file=sys.stderr)
	print("Install it with: pip install openpyxl", file=sys.stderr)
	sys.exit(2)


CREATE_BACKUP = True
STOP_AFTER_CONSECUTIVE_IGNORED_ROWS = 20
DEFAULT_SHEET_MODE = "active"  # This script processes only the active sheet in each workbook.
GRAM_NUMBER_FORMAT = "0.000000"


def CellHasUsableText(Value: object) -> bool:
	return isinstance(Value, str) and bool(Value.strip())


def DecimalToExcelNumber(Value: decimal.Decimal) -> float:
	return float(Value)


def AskElementOrIgnore(FilePath: Path, SheetName: str, Row: int, Text: str, Db: dgm_database.DgmDatabase) -> Optional[dgm_database.ElementRecord]:
	print("\nUnknown text found:")
	print(f"  File:  {FilePath}")
	print(f"  Sheet: {SheetName}")
	print(f"  Row:   {Row}")
	print(f"  Text:  {Text}")

	while True:
		Answer = input("Is this a radio element? [e] element / [i] ignore: ").strip().casefold()
		if Answer in ("e", "element"):
			return AddElementInteractively(Db, Text)
		if Answer in ("i", "ignore"):
			return None
		print("Please enter 'e' or 'i'.")


def AddElementInteractively(Db: dgm_database.DgmDatabase, Name: str) -> dgm_database.ElementRecord:
	PathParts = AskStructureParts(Name)
	ParentParts = PathParts[:-1]
	LeafPart = PathParts[-1] if PathParts else Name

	Siblings = Db.GetSiblingInfos(ParentParts)
	if Siblings:
		print("\nExisting sibling nodes under the selected parent:")
		for Sibling in Siblings:
			Marker = "has DGM" if Sibling.HasDgm else "no DGM"
			if Sibling.Kind == "regex":
				print(f"  {Sibling.Index}. regex text='{Sibling.Text}' pattern='{Sibling.Pattern}' ({Marker})")
			else:
				print(f"  {Sibling.Index}. node text='{Sibling.Text}' ({Marker})")

	print("\nAdd mode:")
	print("  n - add exact structured node")
	print("  r - add regex node with newly entered DGM values")
	print("  c - convert an existing sibling node to regex and reuse its DGM values")

	while True:
		Answer = input("Choose add mode [n/r/c, default n]: ").strip().casefold()
		if Answer == "":
			Answer = "n"

		if Answer == "n":
			Values = AskDgmValues(Name)
			Record = Db.AddElement(Name, Values, PathParts)
			Db.Save()
			return Record

		if Answer == "r":
			Pattern = AskRegexPattern(LeafPart)
			DisplayText = input("Display text for this regex node [pattern]: ").strip()
			Values = AskDgmValues(Name)
			Record = Db.AddRegexElement(Name, Values, ParentParts, Pattern, DisplayText)
			Db.Save()
			return Record

		if Answer == "c":
			if not Siblings:
				print("There are no sibling nodes to convert.")
				continue
			Index = AskSiblingIndex(Siblings)
			Pattern = AskRegexPattern(LeafPart)
			DisplayText = input("Display text for this regex node [pattern]: ").strip()
			Record = Db.ConvertSiblingToRegex(ParentParts, Index, Pattern, DisplayText)
			Db.Save()
			return Record

		print("Please enter 'n', 'r', or 'c'.")


def AskStructureParts(Name: str) -> List[str]:
	print("\nStructure path controls how the element is stored in XML.")
	print("Use '/' between node parts. The regular node parts should concatenate to the element name.")
	print("Example: KM-/5b/-M47")
	print("For a regex leaf, enter only the regular parent path plus the current leaf remainder.")

	while True:
		Raw = input(f"Structure path [{Name}]: ").strip()
		if Raw == "":
			return [Name]
		Parts = [Part.strip() for Part in Raw.split("/") if Part.strip()]
		if Parts:
			return Parts
		print("Please enter at least one non-empty path part.")


def AskRegexPattern(DefaultLeafPart: str) -> str:
	print("The regex is matched case-insensitively against the remaining text after the parent path.")
	print(f"Current leaf remainder: {DefaultLeafPart}")
	while True:
		Pattern = input("Regex pattern: ").strip()
		if Pattern:
			return Pattern
		print("Regex pattern cannot be empty.")


def AskSiblingIndex(Siblings: List[dgm_database.SiblingInfo]) -> int:
	ValidIndexes = {Sibling.Index for Sibling in Siblings if Sibling.HasDgm}
	while True:
		Raw = input("Sibling index to convert: ").strip()
		try:
			Value = int(Raw)
		except ValueError:
			print("Please enter an integer sibling index.")
			continue
		if Value not in ValidIndexes:
			print("Please choose a sibling index that exists and has DGM values.")
			continue
		return Value


def AskDgmValues(Name: str) -> dgm_database.DgmValues:
	print(f"Enter DGM content for: {Name}")
	print("Values are entered in milligrams. Empty input means 0 mg.")

	GoldMg = AskNonNegativeDecimal("Gold, mg")
	SilverMg = AskNonNegativeDecimal("Silver, mg")
	PlatinumMg = AskNonNegativeDecimal("Platinum, mg")
	MpgMg = AskNonNegativeDecimal("MPG/palladium, mg")

	return dgm_database.DgmValues(
		GoldG=GoldMg / decimal.Decimal("1000"),
		SilverG=SilverMg / decimal.Decimal("1000"),
		PlatinumG=PlatinumMg / decimal.Decimal("1000"),
		MpgG=MpgMg / decimal.Decimal("1000"),
	)


def AskNonNegativeDecimal(Prompt: str) -> decimal.Decimal:
	while True:
		Raw = input(f"  {Prompt} [0]: ").strip().replace(",", ".")
		if Raw == "":
			return decimal.Decimal("0")
		try:
			Value = decimal.Decimal(Raw)
		except decimal.InvalidOperation:
			print("Please enter a number, for example: 12.5")
			continue
		if Value < 0:
			print("The value cannot be negative.")
			continue
		return Value


def WriteEntryToRow(Sheet: openpyxl.worksheet.worksheet.Worksheet, Row: int, ColumnsInfo: dgm_database.Columns, Entry: dgm_database.ElementRecord) -> None:
	for MetalKey, _ in dgm_database.METALS:
		PerElementCell = Sheet[f"{ColumnsInfo.PerElement[MetalKey]}{Row}"]
		TotalCell = Sheet[f"{ColumnsInfo.Total[MetalKey]}{Row}"]
		PerElementCell.value = DecimalToExcelNumber(Entry.GetMetalValue(MetalKey))
		PerElementCell.number_format = GRAM_NUMBER_FORMAT
		TotalCell.value = f"={ColumnsInfo.PerElement[MetalKey]}{Row}*{ColumnsInfo.Quantity}{Row}"
		TotalCell.number_format = GRAM_NUMBER_FORMAT


def WriteWorkbookTotals(Sheet: openpyxl.worksheet.worksheet.Worksheet, TotalRow: int, ColumnsInfo: dgm_database.Columns, ProcessedRows: List[int]) -> None:
	for MetalKey, _ in dgm_database.METALS:
		TotalColumn = ColumnsInfo.Total[MetalKey]
		Cell = Sheet[f"{TotalColumn}{TotalRow}"]
		Cell.value = BuildSumFormula(TotalColumn, ProcessedRows)
		Cell.number_format = GRAM_NUMBER_FORMAT


def BuildSumFormula(Column: str, Rows: Iterable[int]) -> str:
	SortedRows = sorted(set(Rows))
	if not SortedRows:
		return "=0"

	Ranges: List[str] = []
	Start = SortedRows[0]
	Previous = SortedRows[0]

	for Row in SortedRows[1:]:
		if Row == Previous + 1:
			Previous = Row
			continue
		Ranges.append(FormatRange(Column, Start, Previous))
		Start = Row
		Previous = Row

	Ranges.append(FormatRange(Column, Start, Previous))
	return f"=SUM({','.join(Ranges)})"


def FormatRange(Column: str, Start: int, End: int) -> str:
	if Start == End:
		return f"{Column}{Start}"
	return f"{Column}{Start}:{Column}{End}"


def ProcessWorkbook(FilePath: Path, Db: dgm_database.DgmDatabase) -> Tuple[int, int]:
	Workbook = openpyxl.load_workbook(FilePath, data_only=False)
	Sheet = Workbook.active

	ProcessedRows: List[int] = []
	IgnoredRows: Set[int] = set()
	LastProcessedRow = 0
	ConsecutiveIgnoredRows = 0
	Row = 1
	MaxRow = Sheet.max_row or 1

	while Row <= MaxRow:
		RawName = Sheet[f"{Db.Columns.Name}{Row}"].value

		if not CellHasUsableText(RawName):
			IgnoredRows.add(Row)
			ConsecutiveIgnoredRows += 1
		else:
			Name = " ".join(str(RawName).strip().split())

			if Db.IsIgnoredText(Name):
				IgnoredRows.add(Row)
				ConsecutiveIgnoredRows += 1
			else:
				SearchResult = Db.FindElement(Name)
				Entry = SearchResult.Record
				if Entry is None:
					Entry = AskElementOrIgnore(FilePath, Sheet.title, Row, Name, Db)
					if Entry is None:
						Db.AddIgnoredText(Name)
						Db.Save()
						IgnoredRows.add(Row)
						ConsecutiveIgnoredRows += 1

				if Entry is not None:
					WriteEntryToRow(Sheet, Row, Db.Columns, Entry)
					ProcessedRows.append(Row)
					LastProcessedRow = Row
					ConsecutiveIgnoredRows = 0

		if ConsecutiveIgnoredRows >= STOP_AFTER_CONSECUTIVE_IGNORED_ROWS:
			break

		Row += 1

	TotalRow = LastProcessedRow + 1 if LastProcessedRow > 0 else 1
	WriteWorkbookTotals(Sheet, TotalRow, Db.Columns, ProcessedRows)

	if CREATE_BACKUP:
		BackupPath = FilePath.with_name(FilePath.name + ".bak")
		if not BackupPath.exists():
			shutil.copy2(FilePath, BackupPath)

	Workbook.save(FilePath)
	return len(ProcessedRows), len(IgnoredRows)


def FindXlsxFiles(Folder: Path) -> List[Path]:
	return sorted(
		PathItem
		for PathItem in Folder.iterdir()
		if PathItem.is_file()
		and PathItem.suffix.casefold() == ".xlsx"
		and not PathItem.name.startswith("~$")
	)


def SelectFolderWithDialog() -> Optional[Path]:
	try:
		import tkinter
		import tkinter.filedialog
	except ImportError as Error:
		print(f"Cannot open folder selection dialog because tkinter is unavailable: {Error}", file=sys.stderr)
		return None

	Root = tkinter.Tk()
	Root.withdraw()
	Root.update()

	try:
		SelectedFolder = tkinter.filedialog.askdirectory(title="Select folder with XLSX files")
	finally:
		Root.destroy()

	if not SelectedFolder:
		return None

	return Path(SelectedFolder).expanduser().resolve()


def ResolveFolderArgument(Argument: Optional[str]) -> Optional[Path]:
	if Argument is not None:
		return Path(Argument).expanduser().resolve()

	return SelectFolderWithDialog()


def Main() -> int:
	if len(sys.argv) not in (2, 3):
		print("Usage: python dgm_inventory.py database.xml [folder_with_xlsx_files]", file=sys.stderr)
		return 2

	DatabasePath = Path(sys.argv[1]).expanduser().resolve()
	FolderArgument = sys.argv[2] if len(sys.argv) == 3 else None
	Folder = ResolveFolderArgument(FolderArgument)

	if Folder is None:
		print("No folder selected.", file=sys.stderr)
		return 2

	if not Folder.exists() or not Folder.is_dir():
		print(f"Folder does not exist or is not a directory: {Folder}", file=sys.stderr)
		return 2

	try:
		Db = dgm_database.OpenDatabase(DatabasePath)
	except Exception as Error:
		print(str(Error), file=sys.stderr)
		return 1

	Files = FindXlsxFiles(Folder)
	if not Files:
		print(f"No .xlsx files found in: {Folder}")
		return 0

	print(f"Database: {DatabasePath}")
	print(f"Folder:   {Folder}")
	print(f"Files:    {len(Files)}")
	print(f"Sheet mode: {DEFAULT_SHEET_MODE}")

	TotalProcessed = 0
	TotalIgnored = 0

	for FilePath in Files:
		print(f"\nProcessing: {FilePath.name}")
		try:
			Processed, Ignored = ProcessWorkbook(FilePath, Db)
		except Exception as Error:
			print(f"ERROR while processing '{FilePath}': {Error}", file=sys.stderr)
			continue

		TotalProcessed += Processed
		TotalIgnored += Ignored
		print(f"  Processed element rows: {Processed}")
		print(f"  Ignored rows:           {Ignored}")

	print("\nDone.")
	print(f"Total processed element rows: {TotalProcessed}")
	print(f"Total ignored rows:           {TotalIgnored}")
	return 0


if __name__ == "__main__":
	sys.exit(Main())
