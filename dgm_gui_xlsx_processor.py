from __future__ import annotations

import decimal
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import tkinter.filedialog
import tkinter.messagebox

import dgm_database
import dgm_xlsx_common
import dgm_xlsx_preprocessor
from dgm_gui_common import GuiConflictRow, GuiMissingElement, GuiProcessResult, WINDOW_TITLE, openpyxl
from dgm_gui_xlsx_review import XlsxReviewWindow


class XlsxProcessingMixin:
	def _SelectAndProcessXlsxFile(self) -> None:
		if openpyxl is None:
			tkinter.messagebox.showerror(WINDOW_TITLE, "Missing dependency: openpyxl", parent=self)
			return

		SelectedFile = tkinter.filedialog.askopenfilename(
			title="Select XLSX inventory file",
			filetypes=(("Excel workbook", "*.xlsx"), ("All files", "*.*")),
			parent=self,
		)
		if not SelectedFile:
			return
		self._ProcessXlsxQueue([Path(SelectedFile).expanduser().resolve()])

	def _SelectAndCleanXlsxFile(self) -> None:
		if openpyxl is None:
			tkinter.messagebox.showerror(WINDOW_TITLE, "Missing dependency: openpyxl", parent=self)
			return

		SelectedFile = tkinter.filedialog.askopenfilename(
			title="Select XLSX inventory file to clean",
			filetypes=(("Excel workbook", "*.xlsx"), ("All files", "*.*")),
			parent=self,
		)
		if not SelectedFile:
			return

		FilePath = Path(SelectedFile).expanduser().resolve()
		try:
			FirstRow, LastRow, ClearedRows = self._CleanXlsxDgmColumns(FilePath)
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, f"Cannot clean '{FilePath}': {Error}", parent=self)
			return

		if ClearedRows == 0:
			tkinter.messagebox.showinfo(WINDOW_TITLE, "No known element rows were found in the selected workbook.", parent=self)
		else:
			tkinter.messagebox.showinfo(
				WINDOW_TITLE,
				f"Cleaned DGM value and total cells for rows {FirstRow}-{LastRow} in '{FilePath.name}'.",
				parent=self,
			)

	def _SelectAndProcessXlsxFolder(self) -> None:
		if openpyxl is None:
			tkinter.messagebox.showerror(WINDOW_TITLE, "Missing dependency: openpyxl", parent=self)
			return

		SelectedFolder = tkinter.filedialog.askdirectory(title="Select folder with XLSX files", parent=self)
		if not SelectedFolder:
			return

		Files = dgm_xlsx_common.FindXlsxFiles(
			Path(SelectedFolder).expanduser().resolve(),
			self.ProcessSubfolders.get(),
		)
		if not Files:
			tkinter.messagebox.showinfo(WINDOW_TITLE, "No .xlsx files found in the selected folder.", parent=self)
			return
		self._ProcessXlsxQueue(Files)

	def _ProcessXlsxQueue(self, Files: List[Path], Index: int = 0) -> None:
		if Index >= len(Files):
			tkinter.messagebox.showinfo(WINDOW_TITLE, "All selected XLSX files were processed.", parent=self)
			return

		try:
			Result = self._ProcessXlsxFile(Files[Index])
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, f"Cannot process '{Files[Index]}': {Error}", parent=self)
			return

		XlsxReviewWindow(self, Result, Files, Index)

	def _ProcessXlsxFile(self, FilePath: Path) -> GuiProcessResult:
		if openpyxl is None:
			raise RuntimeError("Missing dependency: openpyxl")

		Workbook = openpyxl.load_workbook(FilePath, data_only=False)
		Sheet = Workbook.active
		Preprocessor = dgm_xlsx_preprocessor.XlsxPreprocessor(
			self.Database,
			self.DatabasePath.with_name(dgm_xlsx_preprocessor.DEFAULT_RULES_FILENAME),
		)

		ProcessedRows: List[int] = []
		IgnoredRows = 0
		LastProcessedRow = 0
		ConsecutiveIgnoredRows = 0
		MissingByKey: Dict[Tuple[str, int, str], GuiMissingElement] = {}
		Conflicts: List[GuiConflictRow] = []
		Row = 1
		MaxRow = Sheet.max_row or 1

		while Row <= MaxRow:
			RawName = Sheet[f"{self.Database.Columns.Name}{Row}"].value
			if not dgm_xlsx_common.CellHasUsableText(RawName):
				IgnoredRows += 1
				ConsecutiveIgnoredRows += 1
			else:
				Name = " ".join(str(RawName).strip().split())
				if self.Database.IsIgnoredText(Name) or Preprocessor.IsIgnoredText(Name):
					IgnoredRows += 1
					ConsecutiveIgnoredRows += 1
				else:
					SearchResult = self.Database.FindElement(Name)
					Entry = SearchResult.Record
					if Entry is None or not Entry.HasDgm:
						MissingByKey[(Sheet.title, Row, Name)] = GuiMissingElement(FilePath, Sheet.title, Row, Name)
						ConsecutiveIgnoredRows = 0
					elif self._RowHasConflictingDgmValues(Sheet, Row, Entry):
						Conflicts.append(self._BuildConflict(FilePath, Sheet, Row, Name, Entry))
						ConsecutiveIgnoredRows = 0
					else:
						dgm_xlsx_common.WriteEntryToRow(Sheet, Row, self.Database.Columns, Entry)
						ProcessedRows.append(Row)
						LastProcessedRow = Row
						ConsecutiveIgnoredRows = 0

			if ConsecutiveIgnoredRows >= dgm_xlsx_common.STOP_AFTER_CONSECUTIVE_IGNORED_ROWS:
				break
			Row += 1

		if LastProcessedRow > 0:
			TotalRow = LastProcessedRow + 1
			dgm_xlsx_common.WriteWorkbookTotals(Sheet, TotalRow, self.Database.Columns, ProcessedRows)
		Workbook.save(FilePath)
		return GuiProcessResult(FilePath, len(ProcessedRows), IgnoredRows, list(MissingByKey.values()), Conflicts)

	def _CleanXlsxDgmColumns(self, FilePath: Path) -> Tuple[int, int, int]:
		if openpyxl is None:
			raise RuntimeError("Missing dependency: openpyxl")

		Workbook = openpyxl.load_workbook(FilePath, data_only=False)
		Sheet = Workbook.active
		Preprocessor = dgm_xlsx_preprocessor.XlsxPreprocessor(
			self.Database,
			self.DatabasePath.with_name(dgm_xlsx_preprocessor.DEFAULT_RULES_FILENAME),
		)

		KnownRows = self._FindKnownElementRows(Sheet, Preprocessor)
		if not KnownRows:
			return (0, 0, 0)

		FirstRow = min(KnownRows)
		LastRow = max(KnownRows) + 1
		for Row in range(FirstRow, LastRow + 1):
			dgm_xlsx_common.ClearDgmCells(Sheet, Row, self.Database.Columns)

		Workbook.save(FilePath)
		return (FirstRow, LastRow, LastRow - FirstRow + 1)

	def _FindKnownElementRows(self, Sheet: object, Preprocessor: dgm_xlsx_preprocessor.XlsxPreprocessor) -> List[int]:
		KnownRows: List[int] = []
		Row = 1
		MaxRow = Sheet.max_row or 1  # type: ignore[attr-defined]
		ConsecutiveIgnoredRows = 0

		while Row <= MaxRow:
			RawName = Sheet[f"{self.Database.Columns.Name}{Row}"].value  # type: ignore[index]
			if not dgm_xlsx_common.CellHasUsableText(RawName):
				ConsecutiveIgnoredRows += 1
			else:
				Name = " ".join(str(RawName).strip().split())
				if self.Database.IsIgnoredText(Name) or Preprocessor.IsIgnoredText(Name):
					ConsecutiveIgnoredRows += 1
				else:
					SearchResult = self.Database.FindElement(Name)
					if SearchResult.Record is not None and SearchResult.Record.HasDgm:
						KnownRows.append(Row)
					ConsecutiveIgnoredRows = 0

			if ConsecutiveIgnoredRows >= dgm_xlsx_common.STOP_AFTER_CONSECUTIVE_IGNORED_ROWS:
				break
			Row += 1

		return KnownRows

	def _RowHasConflictingDgmValues(self, Sheet: object, Row: int, Entry: dgm_database.ElementRecord) -> bool:
		for MetalKey, _ in dgm_database.METALS:
			Value = self._ReadSheetDgmValue(Sheet, self.Database.Columns.PerElement[MetalKey], Row)
			if Value is not None and Value != 0 and Value != Entry.GetMetalValue(MetalKey):
				return True
		return False

	def _ReadSheetDgmValue(self, Sheet: object, Column: str, Row: int) -> Optional[decimal.Decimal]:
		Value = Sheet[f"{Column}{Row}"].value  # type: ignore[index]
		if Value is None or Value == "":
			return None
		try:
			return dgm_database.ReadDecimal(str(Value))
		except decimal.InvalidOperation:
			return decimal.Decimal("-1")

	def _BuildConflict(self, FilePath: Path, Sheet: object, Row: int, Name: str, Entry: dgm_database.ElementRecord) -> GuiConflictRow:
		SheetValues = dgm_database.DgmValues(decimal.Decimal("0"), decimal.Decimal("0"), decimal.Decimal("0"), decimal.Decimal("0"))
		Details: List[str] = []
		for MetalKey, MetalName in dgm_database.METALS:
			Value = self._ReadSheetDgmValue(Sheet, self.Database.Columns.PerElement[MetalKey], Row) or decimal.Decimal("0")
			SheetValues.SetMetalValue(MetalKey, Value)
			DbValue = Entry.GetMetalValue(MetalKey)
			if Value != 0 and Value != DbValue:
				Details.append(f"{MetalName}: sheet {dgm_database.DecimalToText(Value)} g, database {dgm_database.DecimalToText(DbValue)} g")
		DatabaseValues = dgm_database.DgmValues(Entry.Values.GoldG, Entry.Values.SilverG, Entry.Values.PlatinumG, Entry.Values.MpgG)
		return GuiConflictRow(FilePath, str(Sheet.title), Row, Name, Entry, SheetValues, DatabaseValues, "; ".join(Details))
