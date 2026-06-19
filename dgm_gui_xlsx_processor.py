from __future__ import annotations

import decimal
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import tkinter.filedialog
import tkinter.messagebox

import dgm_database
import dgm_inventory
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

	def _SelectAndProcessXlsxFolder(self) -> None:
		if openpyxl is None:
			tkinter.messagebox.showerror(WINDOW_TITLE, "Missing dependency: openpyxl", parent=self)
			return

		SelectedFolder = tkinter.filedialog.askdirectory(title="Select folder with XLSX files", parent=self)
		if not SelectedFolder:
			return

		Files = dgm_inventory.FindXlsxFiles(Path(SelectedFolder).expanduser().resolve())
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
			if not dgm_inventory.CellHasUsableText(RawName):
				IgnoredRows += 1
				ConsecutiveIgnoredRows += 1
			else:
				Name = " ".join(str(RawName).strip().split())
				if self.Database.IsIgnoredText(Name):
					IgnoredRows += 1
					ConsecutiveIgnoredRows += 1
				else:
					SearchResult = self.Database.FindElement(Name)
					Entry = SearchResult.Record
					if Entry is None:
						MissingByKey[(Sheet.title, Row, Name)] = GuiMissingElement(FilePath, Sheet.title, Row, Name)
						ConsecutiveIgnoredRows = 0
					elif self._RowHasConflictingDgmValues(Sheet, Row, Entry):
						Conflicts.append(self._BuildConflict(FilePath, Sheet, Row, Name, Entry))
						ConsecutiveIgnoredRows = 0
					else:
						dgm_inventory.WriteEntryToRow(Sheet, Row, self.Database.Columns, Entry)
						ProcessedRows.append(Row)
						LastProcessedRow = Row
						ConsecutiveIgnoredRows = 0

			if ConsecutiveIgnoredRows >= dgm_inventory.STOP_AFTER_CONSECUTIVE_IGNORED_ROWS:
				break
			Row += 1

		TotalRow = LastProcessedRow + 1 if LastProcessedRow > 0 else 1
		dgm_inventory.WriteWorkbookTotals(Sheet, TotalRow, self.Database.Columns, ProcessedRows)
		Workbook.save(FilePath)
		return GuiProcessResult(FilePath, len(ProcessedRows), IgnoredRows, list(MissingByKey.values()), Conflicts)

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

