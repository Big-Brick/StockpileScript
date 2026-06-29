from __future__ import annotations

import decimal
from pathlib import Path
from typing import List, Optional

import tkinter as tk
import tkinter.filedialog
import tkinter.messagebox
import tkinter.ttk as ttk

import dgm_database
import dgm_xlsx_common
import dgm_xlsx_preprocessor
from dgm_gui_common import GuiConflictRow, WINDOW_TITLE, openpyxl
from dgm_gui_dialogs import AddElementDialog, RenameElementDialog


def XlsxConflictSortKey(Item: GuiConflictRow) -> tuple[int, str, str, int]:
	ConflictType = 0 if Item.Record is not None and Item.Record.HasDgm else 1
	return (ConflictType, Item.Name.casefold(), str(Item.FilePath).casefold(), Item.Row)


class XlsxConflictReviewWindow(tk.Toplevel):
	def __init__(self, Parent: tk.Toplevel, Files: List[Path], Conflicts: List[GuiConflictRow], IgnoredRows: int) -> None:
		super().__init__(Parent)
		self.ParentViewer = Parent
		self.Files = Files
		self.Conflicts = sorted(Conflicts, key=XlsxConflictSortKey)
		self.IgnoredRows = IgnoredRows

		FileLabel = Files[0].name if len(Files) == 1 else f"{len(Files)} files"
		self.title(f"XLSX/database conflicts - {FileLabel}")
		self.geometry("1180x650")
		self.minsize(860, 430)
		self.columnconfigure(0, weight=1)
		self.rowconfigure(1, weight=1)

		self.SummaryLabel = ttk.Label(self, style="Heading.TLabel")
		self.SummaryLabel.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 4))

		Columns = ("name", "file", "sheet", "row", "details")
		self.Tree = ttk.Treeview(self, columns=Columns, show="headings", selectmode="extended")
		for Column, Label, Width, Stretch in (
			("name", "Element text", 280, True),
			("file", "File", 220, True),
			("sheet", "Sheet", 140, False),
			("row", "Row", 70, False),
			("details", "Conflict", 430, True),
		):
			self.Tree.heading(Column, text=Label)
			self.Tree.column(Column, width=Width, stretch=Stretch, anchor="e" if Column == "row" else "w")
		self.Tree.grid(row=1, column=0, sticky="nsew", padx=10, pady=6)
		Scroll = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.Tree.yview)
		Scroll.grid(row=1, column=1, sticky="ns", pady=6)
		self.Tree.configure(yscrollcommand=Scroll.set)

		Buttons = ttk.Frame(self)
		Buttons.grid(row=2, column=0, sticky="e", padx=10, pady=(4, 10))
		ttk.Button(Buttons, text="Clean selected XLSX row DGM", command=self._CleanSelectedRow).grid(row=0, column=0, padx=(0, 6))
		ttk.Button(Buttons, text="Update database from selected...", command=self._UpdateDatabaseFromSelected).grid(row=0, column=1, padx=(0, 6))
		ttk.Button(Buttons, text="Rename selected in XLSX...", command=self._RenameSelectedInXlsx).grid(row=0, column=2, padx=(0, 6))
		ttk.Button(Buttons, text="Close", command=self.destroy).grid(row=0, column=3)
		self._PopulateTree()

	def _PopulateTree(self) -> None:
		self.Tree.delete(*self.Tree.get_children(""))
		for Item in self.Conflicts:
			self._InsertTreeItem(Item)
		self._UpdateSummaryLabel()

	def _UpdateSummaryLabel(self) -> None:
		self.SummaryLabel.configure(text=f"Conflicts: {len(self.Conflicts)}. Ignored rows: {self.IgnoredRows}.")

	def _TreeItemId(self, Item: GuiConflictRow) -> str:
		return f"conflict:{id(Item)}"

	def _InsertTreeItem(self, Item: GuiConflictRow) -> None:
		Index = self.Conflicts.index(Item)
		self.Tree.insert("", Index, iid=self._TreeItemId(Item), values=(Item.Name, Item.FilePath.name, Item.SheetName, Item.Row, Item.Details))

	def _SelectedConflicts(self) -> List[GuiConflictRow]:
		Selection = set(self.Tree.selection())
		if not Selection:
			return []
		return [Item for Item in self.Conflicts if self._TreeItemId(Item) in Selection]

	def _SelectedConflict(self) -> Optional[GuiConflictRow]:
		Items = self._SelectedConflicts()
		return Items[0] if Items else None

	def _RemoveConflict(self, Item: GuiConflictRow) -> None:
		if Item in self.Conflicts:
			self.Conflicts.remove(Item)
		ItemId = self._TreeItemId(Item)
		if self.Tree.exists(ItemId):
			self.Tree.delete(ItemId)
		self._UpdateSummaryLabel()

	def _CleanSelectedRow(self) -> None:
		Items = self._SelectedConflicts()
		if not Items:
			return
		try:
			for FilePath in sorted({Item.FilePath for Item in Items}, key=str):
				Workbook = openpyxl.load_workbook(FilePath, data_only=False)  # type: ignore[union-attr]
				for Item in [Conflict for Conflict in Items if Conflict.FilePath == FilePath]:
					Sheet = Workbook[Item.SheetName]
					dgm_xlsx_common.ClearDgmCells(Sheet, Item.Row, self.ParentViewer.Database.Columns)
				Workbook.save(FilePath)
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, f"Cannot clean XLSX row: {Error}", parent=self)
			return
		for Item in Items:
			self._RemoveConflict(Item)

	def _UpdateDatabaseFromSelected(self) -> None:
		Items = self._SelectedConflicts()
		if not Items:
			return
		UpdatedItems: List[GuiConflictRow] = []
		try:
			for Item in Items:
				if not self._UpdateDatabaseFromConflict(Item):
					break
				UpdatedItems.append(Item)
			if UpdatedItems:
				self.ParentViewer.Database.Save()
				self.ParentViewer._PopulateDatabaseViews()
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, f"Cannot update database: {Error}", parent=self)
			return
		for Item in UpdatedItems:
			self._RemoveConflict(Item)

	def _UpdateDatabaseFromConflict(self, Item: GuiConflictRow) -> bool:
		if Item.Record is not None and Item.Record.HasDgm:
			Item.Record.SetValues(Item.SheetValues)
		elif Item.Record is not None:
			self.ParentViewer.Database.AddDgmToExistingPath(Item.Name, Item.SheetValues, Item.Record.PathParts)
		else:
			StructuredResult = self.ParentViewer.Database.FindStructuredElement(dgm_database.NormalizeText(Item.Name), Item.Name)
			if StructuredResult.IsEmpty:
				StructuredResult = self.ParentViewer.Database.FindOptionalOnlyPaths()
			Dialog = AddElementDialog(self, Item.Name, StructuredResult, InitialValues=Item.SheetValues)
			if Dialog.Result is None:
				return False
			Result = Dialog.Result
			if Result.Mode == "existing":
				self.ParentViewer.Database.AddDgmToExistingPath(Item.Name, Result.Values, Result.PathParts)
			elif Result.Mode == "regex":
				self.ParentViewer.Database.AddRegexElement(Item.Name, Result.Values, Result.PathParts, Result.RegexText)
			else:
				self.ParentViewer.Database.AddElement(Item.Name, Result.Values, Result.PathParts)
		return True

	def _RenameSelectedInXlsx(self) -> None:
		Items = self._SelectedConflicts()
		if not Items:
			return
		Dialog = RenameElementDialog(self, Items[0].Name)
		if Dialog.Result is None:
			return
		NewName = " ".join(Dialog.Result.strip().split())
		Replacements: List[GuiConflictRow] = []
		try:
			for FilePath in sorted({Item.FilePath for Item in Items}, key=str):
				Workbook = openpyxl.load_workbook(FilePath, data_only=False)  # type: ignore[union-attr]
				for Item in [Conflict for Conflict in Items if Conflict.FilePath == FilePath]:
					Sheet = Workbook[Item.SheetName]
					Sheet[f"{self.ParentViewer.Database.Columns.Name}{Item.Row}"].value = NewName
					Replacement = self.ParentViewer._BuildXlsxConflict(Item.FilePath, Item.SheetName, Item.Row, NewName, Sheet)
					if Replacement is not None:
						Replacements.append(Replacement)
				Workbook.save(FilePath)
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, f"Cannot rename XLSX element: {Error}", parent=self)
			return
		for Item in Items:
			self._RemoveConflict(Item)
		for Replacement in Replacements:
			self.Conflicts.append(Replacement)
		self.Conflicts.sort(key=XlsxConflictSortKey)
		for Replacement in Replacements:
			self._InsertTreeItem(Replacement)
		ReplacementIds = [self._TreeItemId(Replacement) for Replacement in Replacements]
		if ReplacementIds:
			self.Tree.selection_set(*ReplacementIds)
			self.Tree.see(ReplacementIds[0])
		self._UpdateSummaryLabel()


class XlsxConflictsMixin:
	def _SelectAndReviewXlsxConflictsFile(self) -> None:
		if openpyxl is None:
			tkinter.messagebox.showerror(WINDOW_TITLE, "Missing dependency: openpyxl", parent=self)
			return
		SelectedFile = tkinter.filedialog.askopenfilename(title="Select XLSX inventory file", filetypes=(("Excel workbook", "*.xlsx"), ("All files", "*.*")), parent=self)
		if SelectedFile:
			self._OpenXlsxConflictsWindow([Path(SelectedFile).expanduser().resolve()])

	def _SelectAndReviewXlsxConflictsFolder(self) -> None:
		if openpyxl is None:
			tkinter.messagebox.showerror(WINDOW_TITLE, "Missing dependency: openpyxl", parent=self)
			return
		SelectedFolder = tkinter.filedialog.askdirectory(title="Select folder with XLSX files", parent=self)
		if not SelectedFolder:
			return
		Files = dgm_xlsx_common.FindXlsxFiles(Path(SelectedFolder).expanduser().resolve(), self.ProcessSubfolders.get())
		if not Files:
			tkinter.messagebox.showinfo(WINDOW_TITLE, "No .xlsx files found in the selected folder.", parent=self)
			return
		self._OpenXlsxConflictsWindow(Files)

	def _OpenXlsxConflictsWindow(self, Files: List[Path]) -> None:
		try:
			Conflicts, IgnoredRows = self._CollectXlsxConflicts(Files)
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, f"Cannot collect XLSX conflicts: {Error}", parent=self)
			return
		XlsxConflictReviewWindow(self, Files, Conflicts, IgnoredRows)

	def _CollectXlsxConflicts(self, Files: List[Path]) -> tuple[List[GuiConflictRow], int]:
		Preprocessor = dgm_xlsx_preprocessor.XlsxPreprocessor(self.Database, self.DatabasePath.with_name(dgm_xlsx_preprocessor.DEFAULT_RULES_FILENAME))
		Conflicts: List[GuiConflictRow] = []
		IgnoredRows = 0
		for FilePath in Files:
			Workbook = openpyxl.load_workbook(FilePath, data_only=False)  # type: ignore[union-attr]
			Sheet = Workbook.active
			ConsecutiveIgnoredRows = 0
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
						Conflict = self._BuildXlsxConflict(FilePath, Sheet.title, Row, Name, Sheet)
						if Conflict is not None:
							Conflicts.append(Conflict)
						ConsecutiveIgnoredRows = 0
				if ConsecutiveIgnoredRows >= dgm_xlsx_common.STOP_AFTER_CONSECUTIVE_IGNORED_ROWS:
					break
				Row += 1
		return sorted(Conflicts, key=XlsxConflictSortKey), IgnoredRows

	def _BuildXlsxConflict(self, FilePath: Path, SheetName: str, Row: int, Name: str, Sheet: object) -> Optional[GuiConflictRow]:
		SheetValues = self._ReadXlsxRowDgmValues(Sheet, Row)
		if SheetValues.IsZero():
			return None
		SearchResult = self.Database.FindElement(Name)
		Record = SearchResult.Record
		if Record is None or not Record.HasDgm:
			Details = "Database has no DGM data for this XLSX row with non-zero DGM values"
			DatabaseValues = dgm_database.DgmValues(decimal.Decimal("0"), decimal.Decimal("0"), decimal.Decimal("0"), decimal.Decimal("0"))
			return GuiConflictRow(FilePath, SheetName, Row, Name, Record, SheetValues, DatabaseValues, Details)
		Details = self._FormatXlsxDgmDifferences(SheetValues, Record)
		if not Details:
			return None
		DatabaseValues = dgm_database.DgmValues(Record.Values.GoldG, Record.Values.SilverG, Record.Values.PlatinumG, Record.Values.MpgG)
		return GuiConflictRow(FilePath, SheetName, Row, Name, Record, SheetValues, DatabaseValues, Details)

	def _ReadSheetDgmValue(self, Sheet: object, Column: str, Row: int) -> Optional[decimal.Decimal]:
		Value = Sheet[f"{Column}{Row}"].value  # type: ignore[index]
		if Value is None or Value == "":
			return None
		try:
			return dgm_database.ReadDecimal(str(Value))
		except decimal.InvalidOperation:
			return decimal.Decimal("-1")

	def _ReadXlsxRowDgmValues(self, Sheet: object, Row: int) -> dgm_database.DgmValues:
		Values = dgm_database.DgmValues(decimal.Decimal("0"), decimal.Decimal("0"), decimal.Decimal("0"), decimal.Decimal("0"))
		for MetalKey, _ in dgm_database.METALS:
			Value = self._ReadSheetDgmValue(Sheet, self.Database.Columns.PerElement[MetalKey], Row) or decimal.Decimal("0")
			Values.SetMetalValue(MetalKey, Value)
		return Values

	def _FormatXlsxDgmDifferences(self, SheetValues: dgm_database.DgmValues, Record: dgm_database.ElementRecord) -> str:
		Details: List[str] = []
		for MetalKey, MetalName in dgm_database.METALS:
			SheetValue = SheetValues.GetMetalValue(MetalKey)
			DbValue = Record.GetMetalValue(MetalKey)
			if SheetValue != 0 and SheetValue != DbValue:
				Details.append(f"{MetalName}: sheet {dgm_database.DecimalToText(SheetValue)} g, database {dgm_database.DecimalToText(DbValue)} g")
		return "; ".join(Details)
