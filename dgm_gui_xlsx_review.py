from __future__ import annotations

from pathlib import Path
from typing import List, Optional

import tkinter as tk
import tkinter.messagebox
import tkinter.ttk as ttk

import dgm_database
import dgm_xlsx_common
from dgm_gui_common import GuiConflictRow, GuiMissingElement, GuiProcessResult, WINDOW_TITLE, openpyxl
from dgm_gui_dialogs import AddElementDialog, RenameElementDialog


class XlsxReviewWindow(tk.Toplevel):
	def __init__(self, Parent: tk.Toplevel, Result: GuiProcessResult, Files: List[Path], Index: int) -> None:
		super().__init__(Parent)
		self.ParentViewer = Parent
		self.Result = Result
		self.Files = Files
		self.Index = Index
		self.MissingItems = list(Result.Missing)
		self.ConflictItems = list(Result.Conflicts)
		self.MissingSortColumn = "name"
		self.MissingSortReverse = False

		self.title(f"XLSX review - {Result.FilePath.name}")
		self.geometry("1100x620")
		self.minsize(800, 420)
		self.columnconfigure(0, weight=1)
		self.rowconfigure(1, weight=1)

		ttk.Label(
			self,
			text=(
				f"Processed {Result.ProcessedRows} rows, ignored {Result.IgnoredRows} rows. "
				f"Missing: {len(Result.Missing)}. Conflicts: {len(Result.Conflicts)}."
			),
			style="Heading.TLabel",
		).grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 4))

		Content = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
		Content.grid(row=1, column=0, sticky="nsew", padx=10, pady=6)
		MissingPane = ttk.Frame(Content)
		ConflictPane = ttk.Frame(Content)
		Content.add(MissingPane, weight=1)
		Content.add(ConflictPane, weight=1)
		self._BuildMissingPane(MissingPane)
		self._BuildConflictPane(ConflictPane)

		ButtonFrame = ttk.Frame(self)
		ButtonFrame.grid(row=2, column=0, sticky="e", padx=10, pady=(4, 10))
		ttk.Button(ButtonFrame, text="Process current file again", command=self._ProcessAgain).grid(row=0, column=0, padx=(0, 6))
		if Index + 1 < len(Files):
			ttk.Button(ButtonFrame, text="Process next file", command=self._ProcessNext).grid(row=0, column=1, padx=(0, 6))
		ttk.Button(ButtonFrame, text="Close", command=self.destroy).grid(row=0, column=2)

	def _BuildMissingPane(self, Parent: ttk.Frame) -> None:
		Parent.columnconfigure(0, weight=1)
		Parent.rowconfigure(1, weight=1)
		ttk.Label(Parent, text="Missing database elements", style="Heading.TLabel").grid(row=0, column=0, sticky="w")
		self.MissingTree = ttk.Treeview(Parent, columns=("row", "name"), show="headings", selectmode="browse")
		self.MissingTree.heading("row", text="Row", command=lambda: self.SortMissingTree("row"))
		self.MissingTree.heading("name", text="Element text", command=lambda: self.SortMissingTree("name"))
		self.MissingTree.column("row", width=70, stretch=False, anchor="e")
		self.MissingTree.column("name", width=360, stretch=True)
		self.MissingTree.grid(row=1, column=0, sticky="nsew", pady=4)
		for Index, Item in enumerate(self.MissingItems):
			self.MissingTree.insert("", "end", iid=str(Index), values=(Item.Row, Item.Name))
		Buttons = ttk.Frame(Parent)
		Buttons.grid(row=2, column=0, sticky="ew")
		ttk.Button(Buttons, text="Add to database...", command=self._AddMissingToDatabase).grid(row=0, column=0, padx=(0, 6))
		ttk.Button(Buttons, text="Add to ignore list", command=self._IgnoreMissing).grid(row=0, column=1, padx=(0, 6))
		ttk.Button(Buttons, text="Rename in XLSX...", command=self._RenameMissing).grid(row=0, column=2)
		self.SortMissingTree("name")

	def GetMissingTreeSortKey(self, Iid, Column):
		Value = self.MissingTree.set(Iid, Column)

		if Column == "row":
			try:
				return int(Value)
			except ValueError:
				return 0

		return str(Value).casefold()

	def SortMissingTree(self, Column):
		if self.MissingSortColumn == Column:
			Reverse = not self.MissingSortReverse
		else:
			Reverse = False

		Rows = list(self.MissingTree.get_children(""))

		Rows.sort(
			key=lambda Iid: self.GetMissingTreeSortKey(Iid, Column),
			reverse=Reverse
		)

		for Index, Iid in enumerate(Rows):
			self.MissingTree.move(Iid, "", Index)

		self.MissingSortColumn = Column
		self.MissingSortReverse = Reverse

	def _BuildConflictPane(self, Parent: ttk.Frame) -> None:
		Parent.columnconfigure(0, weight=1)
		Parent.rowconfigure(1, weight=1)
		ttk.Label(Parent, text="Rows conflicting with database", style="Heading.TLabel").grid(row=0, column=0, sticky="w")
		self.ConflictTree = ttk.Treeview(Parent, columns=("row", "name", "details"), show="headings", selectmode="browse")
		for Column, Label, Width in (("row", "Row", 70), ("name", "Element text", 220), ("details", "Conflict", 380)):
			self.ConflictTree.heading(Column, text=Label)
			self.ConflictTree.column(Column, width=Width, stretch=(Column == "details"))
		self.ConflictTree.grid(row=1, column=0, sticky="nsew", pady=4)
		for Index, Item in enumerate(self.ConflictItems):
			self.ConflictTree.insert("", "end", iid=str(Index), values=(Item.Row, Item.Name, Item.Details))
		Buttons = ttk.Frame(Parent)
		Buttons.grid(row=2, column=0, sticky="ew")
		ttk.Button(Buttons, text="Clean XLSX row", command=self._CleanConflictRow).grid(row=0, column=0, padx=(0, 6))
		ttk.Button(Buttons, text="Update database from row", command=self._UpdateDatabaseFromConflict).grid(row=0, column=1)

	def _SelectedMissing(self) -> Optional[GuiMissingElement]:
		Selection = self.MissingTree.selection()
		return self.MissingItems[int(Selection[0])] if Selection else None

	def _SelectedConflict(self) -> Optional[GuiConflictRow]:
		Selection = self.ConflictTree.selection()
		return self.ConflictItems[int(Selection[0])] if Selection else None

	def _AddMissingToDatabase(self) -> None:
		Item = self._SelectedMissing()
		if Item is None:
			return
		StructuredResult = self.ParentViewer.Database.FindStructuredElement(dgm_database.NormalizeText(Item.Name), Item.Name)
		if StructuredResult.IsEmpty:
			StructuredResult = self.ParentViewer.Database.FindOptionalOnlyPaths()
		Dialog = AddElementDialog(self, Item.Name, StructuredResult)
		if Dialog.Result is None:
			return
		try:
			Result = Dialog.Result
			if Result.Mode == "existing":
				self.ParentViewer.Database.AddDgmToExistingPath(Item.Name, Result.Values, Result.PathParts)
			elif Result.Mode == "regex":
				self.ParentViewer.Database.AddRegexElement(Item.Name, Result.Values, Result.PathParts, Result.RegexText)
			else:
				self.ParentViewer.Database.AddElement(Item.Name, Result.Values, Result.PathParts)
			self.ParentViewer.Database.Save()
			self.ParentViewer._PopulateDatabaseViews()
			self.destroy()
			self.ParentViewer._ProcessXlsxQueue(self.Files, self.Index)
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, str(Error), parent=self)

	def _IgnoreMissing(self) -> None:
		Item = self._SelectedMissing()
		if Item is None:
			return
		self.ParentViewer.Database.AddIgnoredText(Item.Name)
		self.ParentViewer.Database.Save()
		self.ParentViewer._PopulateIgnoredList()
		self.destroy()
		self.ParentViewer._ProcessXlsxQueue(self.Files, self.Index)

	def _RenameMissing(self) -> None:
		Item = self._SelectedMissing()
		if Item is None:
			return
		Dialog = RenameElementDialog(self, Item.Name)
		if Dialog.Result is None:
			return
		Workbook = openpyxl.load_workbook(Item.FilePath, data_only=False)  # type: ignore[union-attr]
		Workbook[Item.SheetName][f"{self.ParentViewer.Database.Columns.Name}{Item.Row}"].value = Dialog.Result
		Workbook.save(Item.FilePath)
		self.destroy()
		self.ParentViewer._ProcessXlsxQueue(self.Files, self.Index)

	def _CleanConflictRow(self) -> None:
		Item = self._SelectedConflict()
		if Item is None:
			return
		Workbook = openpyxl.load_workbook(Item.FilePath, data_only=False)  # type: ignore[union-attr]
		Sheet = Workbook[Item.SheetName]
		dgm_xlsx_common.ClearDgmCells(Sheet, Item.Row, self.ParentViewer.Database.Columns)
		Workbook.save(Item.FilePath)
		self.destroy()
		self.ParentViewer._ProcessXlsxQueue(self.Files, self.Index)

	def _UpdateDatabaseFromConflict(self) -> None:
		Item = self._SelectedConflict()
		if Item is None:
			return
		Item.Record.SetValues(Item.SheetValues)
		self.ParentViewer.Database.Save()
		self.ParentViewer._PopulateDatabaseViews()
		self.destroy()
		self.ParentViewer._ProcessXlsxQueue(self.Files, self.Index)

	def _ProcessAgain(self) -> None:
		self.destroy()
		self.ParentViewer._ProcessXlsxQueue(self.Files, self.Index)

	def _ProcessNext(self) -> None:
		self.destroy()
		self.ParentViewer._ProcessXlsxQueue(self.Files, self.Index + 1)

