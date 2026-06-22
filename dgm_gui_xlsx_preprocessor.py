from __future__ import annotations

from pathlib import Path
from typing import List, Optional

import tkinter.filedialog
import tkinter.messagebox
import tkinter as tk
import tkinter.ttk as ttk

from dgm_gui_common import WINDOW_TITLE, openpyxl
import dgm_inventory
from dgm_xlsx_preprocessor import DEFAULT_RULES_FILENAME, PreprocessChange, PreprocessResult, XlsxPreprocessor


class XlsxPreprocessingMixin:
	def _SelectAndPreprocessXlsxFile(self) -> None:
		if openpyxl is None:
			tkinter.messagebox.showerror(WINDOW_TITLE, "Missing dependency: openpyxl", parent=self)
			return

		SelectedFile = tkinter.filedialog.askopenfilename(
			title="Select XLSX inventory file to preprocess",
			filetypes=(("Excel workbook", "*.xlsx"), ("All files", "*.*")),
			parent=self,
		)
		if not SelectedFile:
			return

		try:
			RulesPath = self.DatabasePath.with_name(DEFAULT_RULES_FILENAME)
			Preprocessor = XlsxPreprocessor(self.Database, RulesPath)
			Result = Preprocessor.PreprocessWorkbook(Path(SelectedFile).expanduser().resolve())
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, f"Cannot preprocess XLSX file: {Error}", parent=self)
			return

		XlsxPreprocessReviewWindow(self, Preprocessor, Result)

	def _SelectAndPreprocessXlsxFolder(self) -> None:
		if openpyxl is None:
			tkinter.messagebox.showerror(WINDOW_TITLE, "Missing dependency: openpyxl", parent=self)
			return

		SelectedFolder = tkinter.filedialog.askdirectory(title="Select folder with XLSX files to preprocess", parent=self)
		if not SelectedFolder:
			return

		Files = dgm_inventory.FindXlsxFiles(
			Path(SelectedFolder).expanduser().resolve(),
			self.ProcessSubfolders.get(),
		)
		if not Files:
			tkinter.messagebox.showinfo(WINDOW_TITLE, "No .xlsx files found in the selected folder.", parent=self)
			return
		self._PreprocessXlsxQueue(Files)

	def _PreprocessXlsxQueue(self, Files: List[Path], Index: int = 0) -> None:
		if Index >= len(Files):
			tkinter.messagebox.showinfo(WINDOW_TITLE, "All selected XLSX files were preprocessed.", parent=self)
			return

		try:
			RulesPath = self.DatabasePath.with_name(DEFAULT_RULES_FILENAME)
			Preprocessor = XlsxPreprocessor(self.Database, RulesPath)
			Result = Preprocessor.PreprocessWorkbook(Files[Index])
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, f"Cannot preprocess '{Files[Index]}': {Error}", parent=self)
			return

		XlsxPreprocessReviewWindow(self, Preprocessor, Result, Files, Index)


class XlsxPreprocessReviewWindow(tk.Toplevel):
	def __init__(self, Parent: tk.Toplevel, Preprocessor: XlsxPreprocessor, Result: PreprocessResult, Files: Optional[List[Path]] = None, Index: int = 0) -> None:
		super().__init__(Parent)
		self.ParentViewer = Parent
		self.Preprocessor = Preprocessor
		self.Result = Result
		self.Files = Files or [Result.FilePath]
		self.Index = Index
		self.Changes = list(Result.ChangedRows)

		self.title(f"XLSX preprocess review - {Result.FilePath.name}")
		self.geometry("1120x620")
		self.minsize(820, 420)
		self.columnconfigure(0, weight=1)
		self.rowconfigure(1, weight=1)

		Summary = (
			f"Found {len(Result.ChangedRows)} proposed changes, "
			f"{len(Result.AmbiguousRows)} ambiguous rows, "
			f"{len(Result.MissingDatabaseMatches)} unverified changed rows. "
			f"Rules: {Result.RulesPath}"
		)
		ttk.Label(self, text=Summary, style="Heading.TLabel").grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 4))

		self.Tree = ttk.Treeview(self, columns=("row", "original", "new", "verified", "notes"), show="headings", selectmode="extended")
		for Column, Label, Width, Stretch in (
			("row", "Row", 70, False),
			("original", "Original text", 260, True),
			("new", "Proposed text", 260, True),
			("verified", "DB", 70, False),
			("notes", "Notes", 360, True),
		):
			self.Tree.heading(Column, text=Label)
			self.Tree.column(Column, width=Width, stretch=Stretch)
		self.Tree.grid(row=1, column=0, sticky="nsew", padx=10, pady=6)

		Scroll = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.Tree.yview)
		Scroll.grid(row=1, column=1, sticky="ns", pady=6)
		self.Tree.configure(yscrollcommand=Scroll.set)

		for Index, Change in enumerate(self.Changes):
			Verified = "yes" if Change.DatabaseVerified else "no"
			if Change.Ambiguous:
				Verified = "ambiguous"
			self.Tree.insert(
				"",
				"end",
				iid=str(Index),
				values=(Change.Row, Change.OriginalText, Change.NewText, Verified, "; ".join(Change.StageNotes)),
			)

		ButtonFrame = ttk.Frame(self)
		ButtonFrame.grid(row=2, column=0, sticky="e", padx=10, pady=(4, 10))
		ttk.Button(ButtonFrame, text="Apply selected", command=self._ApplySelected).grid(row=0, column=0, padx=(0, 6))
		ttk.Button(ButtonFrame, text="Apply all safe", command=self._ApplySafe).grid(row=0, column=1, padx=(0, 6))
		ttk.Button(ButtonFrame, text="Apply all", command=self._ApplyAll).grid(row=0, column=2, padx=(0, 6))
		if self.Index + 1 < len(self.Files):
			ttk.Button(ButtonFrame, text="Preprocess next file", command=self._PreprocessNext).grid(row=0, column=3, padx=(0, 6))
			CloseColumn = 4
		else:
			CloseColumn = 3
		ttk.Button(ButtonFrame, text="Close", command=self.destroy).grid(row=0, column=CloseColumn)

		if not self.Changes:
			tkinter.messagebox.showinfo(WINDOW_TITLE, "No preprocessing changes were found.", parent=self)

	def _SelectedChanges(self) -> List[PreprocessChange]:
		return [self.Changes[int(Iid)] for Iid in self.Tree.selection()]

	def _ApplySelected(self) -> None:
		Changes = self._SelectedChanges()
		if not Changes:
			return
		self._ApplyChanges(Changes)

	def _ApplySafe(self) -> None:
		Changes = [Change for Change in self.Changes if Change.DatabaseVerified and not Change.Ambiguous]
		self._ApplyChanges(Changes)

	def _ApplyAll(self) -> None:
		if self.Result.AmbiguousRows:
			Proceed = tkinter.messagebox.askyesno(
				WINDOW_TITLE,
				"Some rows have ambiguous type detection. Apply all proposed changes anyway?",
				parent=self,
			)
			if not Proceed:
				return
		self._ApplyChanges(self.Changes)

	def _ApplyChanges(self, Changes: List[PreprocessChange]) -> None:
		if not Changes:
			tkinter.messagebox.showinfo(WINDOW_TITLE, "No changes to apply.", parent=self)
			return
		try:
			self.Preprocessor.ApplyChanges(self.Result, Changes)
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, f"Cannot apply preprocessing changes: {Error}", parent=self)
			return
		Message = f"Applied {len(Changes)} preprocessing changes."
		if self.Index + 1 < len(self.Files):
			if tkinter.messagebox.askyesno(WINDOW_TITLE, f"{Message}\n\nPreprocess next file?", parent=self):
				self._PreprocessNext()
				return
		else:
			tkinter.messagebox.showinfo(WINDOW_TITLE, Message, parent=self)
		self.destroy()

	def _PreprocessNext(self) -> None:
		self.destroy()
		self.ParentViewer._PreprocessXlsxQueue(self.Files, self.Index + 1)
