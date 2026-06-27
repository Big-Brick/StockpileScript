from __future__ import annotations

from pathlib import Path
from typing import List, Optional

import tkinter as tk
import tkinter.filedialog
import tkinter.messagebox
import tkinter.ttk as ttk

import dgm_xlsx_common
import dgm_xlsx_postprocessor
import dgm_xlsx_preprocessor
from dgm_gui_common import WINDOW_TITLE, openpyxl


class FooterPlacementDialog(tk.Toplevel):
	def __init__(self, Parent: tk.Toplevel, FilePath: Path, ReviewRows: List[tuple[int, str]], InitialRow: int) -> None:
		super().__init__(Parent)
		self.Result: Optional[int] = None
		self.title(f"Select footer start - {FilePath.name}")
		self.geometry("900x520")
		self.minsize(700, 360)
		self.columnconfigure(0, weight=1)
		self.rowconfigure(1, weight=1)
		ttk.Label(self, text="Review rows from the first possible footer row to the end, then choose the row where the footer must start.", style="Heading.TLabel").grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 4))
		self.Tree = ttk.Treeview(self, columns=("row", "text"), show="headings", selectmode="browse")
		self.Tree.heading("row", text="Row")
		self.Tree.heading("text", text="Row text")
		self.Tree.column("row", width=70, stretch=False, anchor="e")
		self.Tree.column("text", width=760, stretch=True)
		self.Tree.grid(row=1, column=0, sticky="nsew", padx=10, pady=6)
		Scroll = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.Tree.yview)
		Scroll.grid(row=1, column=1, sticky="ns", pady=6)
		self.Tree.configure(yscrollcommand=Scroll.set)
		for Row, Text in ReviewRows:
			self.Tree.insert("", "end", iid=str(Row), values=(Row, Text))
		Controls = ttk.Frame(self)
		Controls.grid(row=2, column=0, sticky="e", padx=10, pady=(4, 10))
		tk.Label(Controls, text="Footer starts at row:").grid(row=0, column=0, padx=(0, 6))
		self.FooterRow = tk.IntVar(value=InitialRow)
		ttk.Spinbox(Controls, from_=1, to=100000, textvariable=self.FooterRow, width=8).grid(row=0, column=1, padx=(0, 6))
		tk.Button(Controls, text="Use selected row", command=self._UseSelected).grid(row=0, column=2, padx=(0, 6))
		tk.Button(Controls, text="OK", command=self._Ok).grid(row=0, column=3, padx=(0, 6))
		tk.Button(Controls, text="Cancel", command=self.destroy).grid(row=0, column=4)
		self.transient(Parent)
		self.grab_set()
		self.wait_window(self)

	def _UseSelected(self) -> None:
		Selection = self.Tree.selection()
		if Selection:
			self.FooterRow.set(int(Selection[0]))

	def _Ok(self) -> None:
		self.Result = int(self.FooterRow.get())
		self.destroy()


class XlsxPostprocessReviewWindow(tk.Toplevel):
	def __init__(self, Parent: tk.Toplevel, Processor: dgm_xlsx_postprocessor.XlsxPostprocessor, Result: dgm_xlsx_postprocessor.PostprocessResult) -> None:
		super().__init__(Parent)
		self.ParentViewer = Parent
		self.Processor = Processor
		self.Result = Result
		self.Issues = list(Result.Issues)
		self.title(f"Postprocess review - {Result.SavedPath.name}")
		self.geometry("1100x620")
		self.minsize(820, 420)
		self.columnconfigure(0, weight=1)
		self.rowconfigure(2, weight=1)

		Meta = Result.Metadata
		Summary = (
			f"Saved: {Result.SavedPath.name}. Type: {Meta.DocumentType or 'unknown'}. "
			f"Number: {Meta.FileNumber if Meta.FileNumber is not None else 'missing'}. "
			f"Equipment: {Meta.EquipmentName or 'unknown'} {Meta.SerialNumber} {Meta.ManufactureYear}. "
			f"Review rows: {len(self.Issues)}. Warnings: {len(Result.Warnings)}."
		)
		ttk.Label(self, text=Summary, style="Heading.TLabel").grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 4))
		if Meta.Conflicts or Result.Warnings:
			Text = "\n".join(Meta.Conflicts + Result.Warnings)
			ttk.Label(self, text=Text, foreground="#8a4b00", wraplength=1050).grid(row=1, column=0, sticky="ew", padx=10, pady=4)

		Columns = ("row", "name", "reason")
		self.Tree = ttk.Treeview(self, columns=Columns, show="headings", selectmode="extended")
		for Column, Label, Width in (("row", "Row", 70), ("name", "Element text", 520), ("reason", "Reason", 300)):
			self.Tree.heading(Column, text=Label)
			self.Tree.column(Column, width=Width, stretch=(Column == "name"))
		self.Tree.grid(row=2, column=0, sticky="nsew", padx=10, pady=6)
		Scroll = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.Tree.yview)
		Scroll.grid(row=2, column=1, sticky="ns", pady=6)
		self.Tree.configure(yscrollcommand=Scroll.set)
		self._PopulateTree()

		Buttons = ttk.Frame(self)
		Buttons.grid(row=3, column=0, sticky="e", padx=10, pady=(4, 10))
		ttk.Button(Buttons, text="Remove rows", command=self._RemoveRows).grid(row=0, column=0, padx=(0, 6))
		ttk.Button(Buttons, text="Set DGM to zero", command=self._SetZero).grid(row=0, column=1, padx=(0, 6))
		ttk.Button(Buttons, text="Інформація відсутня", command=self._InformationMissing).grid(row=0, column=2, padx=(0, 6))
		ttk.Button(Buttons, text="Put formulas", command=self._PutFormulas).grid(row=0, column=3, padx=(0, 6))
		ttk.Button(Buttons, text="Close", command=self.destroy).grid(row=0, column=4)

	def _PopulateTree(self) -> None:
		for Item in self.Tree.get_children(""):
			self.Tree.delete(Item)
		for Index, Issue in enumerate(self.Issues):
			self.Tree.insert("", "end", iid=str(Index), values=(Issue.Row, Issue.Name, Issue.Reason))

	def _SelectedRows(self) -> List[int]:
		Rows: List[int] = []
		for Iid in self.Tree.selection():
			Rows.append(self.Issues[int(Iid)].Row)
		if not Rows:
			tkinter.messagebox.showinfo(WINDOW_TITLE, "Select one or more rows first.", parent=self)
		return Rows

	def _Refresh(self) -> None:
		try:
			Result = self.Processor.ProcessFile(self.Result.SavedPath, self.Result.Metadata.DocumentType, False)
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, str(Error), parent=self)
			return
		self.Result = Result
		self.Issues = list(Result.Issues)
		self._PopulateTree()

	def _RemoveRows(self) -> None:
		Rows = self._SelectedRows()
		if not Rows:
			return
		self.Processor.RemoveRows(self.Result.SavedPath, Rows)
		self._Refresh()

	def _SetZero(self) -> None:
		Rows = self._SelectedRows()
		if not Rows:
			return
		self.Processor.SetRowsToZero(self.Result.SavedPath, Rows)
		self._Refresh()

	def _InformationMissing(self) -> None:
		Rows = self._SelectedRows()
		if not Rows:
			return
		self.Processor.SetRowsInformationMissing(self.Result.SavedPath, Rows)
		self._Refresh()

	def _PutFormulas(self) -> None:
		Rows = self._SelectedRows()
		if not Rows:
			return
		self.Processor.PutRowFormulas(self.Result.SavedPath, Rows)
		self._Refresh()


class XlsxPostprocessingMixin:
	def _SelectAndPostprocessXlsxFile(self) -> None:
		if openpyxl is None:
			tkinter.messagebox.showerror(WINDOW_TITLE, "Missing dependency: openpyxl", parent=self)
			return
		SelectedFile = tkinter.filedialog.askopenfilename(
			title="Select XLSX file to postprocess",
			filetypes=(("Excel workbook", "*.xlsx"), ("All files", "*.*")),
			parent=self,
		)
		if not SelectedFile:
			return
		self._PostprocessXlsxFile(Path(SelectedFile).expanduser().resolve(), RenameToCanonical=True)

	def _SelectAndPostprocessXlsxFolder(self) -> None:
		if openpyxl is None:
			tkinter.messagebox.showerror(WINDOW_TITLE, "Missing dependency: openpyxl", parent=self)
			return
		SelectedFolder = tkinter.filedialog.askdirectory(title="Select folder with XLSX files to postprocess", parent=self)
		if not SelectedFolder:
			return
		Files = dgm_xlsx_common.FindXlsxFiles(Path(SelectedFolder).expanduser().resolve(), False)
		if not Files:
			tkinter.messagebox.showinfo(WINDOW_TITLE, "No .xlsx files found in the selected folder.", parent=self)
			return
		Results = []
		for FilePath in Files:
			try:
				Results.append(self._PostprocessOneWithFooterPrompt(FilePath, True))
			except Exception as Error:
				tkinter.messagebox.showerror(WINDOW_TITLE, f"Cannot postprocess '{FilePath}': {Error}", parent=self)
				return
		tkinter.messagebox.showinfo(WINDOW_TITLE, f"Postprocessed {len(Results)} file(s).", parent=self)

	def _SelectAndPostprocessRegistry(self) -> None:
		if openpyxl is None:
			tkinter.messagebox.showerror(WINDOW_TITLE, "Missing dependency: openpyxl", parent=self)
			return
		RegistryFile = tkinter.filedialog.askopenfilename(
			title="Select registry XLSX file",
			filetypes=(("Excel workbook", "*.xlsx"), ("All files", "*.*")),
			parent=self,
		)
		if not RegistryFile:
			return
		Folder = tkinter.filedialog.askdirectory(title="Select folder with numbered XLSX files", parent=self)
		if not Folder:
			return
		try:
			Count = self._PostprocessRegistry(Path(RegistryFile).expanduser().resolve(), Path(Folder).expanduser().resolve())
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, str(Error), parent=self)
			return
		tkinter.messagebox.showinfo(WINDOW_TITLE, f"Registry postprocessing complete. Processed {Count} file(s).", parent=self)

	def _PostprocessOneWithFooterPrompt(self, FilePath: Path, RenameToCanonical: bool) -> dgm_xlsx_postprocessor.PostprocessResult:
		Processor = self._BuildPostprocessor()
		try:
			return Processor.ProcessFile(FilePath, None, RenameToCanonical)
		except dgm_xlsx_postprocessor.FooterPlacementRequired as Error:
			FooterStart = self._AskFooterStart(FilePath, Error.ReviewStart)
			if FooterStart is None:
				raise RuntimeError("Footer placement was cancelled")
			return Processor.ProcessFile(FilePath, None, RenameToCanonical, FooterStart)

	def _PostprocessXlsxFile(self, FilePath: Path, RenameToCanonical: bool) -> None:
		Processor = self._BuildPostprocessor()
		try:
			Result = Processor.ProcessFile(FilePath, None, RenameToCanonical)
		except dgm_xlsx_postprocessor.FooterPlacementRequired as Error:
			FooterStart = self._AskFooterStart(FilePath, Error.ReviewStart)
			if FooterStart is None:
				return
			try:
				Result = Processor.ProcessFile(FilePath, None, RenameToCanonical, FooterStart)
			except Exception as InnerError:
				tkinter.messagebox.showerror(WINDOW_TITLE, f"Cannot postprocess '{FilePath}': {InnerError}", parent=self)
				return
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, f"Cannot postprocess '{FilePath}': {Error}", parent=self)
			return
		XlsxPostprocessReviewWindow(self, Processor, Result)

	def _AskFooterStart(self, FilePath: Path, ReviewStart: int) -> Optional[int]:
		Rows = self._ReadFooterReviewRows(FilePath, ReviewStart)
		Dialog = FooterPlacementDialog(self, FilePath, Rows, ReviewStart)
		return Dialog.Result

	def _ReadFooterReviewRows(self, FilePath: Path, ReviewStart: int) -> List[tuple[int, str]]:
		Workbook = openpyxl.load_workbook(FilePath, data_only=False)  # type: ignore[union-attr]
		Sheet = Workbook.active
		Rows: List[tuple[int, str]] = []
		for Row in range(max(1, ReviewStart), (Sheet.max_row or 1) + 1):
			Values = []
			for Cell in Sheet[Row]:
				if Cell.value not in (None, ""):
					Values.append(str(Cell.value))
			Rows.append((Row, " | ".join(Values)))
		return Rows

	def _BuildPostprocessor(self) -> dgm_xlsx_postprocessor.XlsxPostprocessor:
		return dgm_xlsx_postprocessor.XlsxPostprocessor(
			self.Database,
			self.DatabasePath.with_name(dgm_xlsx_preprocessor.DEFAULT_RULES_FILENAME),  # type: ignore[name-defined]
		)

	def _PostprocessRegistry(self, RegistryPath: Path, Folder: Path) -> int:
		Workbook = openpyxl.load_workbook(RegistryPath, data_only=False)  # type: ignore[union-attr]
		Sheet = Workbook.active
		Processor = self._BuildPostprocessor()
		FilesByNumber = self._NumberedFiles(Folder)
		Processed = 0
		NextNumber = 1
		for Row in range(3, (Sheet.max_row or 1) + 1):
			NumberValue = Sheet[f"A{Row}"].value
			try:
				Number = int(NumberValue)
			except (TypeError, ValueError):
				continue
			NextNumber = max(NextNumber, Number + 1)
			FilePath = FilesByNumber.get(Number)
			if FilePath is None:
				continue
			try:
				Result = Processor.ProcessFile(FilePath, None, False)
			except dgm_xlsx_postprocessor.FooterPlacementRequired as Error:
				FooterStart = self._AskFooterStart(FilePath, Error.ReviewStart)
				if FooterStart is None:
					raise RuntimeError("Footer placement was cancelled")
				Result = Processor.ProcessFile(FilePath, None, False, FooterStart)
			Metadata = Result.Metadata
			Metadata.FileNumber = Number
			Sheet[f"C{Row}"].value = Metadata.CanonicalTitle()
			Target = FilePath.with_name(Metadata.CanonicalFilename())
			if Target != FilePath:
				FilePath.rename(Target)
			Processed += 1
		Workbook.save(RegistryPath)
		return Processed

	def _NumberedFiles(self, Folder: Path) -> dict[int, Path]:
		Files: dict[int, Path] = {}
		for FilePath in dgm_xlsx_common.FindXlsxFiles(Folder, False):
			Match = dgm_xlsx_postprocessor.FILE_NUMBER_RE.match(FilePath.name)
			if Match:
				Files.setdefault(int(Match.group("number")), FilePath)
		return Files
