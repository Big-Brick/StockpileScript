from __future__ import annotations

from pathlib import Path
import dataclasses
import decimal
from typing import List, Optional, Set

import tkinter as tk
import tkinter.filedialog
import tkinter.messagebox
import tkinter.ttk as ttk

import dgm_database
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
		ttk.Label(self, text="Review workbook rows, then choose the row where the footer must start. Changing the row number selects and scrolls to that row.", style="Heading.TLabel").grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 4))
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
		self._UpdatingSelection = False
		Controls = ttk.Frame(self)
		Controls.grid(row=2, column=0, sticky="e", padx=10, pady=(4, 10))
		tk.Label(Controls, text="Footer starts at row:").grid(row=0, column=0, padx=(0, 6))
		self.FooterRow = tk.IntVar(value=InitialRow)
		ttk.Spinbox(Controls, from_=1, to=100000, textvariable=self.FooterRow, width=8).grid(row=0, column=1, padx=(0, 6))
		tk.Button(Controls, text="Use selected row", command=self._UseSelected).grid(row=0, column=2, padx=(0, 6))
		tk.Button(Controls, text="OK", command=self._Ok).grid(row=0, column=3, padx=(0, 6))
		tk.Button(Controls, text="Cancel", command=self.destroy).grid(row=0, column=4)
		self.FooterRow.trace_add("write", self._FooterRowChanged)
		self.Tree.bind("<<TreeviewSelect>>", self._TreeSelectionChanged)
		self._SelectFooterRow(InitialRow)
		self.transient(Parent)
		self.grab_set()
		self.wait_window(self)

	def _UseSelected(self) -> None:
		Selection = self.Tree.selection()
		if Selection:
			self.FooterRow.set(int(Selection[0]))

	def _FooterRowChanged(self, *_Args: object) -> None:
		if self._UpdatingSelection:
			return
		try:
			Row = int(self.FooterRow.get())
		except (tk.TclError, ValueError):
			return
		self._SelectFooterRow(Row)

	def _TreeSelectionChanged(self, _Event: object) -> None:
		Selection = self.Tree.selection()
		if not Selection:
			return
		self._UpdatingSelection = True
		try:
			self.FooterRow.set(int(Selection[0]))
		finally:
			self._UpdatingSelection = False

	def _SelectFooterRow(self, Row: int) -> None:
		ItemId = str(Row)
		if not self.Tree.exists(ItemId):
			return
		self.Tree.selection_set(ItemId)
		self.Tree.focus(ItemId)
		self.Tree.see(ItemId)

	def _Ok(self) -> None:
		self.Result = int(self.FooterRow.get())
		self.destroy()


class PostprocessLogWindow(tk.Toplevel):
	def __init__(self, Parent: tk.Toplevel) -> None:
		super().__init__(Parent)
		self.title("XLSX postprocess log")
		self.geometry("760x320")
		self.minsize(520, 220)
		self.Lines: List[str] = []
		self.columnconfigure(0, weight=1)
		self.rowconfigure(0, weight=1)
		self.Text = tk.Text(self, wrap="word", state="disabled", height=10)
		self.Text.grid(row=0, column=0, sticky="nsew", padx=(10, 0), pady=10)
		Scroll = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.Text.yview)
		Scroll.grid(row=0, column=1, sticky="ns", padx=(0, 10), pady=10)
		self.Text.configure(yscrollcommand=Scroll.set)
		Buttons = ttk.Frame(self)
		Buttons.grid(row=1, column=0, sticky="e", padx=10, pady=(0, 10))
		tk.Button(Buttons, text="Save log...", command=self.SaveLog).grid(row=0, column=0, padx=(0, 6))
		tk.Button(Buttons, text="Close", command=self.destroy).grid(row=0, column=1)

	def IsUsable(self) -> bool:
		try:
			return bool(self.winfo_exists()) and bool(self.Text.winfo_exists())
		except tkinter.TclError:
			return False

	def Append(self, Message: str) -> None:
		if not self.IsUsable():
			return
		self.Lines.append(Message)
		self.Text.configure(state="normal")
		self.Text.insert("end", Message + "\n")
		self.Text.see("end")
		self.Text.configure(state="disabled")
		self.update_idletasks()

	def SaveLog(self) -> None:
		SelectedFile = tkinter.filedialog.asksaveasfilename(
			title="Save postprocess log",
			defaultextension=".txt",
			filetypes=(("Text file", "*.txt"), ("All files", "*.*")),
			parent=self,
		)
		if not SelectedFile:
			return
		Path(SelectedFile).write_text("\n".join(self.Lines) + ("\n" if self.Lines else ""), encoding="utf-8")


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


@dataclasses.dataclass
class RegistryResolutionItem:
	Row: int
	Number: int
	RegistryTitle: str
	Reason: str
	SuggestedMetadata: dgm_xlsx_postprocessor.DgmDocumentMetadata


class RegistryResolutionWindow(tk.Toplevel):
	def __init__(self, Parent: tk.Toplevel, RegistryPath: Path, Folder: Path, Items: List[RegistryResolutionItem], Files: List[Path], LogWindow: Optional[PostprocessLogWindow]) -> None:
		super().__init__(Parent)
		self.ParentViewer = Parent
		self.RegistryPath = RegistryPath
		self.Folder = Folder
		self.Items = Items
		self.Files = Files
		self.LogWindow = LogWindow
		self.Processor = Parent._BuildPostprocessor()  # type: ignore[attr-defined]
		self.title("Resolve registry postprocessing conflicts")
		self.geometry("1180x640")
		self.minsize(900, 460)
		self.columnconfigure(0, weight=1)
		self.columnconfigure(1, weight=1)
		self.rowconfigure(1, weight=1)
		ttk.Label(self, text="Select a registry entry and an unprocessed file, approve metadata, then apply it to the registry and workbook.", style="Heading.TLabel").grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=(10, 4))
		self.EntryTree = ttk.Treeview(self, columns=("row", "number", "title", "reason"), show="headings", selectmode="browse")
		for Column, Label, Width in (("row", "Row", 60), ("number", "№", 60), ("title", "Registry entry", 360), ("reason", "Reason", 300)):
			self.EntryTree.heading(Column, text=Label)
			self.EntryTree.column(Column, width=Width, stretch=Column in ("title", "reason"))
		self.EntryTree.grid(row=1, column=0, sticky="nsew", padx=(10, 5), pady=6)
		self.FileTree = ttk.Treeview(self, columns=("file",), show="headings", selectmode="browse")
		self.FileTree.heading("file", text="Unprocessed file")
		self.FileTree.column("file", width=420, stretch=True)
		self.FileTree.grid(row=1, column=1, sticky="nsew", padx=(5, 10), pady=6)
		Form = ttk.Frame(self)
		Form.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=6)
		for Column in range(10):
			Form.columnconfigure(Column, weight=1)
		self.FileNumberVar = tk.StringVar()
		self.DocTypeVar = tk.StringVar(value=dgm_xlsx_postprocessor.DOCUMENT_TYPE_PRESENT)
		self.EquipmentVar = tk.StringVar()
		self.SerialVar = tk.StringVar()
		self.YearVar = tk.StringVar()
		Fields = (("Number", self.FileNumberVar), ("Type", self.DocTypeVar), ("Equipment", self.EquipmentVar), ("Serial", self.SerialVar), ("Year", self.YearVar))
		for Index, (Label, Variable) in enumerate(Fields):
			tk.Label(Form, text=Label).grid(row=0, column=Index * 2, sticky="w")
			if Label == "Type":
				ttk.Combobox(Form, textvariable=Variable, values=(dgm_xlsx_postprocessor.DOCUMENT_TYPE_PRESENT, dgm_xlsx_postprocessor.DOCUMENT_TYPE_MISSING), width=14, state="readonly").grid(row=0, column=Index * 2 + 1, sticky="ew", padx=(0, 6))
			else:
				ttk.Entry(Form, textvariable=Variable).grid(row=0, column=Index * 2 + 1, sticky="ew", padx=(0, 6))
		Buttons = ttk.Frame(self)
		Buttons.grid(row=3, column=0, columnspan=2, sticky="e", padx=10, pady=(4, 10))
		tk.Button(Buttons, text="Apply approved metadata", command=self._Apply).grid(row=0, column=0, padx=(0, 6))
		tk.Button(Buttons, text="Close", command=self.destroy).grid(row=0, column=1)
		self.EntryTree.bind("<<TreeviewSelect>>", lambda _Event: self._Prefill())
		self.FileTree.bind("<<TreeviewSelect>>", lambda _Event: self._Prefill())
		self._Populate()

	def _Populate(self) -> None:
		for Item in self.EntryTree.get_children(""):
			self.EntryTree.delete(Item)
		for Index, Item in enumerate(self.Items):
			self.EntryTree.insert("", "end", iid=str(Index), values=(Item.Row, Item.Number, Item.RegistryTitle, Item.Reason))
		for Item in self.FileTree.get_children(""):
			self.FileTree.delete(Item)
		for Index, FilePath in enumerate(self.Files):
			self.FileTree.insert("", "end", iid=str(Index), values=(FilePath.name,))

	def _SelectedItem(self) -> Optional[RegistryResolutionItem]:
		Selection = self.EntryTree.selection()
		return self.Items[int(Selection[0])] if Selection else None

	def _SelectedFile(self) -> Optional[Path]:
		Selection = self.FileTree.selection()
		return self.Files[int(Selection[0])] if Selection else None

	def _Prefill(self) -> None:
		Item = self._SelectedItem()
		FilePath = self._SelectedFile()
		Metadata = dgm_xlsx_postprocessor.DgmDocumentMetadata()
		if Item is not None:
			Metadata = dataclasses.replace(Item.SuggestedMetadata)
		if FilePath is not None:
			for Field, Value in dgm_xlsx_postprocessor.ParseDocumentText(FilePath.name).items():
				if getattr(Metadata, Field) in (None, ""):
					setattr(Metadata, Field, Value)
		if Item is not None:
			Metadata.FileNumber = Item.Number
		self.FileNumberVar.set("" if Metadata.FileNumber is None else str(Metadata.FileNumber))
		self.DocTypeVar.set(Metadata.DocumentType or dgm_xlsx_postprocessor.DOCUMENT_TYPE_PRESENT)
		self.EquipmentVar.set(Metadata.EquipmentName)
		self.SerialVar.set(Metadata.SerialNumber)
		self.YearVar.set(Metadata.ManufactureYear)

	def _ApprovedMetadata(self) -> dgm_xlsx_postprocessor.DgmDocumentMetadata:
		try:
			Number = int(self.FileNumberVar.get())
		except ValueError as Error:
			raise RuntimeError("File number must be an integer") from Error
		return dgm_xlsx_postprocessor.DgmDocumentMetadata(
			FileNumber=Number,
			DocumentType=self.DocTypeVar.get(),
			EquipmentName=" ".join(self.EquipmentVar.get().split()),
			SerialNumber=dgm_xlsx_postprocessor.CanonicalSerial(self.SerialVar.get()),
			ManufactureYear=" ".join(self.YearVar.get().split()),
		)

	def _Apply(self) -> None:
		Item = self._SelectedItem()
		FilePath = self._SelectedFile()
		if Item is None or FilePath is None:
			tkinter.messagebox.showinfo(WINDOW_TITLE, "Select both a registry entry and a file.", parent=self)
			return
		try:
			Metadata = self._ApprovedMetadata()
			Workbook = openpyxl.load_workbook(self.RegistryPath, data_only=False)  # type: ignore[union-attr]
			Sheet = Workbook.active
			FormularyValues = self.ParentViewer._ReadRegistryDgmValues(Sheet, Item.Row, 4)  # type: ignore[attr-defined]
			Metadata.FormularyValues = FormularyValues
			NewPath = self.Processor.ApplyApprovedMetadata(FilePath, Metadata)
			self.Processor.ApplyFormularyValues(NewPath, FormularyValues)
			Totals = self.Processor.CalculateTotalInProduct(NewPath)
			Sheet[f"A{Item.Row}"].value = Metadata.FileNumber
			Sheet[f"C{Item.Row}"].value = Metadata.CanonicalTitle()
			self.ParentViewer._WriteRegistryDgmValues(Sheet, Item.Row, 8, Totals)  # type: ignore[attr-defined]
			self.ParentViewer._WriteRegistryMissingFormulas(Sheet, Item.Row)  # type: ignore[attr-defined]
			Workbook.save(self.RegistryPath)
		except Exception as Error:
			if self.LogWindow is not None:
				self.LogWindow.Append(f"ERROR resolving registry row {Item.Row}: {Error}")
			tkinter.messagebox.showerror(WINDOW_TITLE, str(Error), parent=self)
			return
		if self.LogWindow is not None:
			self.LogWindow.Append(f"Resolved registry row {Item.Row} with file {NewPath.name}.")
		ItemIndex = self.Items.index(Item)
		FileIndex = self.Files.index(FilePath)
		del self.Items[ItemIndex]
		del self.Files[FileIndex]
		self._Populate()


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
		LogWindow = PostprocessLogWindow(self)
		self._PostprocessXlsxFile(Path(SelectedFile).expanduser().resolve(), RenameToCanonical=True, LogWindow=LogWindow)

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
		LogWindow = PostprocessLogWindow(self)
		LogWindow.Append(f"Found {len(Files)} XLSX file(s) to postprocess in {SelectedFolder}.")
		Results = []
		for Index, FilePath in enumerate(Files, start=1):
			try:
				LogWindow.Append(f"[{Index}/{len(Files)}] Processing {FilePath.name}...")
				Result = self._PostprocessOneWithFooterPrompt(FilePath, True, LogWindow)
				Results.append(Result)
				LogWindow.Append(f"[{Index}/{len(Files)}] Saved {Result.SavedPath.name}; review issues: {len(Result.Issues)}; warnings: {len(Result.Warnings)}.")
			except Exception as Error:
				LogWindow.Append(f"ERROR processing {FilePath.name}: {Error}")
				tkinter.messagebox.showerror(WINDOW_TITLE, f"Cannot postprocess '{FilePath}': {Error}", parent=self)
				return
		LogWindow.Append(f"Finished postprocessing {len(Results)} file(s).")
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
		LogWindow = PostprocessLogWindow(self)
		RegistryPath = Path(RegistryFile).expanduser().resolve()
		FolderPath = Path(Folder).expanduser().resolve()
		try:
			Count, ResolutionItems, UnprocessedFiles = self._PostprocessRegistry(RegistryPath, FolderPath, LogWindow)
		except Exception as Error:
			LogWindow.Append(f"ERROR: {Error}")
			tkinter.messagebox.showerror(WINDOW_TITLE, str(Error), parent=self)
			return
		LogWindow.Append(f"Registry postprocessing complete. Processed {Count} file(s).")
		if ResolutionItems and UnprocessedFiles:
			LogWindow.Append(f"Opening resolver for {len(ResolutionItems)} registry entr(y/ies) and {len(UnprocessedFiles)} unprocessed file(s).")
			RegistryResolutionWindow(self, RegistryPath, FolderPath, ResolutionItems, UnprocessedFiles, LogWindow)
		tkinter.messagebox.showinfo(WINDOW_TITLE, f"Registry postprocessing complete. Processed {Count} file(s).", parent=self)

	def _PostprocessOneWithFooterPrompt(self, FilePath: Path, RenameToCanonical: bool, LogWindow: Optional[PostprocessLogWindow] = None) -> dgm_xlsx_postprocessor.PostprocessResult:
		Processor = self._BuildPostprocessor()
		try:
			return Processor.ProcessFile(FilePath, None, RenameToCanonical)
		except dgm_xlsx_postprocessor.FooterPlacementRequired as Error:
			if LogWindow is not None:
				LogWindow.Append(f"Footer row needs manual selection for {FilePath.name}; suggested row {Error.ReviewStart}.")
			FooterStart = self._AskFooterStart(FilePath, Error.ReviewStart)
			if FooterStart is None:
				raise RuntimeError("Footer placement was cancelled")
			if LogWindow is not None:
				LogWindow.Append(f"Using footer start row {FooterStart} for {FilePath.name}.")
			return Processor.ProcessFile(FilePath, None, RenameToCanonical, FooterStart)

	def _PostprocessXlsxFile(self, FilePath: Path, RenameToCanonical: bool, LogWindow: Optional[PostprocessLogWindow] = None) -> None:
		Processor = self._BuildPostprocessor()
		if LogWindow is not None:
			LogWindow.Append(f"Processing {FilePath}...")
		try:
			Result = Processor.ProcessFile(FilePath, None, RenameToCanonical)
		except dgm_xlsx_postprocessor.FooterPlacementRequired as Error:
			if LogWindow is not None:
				LogWindow.Append(f"Footer row needs manual selection for {FilePath.name}; suggested row {Error.ReviewStart}.")
			FooterStart = self._AskFooterStart(FilePath, Error.ReviewStart)
			if FooterStart is None:
				if LogWindow is not None:
					LogWindow.Append(f"Cancelled footer selection for {FilePath.name}.")
				return
			if LogWindow is not None:
				LogWindow.Append(f"Using footer start row {FooterStart} for {FilePath.name}.")
			try:
				Result = Processor.ProcessFile(FilePath, None, RenameToCanonical, FooterStart)
			except Exception as InnerError:
				if LogWindow is not None:
					LogWindow.Append(f"ERROR processing {FilePath.name}: {InnerError}")
				tkinter.messagebox.showerror(WINDOW_TITLE, f"Cannot postprocess '{FilePath}': {InnerError}", parent=self)
				return
		except Exception as Error:
			if LogWindow is not None:
				LogWindow.Append(f"ERROR processing {FilePath.name}: {Error}")
			tkinter.messagebox.showerror(WINDOW_TITLE, f"Cannot postprocess '{FilePath}': {Error}", parent=self)
			return
		if LogWindow is not None:
			LogWindow.Append(f"Saved {Result.SavedPath.name}; review issues: {len(Result.Issues)}; warnings: {len(Result.Warnings)}.")
		XlsxPostprocessReviewWindow(self, Processor, Result)

	def _AskFooterStart(self, FilePath: Path, ReviewStart: int) -> Optional[int]:
		Rows = self._ReadFooterReviewRows(FilePath, ReviewStart)
		Dialog = FooterPlacementDialog(self, FilePath, Rows, ReviewStart)
		return Dialog.Result

	def _ReadFooterReviewRows(self, FilePath: Path, ReviewStart: int) -> List[tuple[int, str]]:
		Workbook = openpyxl.load_workbook(FilePath, data_only=False)  # type: ignore[union-attr]
		Sheet = Workbook.active
		Rows: List[tuple[int, str]] = []
		for Row in range(1, (Sheet.max_row or 1) + 1):
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

	def _PostprocessRegistry(self, RegistryPath: Path, Folder: Path, LogWindow: Optional[PostprocessLogWindow] = None) -> tuple[int, List[RegistryResolutionItem], List[Path]]:
		if LogWindow is not None:
			LogWindow.Append(f"Opening registry {RegistryPath}.")
		Workbook = openpyxl.load_workbook(RegistryPath, data_only=False)  # type: ignore[union-attr]
		Sheet = Workbook.active
		Processor = self._BuildPostprocessor()
		FilesByNumber = self._NumberedFiles(Folder)
		if LogWindow is not None:
			LogWindow.Append(f"Found {len(FilesByNumber)} numbered XLSX file(s) in {Folder}.")
		Processed = 0
		ProcessedFiles: Set[Path] = set()
		ResolutionItems: List[RegistryResolutionItem] = []
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
				RegistryMetadata = self._RegistryRowMetadata(Sheet, Row)
				ResolutionItems.append(RegistryResolutionItem(Row, Number, str(Sheet[f"C{Row}"].value or ""), "No matching numbered XLSX file", RegistryMetadata))
				if LogWindow is not None:
					LogWindow.Append(f"Registry row {Row}, number {Number}: no matching numbered XLSX file found.")
				continue
			if LogWindow is not None:
				LogWindow.Append(f"Registry row {Row}, number {Number}: processing {FilePath.name}...")
			try:
				Result = Processor.ProcessFile(FilePath, None, False)
			except dgm_xlsx_postprocessor.FooterPlacementRequired as Error:
				if LogWindow is not None:
					LogWindow.Append(f"Footer row needs manual selection for {FilePath.name}; suggested row {Error.ReviewStart}.")
				FooterStart = self._AskFooterStart(FilePath, Error.ReviewStart)
				if FooterStart is None:
					raise RuntimeError("Footer placement was cancelled")
				if LogWindow is not None:
					LogWindow.Append(f"Using footer start row {FooterStart} for {FilePath.name}.")
				Result = Processor.ProcessFile(FilePath, None, False, FooterStart)
			Metadata = Result.Metadata
			Metadata.FileNumber = Number
			RegistryMetadata = self._RegistryRowMetadata(Sheet, Row)
			RegistryConflicts = self._RegistryMetadataConflicts(Metadata, RegistryMetadata)
			HasMetadataConflict = bool(Metadata.Conflicts or RegistryConflicts)
			if HasMetadataConflict:
				Reason = "; ".join(Metadata.Conflicts + RegistryConflicts)
				ResolutionItems.append(RegistryResolutionItem(Row, Number, str(Sheet[f"C{Row}"].value or ""), Reason, Metadata))
			Sheet[f"C{Row}"].value = Metadata.CanonicalTitle()
			Target = FilePath.with_name(Metadata.CanonicalFilename())
			if Target != FilePath:
				FilePath.rename(Target)
				if LogWindow is not None:
					LogWindow.Append(f"Renamed file to {Target.name}.")
				FilePath = Target
			FormularyValues = self._ReadRegistryDgmValues(Sheet, Row, 4)
			if not Metadata.Conflicts and not RegistryConflicts:
				Processor.ApplyFormularyValues(FilePath, FormularyValues)
				Totals = Processor.CalculateTotalInProduct(FilePath)
				self._WriteRegistryDgmValues(Sheet, Row, 8, Totals)
				self._WriteRegistryMissingFormulas(Sheet, Row)
				if LogWindow is not None:
					LogWindow.Append(f"Registry row {Row}: applied formulary values, copied totals, and wrote missing-DGM formulas.")
			else:
				if LogWindow is not None:
					ConflictText = "; ".join(Metadata.Conflicts + RegistryConflicts)
					LogWindow.Append(f"Registry row {Row}: skipped DGM sync because metadata conflicts remain: {ConflictText}")
			Processed += 1
			if not HasMetadataConflict:
				ProcessedFiles.add(FilePath)
			if LogWindow is not None:
				LogWindow.Append(f"Registry row {Row}: finished {FilePath.name}.")
		if LogWindow is not None:
			LogWindow.Append(f"Saving registry {RegistryPath.name}.")
		Workbook.save(RegistryPath)
		UnprocessedFiles = [FilePath for FilePath in dgm_xlsx_common.FindXlsxFiles(Folder, False) if FilePath not in ProcessedFiles]
		return Processed, ResolutionItems, UnprocessedFiles

	def _RegistryRowMetadata(self, Sheet: object, Row: int) -> dgm_xlsx_postprocessor.DgmDocumentMetadata:
		Metadata = dgm_xlsx_postprocessor.DgmDocumentMetadata()
		try:
			Metadata.FileNumber = int(Sheet[f"A{Row}"].value)
		except (TypeError, ValueError):
			Metadata.FileNumber = None
		Values = dgm_xlsx_postprocessor.ParseDocumentText(str(Sheet[f"C{Row}"].value or ""))
		for Field, Value in Values.items():
			setattr(Metadata, Field, Value)
		if Metadata.FileNumber is None and Values.get("FileNumber") is not None:
			Metadata.FileNumber = Values.get("FileNumber")  # type: ignore[assignment]
		return Metadata

	def _RegistryMetadataConflicts(self, FileMetadata: dgm_xlsx_postprocessor.DgmDocumentMetadata, RegistryMetadata: dgm_xlsx_postprocessor.DgmDocumentMetadata) -> List[str]:
		Conflicts: List[str] = []
		for Field in ("FileNumber", "DocumentType", "EquipmentName", "SerialNumber", "ManufactureYear"):
			FileValue = getattr(FileMetadata, Field)
			RegistryValue = getattr(RegistryMetadata, Field)
			if RegistryValue in (None, "") or FileValue in (None, ""):
				continue
			if Field == "SerialNumber":
				FileValue = dgm_xlsx_postprocessor.CanonicalSerial(FileValue)
				RegistryValue = dgm_xlsx_postprocessor.CanonicalSerial(RegistryValue)
			if FileValue != RegistryValue:
				Conflicts.append(f"{Field}: file has {FileValue}, registry has {RegistryValue}")
		return Conflicts

	def _ReadRegistryDgmValues(self, Sheet: object, Row: int, FirstColumn: int) -> dgm_database.DgmValues:
		Values = []
		for Offset, _Metal in enumerate(dgm_database.METALS):
			CellValue = Sheet.cell(Row, FirstColumn + Offset).value
			Value = dgm_xlsx_postprocessor.ValueToDecimal(CellValue)
			Values.append(Value if Value is not None else decimal.Decimal("0"))
		return dgm_database.DgmValues(*Values)

	def _WriteRegistryDgmValues(self, Sheet: object, Row: int, FirstColumn: int, Values: dgm_database.DgmValues) -> None:
		for Offset, (MetalKey, _MetalName) in enumerate(dgm_database.METALS):
			Cell = Sheet.cell(Row, FirstColumn + Offset)
			Cell.value = dgm_xlsx_postprocessor.DecimalToWorkbookNumber(Values.GetMetalValue(MetalKey))
			Cell.number_format = dgm_xlsx_postprocessor.DGM_NUMBER_FORMAT

	def _WriteRegistryMissingFormulas(self, Sheet: object, Row: int) -> None:
		for Offset, _Metal in enumerate(dgm_database.METALS):
			FormularyCell = Sheet.cell(Row, 4 + Offset).coordinate
			TotalCell = Sheet.cell(Row, 8 + Offset).coordinate
			MissingCell = Sheet.cell(Row, 12 + Offset)
			MissingCell.value = f"={FormularyCell}-{TotalCell}"
			MissingCell.number_format = dgm_xlsx_postprocessor.DGM_NUMBER_FORMAT

	def _NumberedFiles(self, Folder: Path) -> dict[int, Path]:
		Files: dict[int, Path] = {}
		for FilePath in dgm_xlsx_common.FindXlsxFiles(Folder, False):
			Match = dgm_xlsx_postprocessor.FILE_NUMBER_RE.match(FilePath.name)
			if Match:
				Files.setdefault(int(Match.group("number")), FilePath)
		return Files
