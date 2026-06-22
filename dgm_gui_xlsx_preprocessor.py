from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

import tkinter.filedialog
import tkinter.messagebox
import tkinter as tk
import tkinter.ttk as ttk

from dgm_gui_common import WINDOW_TITLE, openpyxl
import dgm_inventory
from dgm_xlsx_preprocessor import (
	DEFAULT_RULES_FILENAME,
	PREFIX_SAFETY_LABELS,
	PREFIX_SAFETY_ORDER,
	PREPROCESS_STAGES,
	PreprocessChange,
	PreprocessResult,
	XlsxPreprocessor,
)


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

		self._PreprocessXlsxQueue([Path(SelectedFile).expanduser().resolve()])

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

	def _PreprocessXlsxQueue(self, Files: List[Path], Index: int = 0, StageIndex: int = 0, LogWindow: Optional["PreprocessLogWindow"] = None) -> None:
		if Index >= len(Files):
			if LogWindow is not None:
				LogWindow.Append("All selected XLSX files were preprocessed.")
			else:
				tkinter.messagebox.showinfo(WINDOW_TITLE, "All selected XLSX files were preprocessed.", parent=self)
			return

		if StageIndex >= len(PREPROCESS_STAGES):
			self._PreprocessXlsxQueue(Files, Index + 1, 0, LogWindow)
			return

		Stage = PREPROCESS_STAGES[StageIndex]
		try:
			RulesPath = self.DatabasePath.with_name(DEFAULT_RULES_FILENAME)
			Preprocessor = XlsxPreprocessor(self.Database, RulesPath)
			Result = Preprocessor.PreprocessWorkbook(Files[Index], Stage.Id)
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, f"Cannot preprocess '{Files[Index]}': {Error}", parent=self)
			return

		if not Result.ChangedRows:
			if LogWindow is None:
				LogWindow = PreprocessLogWindow(self)
			LogWindow.Append(f"{Files[Index].name}: skipped {Stage.Name}; no corrections found.")
			self._PreprocessXlsxQueue(Files, Index, StageIndex + 1, LogWindow)
			return

		XlsxPreprocessReviewWindow(self, Preprocessor, Result, Files, Index, StageIndex, LogWindow)


class PreprocessLogWindow(tk.Toplevel):
	def __init__(self, Parent: tk.Toplevel) -> None:
		super().__init__(Parent)
		self.title("XLSX preprocess log")
		self.geometry("640x240")
		self.minsize(420, 180)
		self.columnconfigure(0, weight=1)
		self.rowconfigure(0, weight=1)
		self.Text = tk.Text(self, wrap="word", state="disabled", height=8)
		self.Text.grid(row=0, column=0, sticky="nsew", padx=(10, 0), pady=10)
		Scroll = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.Text.yview)
		Scroll.grid(row=0, column=1, sticky="ns", padx=(0, 10), pady=10)
		self.Text.configure(yscrollcommand=Scroll.set)

	def Append(self, Message: str) -> None:
		self.Text.configure(state="normal")
		self.Text.insert("end", Message + "\n")
		self.Text.see("end")
		self.Text.configure(state="disabled")


class XlsxPreprocessReviewWindow(tk.Toplevel):
	def __init__(self, Parent: tk.Toplevel, Preprocessor: XlsxPreprocessor, Result: PreprocessResult, Files: Optional[List[Path]] = None, Index: int = 0, StageIndex: int = 0, LogWindow: Optional[PreprocessLogWindow] = None) -> None:
		super().__init__(Parent)
		self.ParentViewer = Parent
		self.Preprocessor = Preprocessor
		self.Result = Result
		self.Files = Files or [Result.FilePath]
		self.Index = Index
		self.StageIndex = StageIndex
		self.LogWindow = LogWindow
		self.Changes = self._SortedChanges(list(Result.ChangedRows))
		self.ItemToChange: Dict[str, PreprocessChange] = {}

		self.title(f"XLSX preprocess review - {Result.FilePath.name}")
		self.geometry("1180x680")
		self.minsize(900, 460)
		self.columnconfigure(0, weight=1)
		self.rowconfigure(2, weight=1)

		Summary = (
			f"Stage {self.StageIndex + 1}/{len(PREPROCESS_STAGES)}: {Result.StageName}. "
			f"Found {len(Result.ChangedRows)} proposed corrections. Rules: {Result.RulesPath}"
		)
		ttk.Label(self, text=Summary, style="Heading.TLabel").grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=(10, 4))

		self._BuildPrefixControls()
		self._BuildTree()
		self._PopulateTree()
		self._BuildButtons()

	def _BuildPrefixControls(self) -> None:
		self.PrefixFrame = ttk.LabelFrame(self, text="Manual prefix assignment")
		if self.Result.StageId != "prefix":
			return
		self.PrefixFrame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=4)
		self.PrefixFrame.columnconfigure(1, weight=1)
		ttk.Label(self.PrefixFrame, text="Element type").grid(row=0, column=0, sticky="w", padx=6, pady=6)
		self.TypeChoices = [ElementType.Canonical for ElementType in self.Preprocessor.Rules.ElementTypes]
		self.TypeVar = tk.StringVar(value=self.TypeChoices[0] if self.TypeChoices else "")
		self.TypeCombo = ttk.Combobox(self.PrefixFrame, values=self.TypeChoices, textvariable=self.TypeVar, state="readonly")
		self.TypeCombo.grid(row=0, column=1, sticky="ew", padx=(0, 6), pady=6)
		ttk.Button(self.PrefixFrame, text="Set selected rows", command=self._SetSelectedPrefix).grid(row=0, column=2, padx=(0, 6), pady=6)

	def _BuildTree(self) -> None:
		self.Tree = ttk.Treeview(self, columns=("row", "original", "new", "type", "safety", "notes"), show="tree headings", selectmode="extended")
		self.Tree.heading("#0", text="Group")
		self.Tree.column("#0", width=240, stretch=False)
		for Column, Label, Width, Stretch in (
			("row", "Row", 70, False),
			("original", "Original text", 240, True),
			("new", "Proposed text", 240, True),
			("type", "Element type", 140, False),
			("safety", "Safety", 170, False),
			("notes", "Notes", 300, True),
		):
			self.Tree.heading(Column, text=Label)
			self.Tree.column(Column, width=Width, stretch=Stretch)
		self.Tree.grid(row=2, column=0, sticky="nsew", padx=(10, 0), pady=6)
		Scroll = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.Tree.yview)
		Scroll.grid(row=2, column=1, sticky="ns", padx=(0, 10), pady=6)
		self.Tree.configure(yscrollcommand=Scroll.set)

	def _BuildButtons(self) -> None:
		ButtonFrame = ttk.Frame(self)
		ButtonFrame.grid(row=3, column=0, sticky="e", padx=10, pady=(4, 10))
		ttk.Button(ButtonFrame, text="Apply selected", command=self._ApplySelected).grid(row=0, column=0, padx=(0, 6))
		ttk.Button(ButtonFrame, text="Apply all safe", command=self._ApplySafe).grid(row=0, column=1, padx=(0, 6))
		ttk.Button(ButtonFrame, text="Apply all visible", command=self._ApplyAll).grid(row=0, column=2, padx=(0, 6))
		ttk.Button(ButtonFrame, text="Next stage", command=self._NextStage).grid(row=0, column=3, padx=(0, 6))
		ttk.Button(ButtonFrame, text="Close", command=self.destroy).grid(row=0, column=4)

	def _SortedChanges(self, Changes: List[PreprocessChange]) -> List[PreprocessChange]:
		if self.Result.StageId != "prefix":
			return sorted(Changes, key=lambda Change: (Change.Row, Change.OriginalText.casefold()))
		return sorted(
			Changes,
			key=lambda Change: (
				PREFIX_SAFETY_ORDER.get(Change.SafetyLevel, 99),
				Change.ElementType.casefold(),
				Change.OriginalText.casefold(),
				Change.Row,
			),
		)

	def _GroupKey(self, Change: PreprocessChange) -> str:
		if self.Result.StageId != "prefix":
			return "Corrections"
		SafetyLabel = PREFIX_SAFETY_LABELS.get(Change.SafetyLevel, Change.SafetyLevel or "Other")
		TypeLabel = Change.ElementType or "No type selected"
		return f"{SafetyLabel} / {TypeLabel}"

	def _PopulateTree(self) -> None:
		for Item in self.Tree.get_children(""):
			self.Tree.delete(Item)
		self.ItemToChange.clear()
		Groups: Dict[str, str] = {}
		for Index, Change in enumerate(self.Changes):
			GroupKey = self._GroupKey(Change)
			GroupId = Groups.get(GroupKey)
			if GroupId is None:
				GroupId = f"group:{len(Groups)}"
				Groups[GroupKey] = GroupId
				self.Tree.insert("", "end", iid=GroupId, text=GroupKey, open=True, values=("", "", "", "", "", ""))
			ItemId = f"change:{Index}"
			self.ItemToChange[ItemId] = Change
			self.Tree.insert(
				GroupId,
				"end",
				iid=ItemId,
				text="",
				values=(Change.Row, Change.OriginalText, Change.NewText, Change.ElementType, Change.SafetyLevel, "; ".join(Change.StageNotes)),
			)

	def _SelectedChanges(self) -> List[PreprocessChange]:
		Changes: List[PreprocessChange] = []
		for Iid in self.Tree.selection():
			if Iid in self.ItemToChange:
				Changes.append(self.ItemToChange[Iid])
			else:
				for Child in self.Tree.get_children(Iid):
					Change = self.ItemToChange.get(Child)
					if Change is not None:
						Changes.append(Change)
		return Changes

	def _SetSelectedPrefix(self) -> None:
		TypeName = self.TypeVar.get()
		if not TypeName:
			return
		Selected = self._SelectedChanges()
		if not Selected:
			return
		for Change in Selected:
			BaseText = Change.OriginalText
			Explicit = self.Preprocessor._FindExplicitType(BaseText)
			if Explicit is not None:
				BaseText = Explicit[1]
			Change.NewText = self.Preprocessor._CleanWhitespace(f"{TypeName} {BaseText}")
			Change.ElementType = TypeName
			Change.SafetyLevel = self.Preprocessor._ClassifyPrefixCandidate(Change.NewText)
			Change.DatabaseVerified = Change.SafetyLevel == "safe"
			Change.Ambiguous = Change.SafetyLevel in ("ambiguous", "unidentified")
			Change.StageNotes = [f"Stage 2: manually assigned {TypeName}"]
		self.Changes = self._SortedChanges(self.Changes)
		self._PopulateTree()

	def _ApplySelected(self) -> None:
		Changes = self._SelectedChanges()
		if not Changes:
			return
		self._ApplyChanges(Changes)

	def _ApplySafe(self) -> None:
		Changes = [Change for Change in self.Changes if Change.DatabaseVerified and not Change.Ambiguous]
		self._ApplyChanges(Changes)

	def _ApplyAll(self) -> None:
		if self.Result.StageId == "prefix" and any(Change.SafetyLevel in ("ambiguous", "unidentified") for Change in self.Changes):
			Proceed = tkinter.messagebox.askyesno(
				WINDOW_TITLE,
				"Some rows are ambiguous or unidentified. Apply all visible changes anyway?",
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
		if self.LogWindow is None:
			self.LogWindow = PreprocessLogWindow(self.ParentViewer)
		self.LogWindow.Append(f"{self.Result.FilePath.name}: applied {len(Changes)} corrections for {self.Result.StageName}.")
		self._NextStage()

	def _NextStage(self) -> None:
		self.destroy()
		self.ParentViewer._PreprocessXlsxQueue(self.Files, self.Index, self.StageIndex + 1, self.LogWindow)
