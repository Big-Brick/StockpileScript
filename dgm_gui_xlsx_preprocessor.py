from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import tkinter.filedialog
import tkinter.messagebox
import tkinter.simpledialog
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
	SavePreprocessRules,
	XlsxPreprocessor,
)

@dataclass
class PreprocessChangeOccurrence:
	Result: PreprocessResult
	Change: PreprocessChange


@dataclass
class GroupedPreprocessChange:
	OriginalText: str
	NewText: str
	StageNotes: List[str]
	DatabaseVerified: bool
	Ambiguous: bool
	StageId: str
	ElementType: str
	SafetyLevel: str
	Occurrences: List[PreprocessChangeOccurrence]

	@property
	def OccurrenceCount(self) -> int:
		return len(self.Occurrences)

	def ApplyToOccurrences(self) -> None:
		for Occurrence in self.Occurrences:
			Occurrence.Change.NewText = self.NewText
			Occurrence.Change.StageNotes = list(self.StageNotes)
			Occurrence.Change.DatabaseVerified = self.DatabaseVerified
			Occurrence.Change.Ambiguous = self.Ambiguous
			Occurrence.Change.ElementType = self.ElementType
			Occurrence.Change.SafetyLevel = self.SafetyLevel


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
		if StageIndex >= len(PREPROCESS_STAGES):
			if LogWindow is not None:
				LogWindow.Append("All selected XLSX files were preprocessed.")
			else:
				tkinter.messagebox.showinfo(WINDOW_TITLE, "All selected XLSX files were preprocessed.", parent=self)
			return

		Stage = PREPROCESS_STAGES[StageIndex]
		try:
			RulesPath = self.DatabasePath.with_name(DEFAULT_RULES_FILENAME)
			Preprocessor = XlsxPreprocessor(self.Database, RulesPath)
			Results = [Preprocessor.PreprocessWorkbook(FilePath, Stage.Id) for FilePath in Files]
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, f"Cannot preprocess selected files: {Error}", parent=self)
			return

		GroupedChanges = self._GroupPreprocessChanges(Results)
		if not GroupedChanges:
			if LogWindow is None:
				LogWindow = PreprocessLogWindow(self)
			LogWindow.Append(f"Skipped {Stage.Name}; no corrections found in {len(Files)} selected file(s).")
			self._PreprocessXlsxQueue(Files, 0, StageIndex + 1, LogWindow)
			return

		XlsxPreprocessReviewWindow(self, Preprocessor, Results, GroupedChanges, Files, StageIndex, LogWindow)

	def _GroupPreprocessChanges(self, Results: List[PreprocessResult]) -> List[GroupedPreprocessChange]:
		Groups: Dict[Tuple[str, str], GroupedPreprocessChange] = {}
		for Result in Results:
			for Change in Result.ChangedRows:
				Key = (Change.StageId, Change.OriginalText.casefold())
				Group = Groups.get(Key)
				if Group is None:
					Group = GroupedPreprocessChange(
						OriginalText=Change.OriginalText,
						NewText=Change.NewText,
						StageNotes=list(Change.StageNotes),
						DatabaseVerified=Change.DatabaseVerified,
						Ambiguous=Change.Ambiguous,
						StageId=Change.StageId,
						ElementType=Change.ElementType,
						SafetyLevel=Change.SafetyLevel,
						Occurrences=[],
					)
					Groups[Key] = Group
				Group.Occurrences.append(PreprocessChangeOccurrence(Result, Change))
		return list(Groups.values())


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
	def __init__(self, Parent: tk.Toplevel, Preprocessor: XlsxPreprocessor, Results: List[PreprocessResult], Changes: List[GroupedPreprocessChange], Files: List[Path], StageIndex: int = 0, LogWindow: Optional[PreprocessLogWindow] = None) -> None:
		super().__init__(Parent)
		self.ParentViewer = Parent
		self.Preprocessor = Preprocessor
		self.Results = Results
		self.Result = Results[0]
		self.Files = Files
		self.StageIndex = StageIndex
		self.LogWindow = LogWindow
		self.Changes = self._SortedChanges(Changes)
		self.ItemToChange: Dict[str, GroupedPreprocessChange] = {}

		FileLabel = self.Result.FilePath.name if len(Files) == 1 else f"{len(Files)} files"
		self.title(f"XLSX preprocess review - {FileLabel}")
		self.geometry("1180x680")
		self.minsize(900, 460)
		self.columnconfigure(0, weight=1)
		self.rowconfigure(2, weight=1)

		TotalOccurrences = sum(Change.OccurrenceCount for Change in self.Changes)
		Summary = (
			f"Stage {self.StageIndex + 1}/{len(PREPROCESS_STAGES)}: {self.Result.StageName}. "
			f"Found {len(self.Changes)} unique corrections across {TotalOccurrences} row occurrence(s) in {len(self.Files)} file(s). "
			f"Rules: {self.Result.RulesPath}"
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
		self.Tree = ttk.Treeview(self, columns=("count", "locations", "original", "new", "type", "safety", "notes"), show="tree headings", selectmode="extended")
		self.Tree.heading("#0", text="Group")
		self.Tree.column("#0", width=240, stretch=False)
		for Column, Label, Width, Stretch in (
			("count", "Rows", 70, False),
			("locations", "Files / rows", 220, True),
			("original", "Original text", 240, True),
			("new", "Proposed text", 240, True),
			("type", "Element type", 140, False),
			("safety", "Safety", 170, False),
			("notes", "Notes", 260, True),
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
		ttk.Button(ButtonFrame, text="Edit original in XLSX", command=lambda: self._EditSelectedText("original")).grid(
			row=0,
			column=3,
			padx=(0, 6),
		)
		ttk.Button(ButtonFrame, text="Edit proposed in XLSX", command=lambda: self._EditSelectedText("proposed")).grid(
			row=0,
			column=4,
			padx=(0, 6),
		)
		ttk.Button(ButtonFrame, text="Next stage", command=self._NextStage).grid(row=0, column=5, padx=(0, 6))
		ttk.Button(ButtonFrame, text="Close", command=self.destroy).grid(row=0, column=6)

	def _SortedChanges(self, Changes: List[GroupedPreprocessChange]) -> List[GroupedPreprocessChange]:
		if self.Result.StageId != "prefix":
			return sorted(Changes, key=lambda Change: (self._FirstOccurrenceRow(Change), Change.OriginalText.casefold()))
		return sorted(
			Changes,
			key=lambda Change: (
				PREFIX_SAFETY_ORDER.get(Change.SafetyLevel, 99),
				Change.ElementType.casefold(),
				Change.OriginalText.casefold(),
				self._FirstOccurrenceRow(Change),
			),
		)

	def _FirstOccurrenceRow(self, Change: GroupedPreprocessChange) -> int:
		return min((Occurrence.Change.Row for Occurrence in Change.Occurrences), default=0)

	def _GroupKey(self, Change: GroupedPreprocessChange) -> str:
		if self.Result.StageId != "prefix":
			return "Corrections"
		SafetyLabel = PREFIX_SAFETY_LABELS.get(Change.SafetyLevel, Change.SafetyLevel or "Other")
		TypeLabel = Change.ElementType or "No type selected"
		return f"{SafetyLabel} / {TypeLabel}"

	def _PopulateTree(self) -> None:
		OpenStates = {
			self.Tree.item(Item, "text"): bool(self.Tree.item(Item, "open"))
			for Item in self.Tree.get_children("")
		}
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
				self.Tree.insert(
					"",
					"end",
					iid=GroupId,
					text=GroupKey,
					open=OpenStates.get(GroupKey, True),
					values=("", "", "", "", "", "", ""),
				)
			ItemId = f"change:{Index}"
			self.ItemToChange[ItemId] = Change
			self.Tree.insert(
				GroupId,
				"end",
				iid=ItemId,
				text="",
				values=(
					Change.OccurrenceCount,
					self._FormatLocations(Change),
					Change.OriginalText,
					Change.NewText,
					Change.ElementType,
					Change.SafetyLevel,
					"; ".join(Change.StageNotes),
				),
			)

	def _FormatLocations(self, Change: GroupedPreprocessChange) -> str:
		LocationsByFile: Dict[Path, List[int]] = {}
		for Occurrence in Change.Occurrences:
			LocationsByFile.setdefault(Occurrence.Result.FilePath, []).append(Occurrence.Change.Row)
		Parts = []
		for FilePath, Rows in sorted(LocationsByFile.items(), key=lambda Item: str(Item[0])):
			RowText = ",".join(str(Row) for Row in sorted(Rows))
			Parts.append(f"{FilePath.name}: {RowText}")
		return "; ".join(Parts)

	def _SelectedChanges(self) -> List[GroupedPreprocessChange]:
		Changes: List[GroupedPreprocessChange] = []
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

	def _SelectedSingleChangeForEdit(self) -> Optional[GroupedPreprocessChange]:
		Selected = self._SelectedChanges()
		Unique: Dict[int, GroupedPreprocessChange] = {id(Change): Change for Change in Selected}
		if not Unique:
			tkinter.messagebox.showinfo(WINDOW_TITLE, "Select one correction row to edit.", parent=self)
			return None
		if len(Unique) != 1:
			tkinter.messagebox.showinfo(WINDOW_TITLE, "Select only one correction row to edit.", parent=self)
			return None
		return next(iter(Unique.values()))

	def _EditSelectedText(self, Source: str) -> None:
		Change = self._SelectedSingleChangeForEdit()
		if Change is None:
			return
		if Change.OccurrenceCount > 1:
			Proceed = tkinter.messagebox.askyesno(
				WINDOW_TITLE,
				f"This correction represents {Change.OccurrenceCount} row occurrences. Edit all of them?",
				parent=self,
			)
			if not Proceed:
				return

		InitialValue = Change.OriginalText if Source == "original" else Change.NewText
		EditedText = tkinter.simpledialog.askstring(
			"Edit XLSX text",
			"Text to write to XLSX:",
			initialvalue=InitialValue,
			parent=self,
		)
		if EditedText is None:
			return
		EditedText = " ".join(EditedText.strip().split())
		if not EditedText:
			tkinter.messagebox.showerror(WINDOW_TITLE, "Text cannot be empty.", parent=self)
			return

		try:
			self._WriteTextToOccurrences(Change, EditedText)
			self._RemoveChangeOccurrences(Change)
			self._ReprocessEditedOccurrences(Change, EditedText)
			if self.Preprocessor.CacheChanged:
				SavePreprocessRules(self.Preprocessor.Rules)
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, f"Cannot edit XLSX text: {Error}", parent=self)
			return

		self.Changes = self._SortedChanges(self.ParentViewer._GroupPreprocessChanges(self.Results))
		self._PopulateTree()
		if not self.Changes:
			tkinter.messagebox.showinfo(WINDOW_TITLE, "No corrections remain for this stage.", parent=self)

	def _WriteTextToOccurrences(self, Change: GroupedPreprocessChange, Text: str) -> None:
		OccurrencesByResult: Dict[int, Tuple[PreprocessResult, List[PreprocessChangeOccurrence]]] = {}
		for Occurrence in Change.Occurrences:
			_, ResultOccurrences = OccurrencesByResult.setdefault(id(Occurrence.Result), (Occurrence.Result, []))
			ResultOccurrences.append(Occurrence)

		for Result, Occurrences in OccurrencesByResult.values():
			if openpyxl is None:
				raise RuntimeError("Missing dependency: openpyxl")
			Workbook = openpyxl.load_workbook(Result.FilePath, data_only=False)
			Sheet = Workbook[Result.SheetName]
			for Occurrence in Occurrences:
				Sheet[f"{self.Preprocessor.Database.Columns.Name}{Occurrence.Change.Row}"].value = Text
			Workbook.save(Result.FilePath)

	def _RemoveChangeOccurrences(self, Change: GroupedPreprocessChange) -> None:
		ChangedIds = {id(Occurrence.Change) for Occurrence in Change.Occurrences}
		for Result in self.Results:
			Result.ChangedRows = [RowChange for RowChange in Result.ChangedRows if id(RowChange) not in ChangedIds]
			Result.AmbiguousRows = [RowChange for RowChange in Result.AmbiguousRows if id(RowChange) not in ChangedIds]
			Result.MissingDatabaseMatches = [RowChange for RowChange in Result.MissingDatabaseMatches if id(RowChange) not in ChangedIds]

	def _ReprocessEditedOccurrences(self, Change: GroupedPreprocessChange, Text: str) -> None:
		for Occurrence in Change.Occurrences:
			if self.Preprocessor.IsIgnoredText(Text):
				continue
			NewChange = self.Preprocessor.PreprocessText(Occurrence.Change.Row, Text, Occurrence.Result.StageId)
			if not self.Preprocessor._ShouldOfferChange(NewChange, Occurrence.Result.StageId):
				continue
			Occurrence.Result.ChangedRows.append(NewChange)
			if NewChange.Ambiguous:
				Occurrence.Result.AmbiguousRows.append(NewChange)
			elif not NewChange.DatabaseVerified:
				Occurrence.Result.MissingDatabaseMatches.append(NewChange)

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

	def _ApplyChanges(self, Changes: List[GroupedPreprocessChange]) -> None:
		if not Changes:
			tkinter.messagebox.showinfo(WINDOW_TITLE, "No changes to apply.", parent=self)
			return
		try:
			ChangesByResult: Dict[int, Tuple[PreprocessResult, List[PreprocessChange]]] = {}
			for Change in Changes:
				Change.ApplyToOccurrences()
				for Occurrence in Change.Occurrences:
					_, ResultChanges = ChangesByResult.setdefault(id(Occurrence.Result), (Occurrence.Result, []))
					ResultChanges.append(Occurrence.Change)
			for Result, ResultChanges in ChangesByResult.values():
				self.Preprocessor.ApplyChanges(Result, ResultChanges)
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, f"Cannot apply preprocessing changes: {Error}", parent=self)
			return
		if self.LogWindow is None:
			self.LogWindow = PreprocessLogWindow(self.ParentViewer)
		OccurrenceCount = sum(Change.OccurrenceCount for Change in Changes)
		self.LogWindow.Append(f"Applied {len(Changes)} unique corrections ({OccurrenceCount} row occurrence(s)) for {self.Result.StageName}.")
		AppliedIds = {id(Change) for Change in Changes}
		self.Changes = [Change for Change in self.Changes if id(Change) not in AppliedIds]
		self._PopulateTree()
		if not self.Changes:
			tkinter.messagebox.showinfo(WINDOW_TITLE, "All visible corrections for this stage were applied. Click Next stage when you are ready to continue.", parent=self)

	def _NextStage(self) -> None:
		self.destroy()
		self.ParentViewer._PreprocessXlsxQueue(self.Files, 0, self.StageIndex + 1, self.LogWindow)
