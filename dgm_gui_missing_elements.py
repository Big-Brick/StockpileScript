from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

import tkinter as tk
import tkinter.filedialog
import tkinter.messagebox
import tkinter.simpledialog
import tkinter.ttk as ttk

import dgm_database
import dgm_xlsx_common
import dgm_xlsx_preprocessor
from dgm_gui_common import GuiMissingElement, WINDOW_TITLE, openpyxl
from dgm_gui_dialogs import AddElementDialog


@dataclass
class MissingElementGroup:
	Key: str
	Title: str
	SortText: str
	ItemsByName: Dict[str, "MissingElementSummary"] = field(default_factory=dict)


@dataclass
class MissingElementSummary:
	Name: str
	Occurrences: List[GuiMissingElement] = field(default_factory=list)


class MissingElementsWindow(tk.Toplevel):
	def __init__(self, Parent: tk.Toplevel, Files: List[Path], Groups: List[MissingElementGroup], IgnoredRows: int) -> None:
		super().__init__(Parent)
		self.ParentViewer = Parent
		self.Files = Files
		self.Groups = Groups
		self.IgnoredRows = IgnoredRows
		self.TreeItems: Dict[str, Tuple[MissingElementGroup, Optional[MissingElementSummary]]] = {}
		self.OccurrenceItems: Dict[str, Tuple[MissingElementGroup, MissingElementSummary, GuiMissingElement]] = {}

		FileLabel = Files[0].name if len(Files) == 1 else f"{len(Files)} files"
		self.title(f"Missing elements - {FileLabel}")
		self.geometry("1050x650")
		self.minsize(780, 420)
		self.columnconfigure(0, weight=1)
		self.rowconfigure(1, weight=1)

		self.SummaryLabel = ttk.Label(
			self,
			style="Heading.TLabel",
		)
		self.SummaryLabel.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 4))
		self._UpdateSummaryLabel()

		Columns = ("count", "files", "first_location")
		self.Tree = ttk.Treeview(self, columns=Columns, show="tree headings", selectmode="browse")
		self.Tree.heading("#0", text="Database partial match / element text")
		self.Tree.heading("count", text="Count")
		self.Tree.heading("files", text="Files")
		self.Tree.heading("first_location", text="First location")
		self.Tree.column("#0", width=500, minwidth=260, stretch=True)
		self.Tree.column("count", width=70, stretch=False, anchor="e")
		self.Tree.column("files", width=240, stretch=True)
		self.Tree.column("first_location", width=180, stretch=False)
		self.Tree.grid(row=1, column=0, sticky="nsew", padx=10, pady=6)
		Scroll = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.Tree.yview)
		Scroll.grid(row=1, column=1, sticky="ns", pady=6)
		self.Tree.configure(yscrollcommand=Scroll.set)
		self._PopulateTree()

		Buttons = ttk.Frame(self)
		Buttons.grid(row=2, column=0, sticky="e", padx=10, pady=(4, 10))
		ttk.Button(Buttons, text="Add selected to database...", command=self._AddSelectedToDatabase).grid(row=0, column=0, padx=(0, 6))
		ttk.Button(Buttons, text="Add selected to ignore list", command=self._IgnoreSelected).grid(row=0, column=1, padx=(0, 6))
		ttk.Button(Buttons, text="Edit selected in XLSX", command=self._EditSelectedInXlsx).grid(row=0, column=2, padx=(0, 6))
		ttk.Button(Buttons, text="Dump no-partial originals", command=self._DumpNoPartialOriginals).grid(row=0, column=3, padx=(0, 6))
		ttk.Button(Buttons, text="Close", command=self.destroy).grid(row=0, column=4)

	def _UpdateSummaryLabel(self) -> None:
		MissingCount = sum(len(Group.ItemsByName) for Group in self.Groups)
		OccurrenceCount = sum(len(Item.Occurrences) for Group in self.Groups for Item in Group.ItemsByName.values())
		self.SummaryLabel.configure(text=f"Missing unique elements: {MissingCount}. Occurrences: {OccurrenceCount}. Ignored rows: {self.IgnoredRows}.")

	def _PopulateTree(self, PreserveExpansion: bool = False) -> None:
		ExpandedItems = self._GetExpandedItems() if PreserveExpansion else set()
		for Existing in self.Tree.get_children(""):
			self.Tree.delete(Existing)
		self.TreeItems.clear()
		self.OccurrenceItems.clear()
		self._UpdateSummaryLabel()
		for Group in sorted(self.Groups, key=lambda Group: Group.SortText.casefold()):
			GroupOccurrenceCount = sum(len(Item.Occurrences) for Item in Group.ItemsByName.values())
			GroupFiles = self._FormatFiles([Occurrence for Item in Group.ItemsByName.values() for Occurrence in Item.Occurrences])
			GroupId = self._GroupTreeId(Group)
			self.Tree.insert("", "end", iid=GroupId, text=Group.Title, values=(GroupOccurrenceCount, GroupFiles, ""), open=(GroupId in ExpandedItems))
			self.TreeItems[GroupId] = (Group, None)
			for Item in sorted(Group.ItemsByName.values(), key=lambda Item: Item.Name.casefold()):
				First = Item.Occurrences[0]
				ItemId = self._SummaryTreeId(Group, Item)
				self.Tree.insert(GroupId, "end", iid=ItemId, text=Item.Name, values=(len(Item.Occurrences), self._FormatFiles(Item.Occurrences), self._FormatLocation(First)), open=(ItemId in ExpandedItems))
				self.TreeItems[ItemId] = (Group, Item)
				for OccurrenceIndex, Occurrence in enumerate(Item.Occurrences):
					OccurrenceId = self._OccurrenceTreeId(Group, Item, OccurrenceIndex)
					self.Tree.insert(ItemId, "end", iid=OccurrenceId, text=self._FormatLocation(Occurrence), values=("", Occurrence.FilePath.name, ""), open=False)
					self.OccurrenceItems[OccurrenceId] = (Group, Item, Occurrence)

	def _GetExpandedItems(self) -> set[str]:
		Expanded: set[str] = set()

		def Collect(ItemId: str) -> None:
			if self.Tree.item(ItemId, "open"):
				Expanded.add(ItemId)
			for ChildId in self.Tree.get_children(ItemId):
				Collect(ChildId)

		for RootId in self.Tree.get_children(""):
			Collect(RootId)
		return Expanded

	def _GroupTreeId(self, Group: MissingElementGroup) -> str:
		return f"group:{Group.Key}"

	def _SummaryTreeId(self, Group: MissingElementGroup, Summary: MissingElementSummary) -> str:
		return f"item:{Group.Key}:{dgm_database.NormalizeText(Summary.Name)}"

	def _OccurrenceTreeId(self, Group: MissingElementGroup, Summary: MissingElementSummary, OccurrenceIndex: int) -> str:
		return f"occ:{Group.Key}:{dgm_database.NormalizeText(Summary.Name)}:{OccurrenceIndex}"

	def _FormatFiles(self, Occurrences: List[GuiMissingElement]) -> str:
		Names = sorted({Occurrence.FilePath.name for Occurrence in Occurrences}, key=str.casefold)
		return ", ".join(Names[:3]) + (f" (+{len(Names) - 3})" if len(Names) > 3 else "")

	def _FormatLocation(self, Occurrence: GuiMissingElement) -> str:
		return f"{Occurrence.FilePath.name} / {Occurrence.SheetName} / row {Occurrence.Row}"

	def _DumpNoPartialOriginals(self) -> None:
		Items = self._ItemsWithoutUsablePartialMatches()
		if not Items:
			tkinter.messagebox.showinfo(WINDOW_TITLE, "No missing rows without usable database partial matches to dump.", parent=self)
			return
		SelectedFile = tkinter.filedialog.asksaveasfilename(
			title="Save missing original texts",
			defaultextension=".txt",
			filetypes=(("Text file", "*.txt"), ("All files", "*.*")),
			parent=self,
		)
		if not SelectedFile:
			return
		OutputPath = Path(SelectedFile).expanduser().resolve()
		Lines = []
		for Item in Items:
			FirstOccurrence = Item.Occurrences[0] if Item.Occurrences else None
			FileName = FirstOccurrence.FilePath.name if FirstOccurrence is not None else ""
			Lines.append(f"{FileName}\t{Item.Name}")
		OutputPath.write_text("\n".join(Lines) + "\n", encoding="utf-8")
		tkinter.messagebox.showinfo(WINDOW_TITLE, f"Dumped {len(Lines)} missing original text row(s) to '{OutputPath}'.", parent=self)

	def _ItemsWithoutUsablePartialMatches(self) -> List[MissingElementSummary]:
		Items: List[MissingElementSummary] = []
		for Group in self.Groups:
			for Item in Group.ItemsByName.values():
				if not self._HasUsablePartialMatch(Item.Name):
					Items.append(Item)
		return sorted(Items, key=lambda Item: Item.Name.casefold())

	def _HasUsablePartialMatch(self, Text: str) -> bool:
		SearchResult = self.ParentViewer.Database.FindElement(Text)
		Matches = list(SearchResult.PartialMatches)
		if SearchResult.Record is not None and SearchResult.Record.Node.tag == "node":
			Matches.insert(0, dgm_database.PartialElementMatch(Record=SearchResult.Record))
		return any(not self._MatchUsesUnconsumedOptionalNode(Match) for Match in Matches)

	def _MatchUsesUnconsumedOptionalNode(self, Match: dgm_database.PartialElementMatch) -> bool:
		return any(
			dgm_database.IsOptionalNode(Record.Node) and not Record.ConsumedText
			for Record in Match.Record.IterPath()
		)

	def _SelectedItem(self) -> Optional[MissingElementSummary]:
		Selection = self.Tree.selection()
		if not Selection:
			return None
		ItemId = Selection[0]
		while ItemId and ItemId not in self.TreeItems:
			ItemId = self.Tree.parent(ItemId)
		return self.TreeItems.get(ItemId, (None, None))[1]  # type: ignore[return-value]

	def _AddSelectedToDatabase(self) -> None:
		Item = self._SelectedItem()
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
			self._RemoveResolvedSummaries()
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, str(Error), parent=self)

	def _IgnoreSelected(self) -> None:
		Item = self._SelectedItem()
		if Item is None:
			return
		self.ParentViewer.Database.AddIgnoredText(Item.Name)
		self.ParentViewer.Database.Save()
		self.ParentViewer._PopulateIgnoredList()
		self._RemoveSummary(Item)

	def _SelectedOccurrencesForEdit(self) -> Tuple[Optional[MissingElementSummary], List[GuiMissingElement]]:
		Selection = self.Tree.selection()
		if not Selection:
			tkinter.messagebox.showinfo(WINDOW_TITLE, "Select a missing element row to edit.", parent=self)
			return None, []
		ItemId = Selection[0]
		OccurrenceInfo = self.OccurrenceItems.get(ItemId)
		if OccurrenceInfo is not None:
			_, Summary, Occurrence = OccurrenceInfo
			return Summary, [Occurrence]
		Summary = self._SelectedItem()
		if Summary is None:
			tkinter.messagebox.showinfo(WINDOW_TITLE, "Select a missing element row to edit.", parent=self)
			return None, []
		return Summary, list(Summary.Occurrences)

	def _EditSelectedInXlsx(self) -> None:
		Summary, Occurrences = self._SelectedOccurrencesForEdit()
		if Summary is None or not Occurrences:
			return
		if len(Occurrences) > 1:
			Proceed = tkinter.messagebox.askyesno(
				WINDOW_TITLE,
				f"This missing element has {len(Occurrences)} row occurrences. Edit all of them?",
				parent=self,
			)
			if not Proceed:
				return

		EditedText = tkinter.simpledialog.askstring(
			"Edit XLSX text",
			"Text to write to XLSX:",
			initialvalue=Summary.Name,
			parent=self,
		)
		if EditedText is None:
			return
		EditedText = " ".join(EditedText.strip().split())
		if not EditedText:
			tkinter.messagebox.showerror(WINDOW_TITLE, "Text cannot be empty.", parent=self)
			return

		try:
			self._WriteOccurrencesToXlsx(Occurrences, EditedText)
			self._RemoveOccurrences(Occurrences)
			self._RecheckEditedOccurrences(Occurrences, EditedText)
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, f"Cannot edit XLSX text: {Error}", parent=self)
			return
		self._PopulateTree(PreserveExpansion=True)

	def _WriteOccurrencesToXlsx(self, Occurrences: List[GuiMissingElement], Text: str) -> None:
		if openpyxl is None:
			raise RuntimeError("Missing dependency: openpyxl")
		OccurrencesByFile: Dict[Path, List[GuiMissingElement]] = {}
		for Occurrence in Occurrences:
			OccurrencesByFile.setdefault(Occurrence.FilePath, []).append(Occurrence)
		for FilePath, FileOccurrences in OccurrencesByFile.items():
			Workbook = openpyxl.load_workbook(FilePath, data_only=False)  # type: ignore[union-attr]
			SheetsByName = {Sheet.title: Sheet for Sheet in Workbook.worksheets}
			for Occurrence in FileOccurrences:
				Sheet = SheetsByName[Occurrence.SheetName]
				Sheet[f"{self.ParentViewer.Database.Columns.Name}{Occurrence.Row}"].value = Text
			Workbook.save(FilePath)

	def _RemoveOccurrences(self, Occurrences: List[GuiMissingElement]) -> None:
		OccurrenceIds = {id(Occurrence) for Occurrence in Occurrences}
		for Group in self.Groups:
			for Key, Item in list(Group.ItemsByName.items()):
				Item.Occurrences = [Occurrence for Occurrence in Item.Occurrences if id(Occurrence) not in OccurrenceIds]
				if not Item.Occurrences:
					del Group.ItemsByName[Key]
		self.Groups = [Group for Group in self.Groups if Group.ItemsByName]

	def _RecheckEditedOccurrences(self, Occurrences: List[GuiMissingElement], Text: str) -> None:
		Preprocessor = dgm_xlsx_preprocessor.XlsxPreprocessor(
			self.ParentViewer.Database,
			self.ParentViewer.DatabasePath.with_name(dgm_xlsx_preprocessor.DEFAULT_RULES_FILENAME),
		)
		if self.ParentViewer.Database.IsIgnoredText(Text) or Preprocessor.IsIgnoredText(Text):
			return
		SearchResult = self.ParentViewer.Database.FindElement(Text)
		if SearchResult.Record is not None and SearchResult.Record.HasDgm:
			return
		Matches = list(SearchResult.PartialMatches)
		if SearchResult.Record is not None and SearchResult.Record.Node.tag == "node":
			Matches.insert(0, dgm_database.PartialElementMatch(
				Record=SearchResult.Record,
			))
		GroupsByKey = {Group.Key: Group for Group in self.Groups}
		for Occurrence in Occurrences:
			self.ParentViewer._AddMissingElement(
				GroupsByKey,
				GuiMissingElement(Occurrence.FilePath, Occurrence.SheetName, Occurrence.Row, Text),
				Matches,
			)
		self.Groups = list(GroupsByKey.values())

	def _RemoveSummary(self, Summary: MissingElementSummary) -> None:
		self._RemoveSummaries(lambda Item: Item is Summary)

	def _RemoveResolvedSummaries(self) -> None:
		self._RemoveSummaries(lambda Item: (lambda Record: Record is not None and Record.HasDgm)(self.ParentViewer.Database.FindElement(Item.Name).Record))

	def _RemoveSummaries(self, ShouldRemove: Callable[[MissingElementSummary], bool]) -> None:
		for Group in self.Groups:
			for Key, Item in list(Group.ItemsByName.items()):
				if ShouldRemove(Item):
					del Group.ItemsByName[Key]
		self.Groups = [Group for Group in self.Groups if Group.ItemsByName]
		self._PopulateTree(PreserveExpansion=True)


class MissingElementsMixin:
	def _SelectAndListMissingXlsxFile(self) -> None:
		if openpyxl is None:
			tkinter.messagebox.showerror(WINDOW_TITLE, "Missing dependency: openpyxl", parent=self)
			return
		SelectedFile = tkinter.filedialog.askopenfilename(title="Select XLSX inventory file", filetypes=(("Excel workbook", "*.xlsx"), ("All files", "*.*")), parent=self)
		if SelectedFile:
			self._OpenMissingElementsWindow([Path(SelectedFile).expanduser().resolve()])

	def _SelectAndListMissingXlsxFolder(self) -> None:
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
		self._OpenMissingElementsWindow(Files)

	def _OpenMissingElementsWindow(self, Files: List[Path]) -> None:
		try:
			Groups, IgnoredRows = self._CollectMissingElements(Files)
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, f"Cannot collect missing elements: {Error}", parent=self)
			return
		MissingElementsWindow(self, Files, Groups, IgnoredRows)

	def _CollectMissingElements(self, Files: List[Path]) -> Tuple[List[MissingElementGroup], int]:
		Preprocessor = dgm_xlsx_preprocessor.XlsxPreprocessor(self.Database, self.DatabasePath.with_name(dgm_xlsx_preprocessor.DEFAULT_RULES_FILENAME))
		GroupsByKey: Dict[str, MissingElementGroup] = {}
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
						SearchResult = self.Database.FindElement(Name)
						if SearchResult.Record is None or not SearchResult.Record.HasDgm:
							Matches = list(SearchResult.PartialMatches)
							if SearchResult.Record is not None and SearchResult.Record.Node.tag == "node":
								Matches.insert(0, dgm_database.PartialElementMatch(
									Record=SearchResult.Record,
								))
							self._AddMissingElement(GroupsByKey, GuiMissingElement(FilePath, Sheet.title, Row, Name), Matches)
						ConsecutiveIgnoredRows = 0
				if ConsecutiveIgnoredRows >= dgm_xlsx_common.STOP_AFTER_CONSECUTIVE_IGNORED_ROWS:
					break
				Row += 1
		return list(GroupsByKey.values()), IgnoredRows

	def _AddMissingElement(self, GroupsByKey: Dict[str, MissingElementGroup], Item: GuiMissingElement, Matches: List[dgm_database.PartialElementMatch]) -> None:
		Best = sorted(Matches, key=lambda Match: len(dgm_database.NormalizeText(Match.DisplayName)), reverse=True)
		if Best:
			GroupKey = "match:" + dgm_database.NormalizeText(Best[0].DisplayName)
			Title = Best[0].DisplayName
		else:
			Simple = Item.Name[:1].upper() if Item.Name else "#"
			GroupKey = "text:" + Simple.casefold()
			Title = f"Text: {Simple}"
		Group = GroupsByKey.setdefault(GroupKey, MissingElementGroup(GroupKey, Title, Title))
		NameKey = dgm_database.NormalizeText(Item.Name)
		Summary = Group.ItemsByName.get(NameKey)
		if Summary is None:
			Summary = MissingElementSummary(Item.Name)
			Group.ItemsByName[NameKey] = Summary
		Summary.Occurrences.append(Item)
