from __future__ import annotations

import decimal
import tkinter as tk
import tkinter.messagebox
import tkinter.ttk as ttk
import xml.etree.ElementTree as XmlTree
from typing import Dict, List, Optional

import dgm_database
from dgm_gui_common import GuiAddElementResult, WINDOW_TITLE
from dgm_gui_widgets import NewNodePathEditor


class CatalogNodeEditDialog(tk.Toplevel):
	def __init__(self, Parent: tk.Tk, Node: XmlTree.Element) -> None:
		super().__init__(Parent)
		self.Result: Optional[Dict[str, str]] = None
		self.title("Modify database node")
		self.transient(Parent)
		self.grab_set()
		self.resizable(False, False)
		self.columnconfigure(1, weight=1)

		self._Entries: Dict[str, tk.Entry] = {}
		self._Kind = tk.StringVar(value=dgm_database.REGEX_LEAF_TAG if Node.tag == dgm_database.REGEX_LEAF_TAG else "node")
		self._Optional = tk.BooleanVar(value=dgm_database.IsOptionalNode(Node))
		self._HasDgm = tk.BooleanVar(value=self._NodeHasDgmValues(Node))

		Row = 0
		ttk.Label(self, text="Node type").grid(row=Row, column=0, sticky="w", padx=(10, 6), pady=4)
		TypeFrame = ttk.Frame(self)
		TypeFrame.grid(row=Row, column=1, sticky="w", padx=(0, 10), pady=4)
		ttk.Radiobutton(TypeFrame, text="Regular node", variable=self._Kind, value="node", command=self._UpdateNodeState).grid(row=0, column=0, sticky="w", padx=(0, 10))
		ttk.Radiobutton(TypeFrame, text="Regex leaf node", variable=self._Kind, value=dgm_database.REGEX_LEAF_TAG, command=self._UpdateNodeState).grid(row=0, column=1, sticky="w")
		Row += 1

		ttk.Label(self, text="Optional").grid(row=Row, column=0, sticky="w", padx=(10, 6), pady=4)
		self._OptionalCheck = ttk.Checkbutton(self, text="Search children even when this node text is absent", variable=self._Optional)
		self._OptionalCheck.grid(row=Row, column=1, sticky="w", padx=(0, 10), pady=4)
		Row += 1

		ttk.Label(self, text="DGM values").grid(row=Row, column=0, sticky="w", padx=(10, 6), pady=4)
		self._HasDgmCheck = ttk.Checkbutton(
			self,
			text="Store DGM values on this node",
			variable=self._HasDgm,
			command=self._UpdateDgmState,
		)
		self._HasDgmCheck.grid(row=Row, column=1, sticky="w", padx=(0, 10), pady=4)
		Row += 1

		for Key, Label, Value in self._BuildFields(Node):
			ttk.Label(self, text=Label).grid(row=Row, column=0, sticky="w", padx=(10, 6), pady=4)
			Entry = ttk.Entry(self, width=42)
			Entry.insert(0, Value)
			Entry.grid(row=Row, column=1, sticky="ew", padx=(0, 10), pady=4)
			self._Entries[Key] = Entry
			Row += 1
		self._UpdateNodeState()

		ButtonFrame = ttk.Frame(self)
		ButtonFrame.grid(row=Row, column=0, columnspan=2, sticky="e", padx=10, pady=(8, 10))
		ttk.Button(ButtonFrame, text="Cancel", command=self._Cancel).grid(row=0, column=0, padx=(0, 6))
		ttk.Button(ButtonFrame, text="Save", command=self._Save).grid(row=0, column=1)

		self.bind("<Escape>", lambda _Event: self._Cancel())
		self.bind("<Return>", lambda _Event: self._Save())
		self.protocol("WM_DELETE_WINDOW", self._Cancel)
		self._Entries["text"].focus_set()
		self.wait_window(self)

	def _NodeHasDgmValues(self, Node: XmlTree.Element) -> bool:
		if Node.find("dgm") is not None:
			return True
		return any(Node.get(f"{MetalKey}_g") is not None for MetalKey, _ in dgm_database.METALS)

	def _BuildFields(self, Node: XmlTree.Element) -> list[tuple[str, str, str]]:
		DgmNode = Node.find("dgm")
		SourceNode = DgmNode if DgmNode is not None else Node
		TextLabel = "Regex pattern" if Node.tag == dgm_database.REGEX_LEAF_TAG else "Display text"
		Fields = [
			("text", TextLabel, Node.get("text", Node.get("name", ""))),
			("name", "Name", Node.get("name", "")),
		]
		for MetalKey, MetalName in dgm_database.METALS:
			Fields.append((MetalKey, f"{MetalName} g", SourceNode.get(f"{MetalKey}_g", "0")))
		return Fields

	def _UpdateNodeState(self) -> None:
		IsRegexLeaf = self._Kind.get() == dgm_database.REGEX_LEAF_TAG
		self._OptionalCheck.configure(state="disabled" if IsRegexLeaf else "normal")
		if IsRegexLeaf:
			self._Optional.set(False)
		self._UpdateDgmState()

	def _UpdateDgmState(self) -> None:
		State = "normal" if self._HasDgm.get() else "disabled"
		for MetalKey, _ in dgm_database.METALS:
			Entry = self._Entries.get(MetalKey)
			if Entry is not None:
				Entry.configure(state=State)

	def _Save(self) -> None:
		if self._HasDgm.get():
			try:
				for MetalKey, _ in dgm_database.METALS:
					dgm_database.ReadDecimal(self._Entries[MetalKey].get())
			except decimal.InvalidOperation as Error:
				tkinter.messagebox.showerror(WINDOW_TITLE, f"Invalid decimal value: {Error}", parent=self)
				return

		self.Result = {Key: Entry.get() for Key, Entry in self._Entries.items()}
		self.Result["kind"] = self._Kind.get()
		self.Result["optional"] = "true" if self._Optional.get() else "false"
		self.Result["has_dgm"] = "true" if self._HasDgm.get() else "false"
		self.destroy()

	def _Cancel(self) -> None:
		self.Result = None
		self.destroy()


class MoveCatalogNodeDialog(tk.Toplevel):
	STOP_VALUE = "<use selected existing node>"

	def __init__(self, Parent: tk.Tk, Database: dgm_database.DgmDatabase, CurrentParentPath: List[str]) -> None:
		super().__init__(Parent)
		self.Database = Database
		self.Result: Optional[List[str]] = None
		self._ExistingCombos: List[ttk.Combobox] = []
		self._ExistingSelections: List[tk.StringVar] = []

		self.title("Move catalog node")
		self.transient(Parent)
		self.grab_set()
		self.resizable(False, False)
		self.columnconfigure(0, weight=1)
		self.rowconfigure(1, weight=1)

		ExistingFrame = ttk.LabelFrame(self, text="Existing parent node")
		ExistingFrame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))
		ExistingFrame.columnconfigure(0, weight=1)
		self.ExistingComboFrame = ttk.Frame(ExistingFrame)
		self.ExistingComboFrame.grid(row=0, column=0, sticky="ew", padx=6, pady=6)
		self.ExistingComboFrame.columnconfigure(0, weight=1)
		ttk.Label(ExistingFrame, text="Pick children one level at a time; choose the stop value to use the previous node.").grid(row=1, column=0, sticky="w", padx=6, pady=(0, 6))

		NewFrame = ttk.LabelFrame(self, text="New nodes to create under selected existing node")
		NewFrame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 6))
		NewFrame.columnconfigure(0, weight=1)
		NewFrame.rowconfigure(0, weight=1)
		self.NewPathEditor = NewNodePathEditor(NewFrame, ListRows=8, ListWidth=54)
		self.NewPathEditor.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)

		ButtonFrame = ttk.Frame(self)
		ButtonFrame.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 10))
		ttk.Button(ButtonFrame, text="Cancel", command=self._Cancel).grid(row=0, column=0, sticky="e", padx=(20, 6))
		ttk.Button(ButtonFrame, text="Move", command=self._Move).grid(row=0, column=1, sticky="e")

		self._SetExistingPath(CurrentParentPath)
		self.bind("<Escape>", lambda _Event: self._Cancel())
		self.bind("<Return>", lambda _Event: self._Move())
		self.protocol("WM_DELETE_WINDOW", self._Cancel)
		self.NewPathEditor.FocusEntry()
		self.wait_window(self)

	def _NodeText(self, Node: XmlTree.Element) -> str:
		return Node.get("text", Node.get("name", Node.tag))

	def _ExistingPathParts(self) -> List[str]:
		Parts: List[str] = []
		for Var in self._ExistingSelections:
			Value = Var.get()
			if Value == self.STOP_VALUE or not Value:
				break
			Parts.append(Value)
			NextNode = self.Database.FindParent(Parts)
			if NextNode is None:
				break
		return Parts

	def _ChildNames(self, ParentPath: List[str]) -> List[str]:
		Parent = self.Database.FindParent(ParentPath)
		if Parent is None:
			return []
		return [self._NodeText(Child) for Child in list(Parent) if Child.tag == "node"]

	def _SetExistingPath(self, Parts: List[str]) -> None:
		for Widget in self._ExistingCombos:
			Widget.destroy()
		self._ExistingCombos.clear()
		self._ExistingSelections.clear()
		ParentPath: List[str] = []
		for Part in Parts:
			if Part not in self._ChildNames(ParentPath):
				break
			self._AddExistingCombo(ParentPath, Part)
			ParentPath.append(Part)
		self._AddExistingCombo(ParentPath, self.STOP_VALUE)

	def _AddExistingCombo(self, ParentPath: List[str], Selected: str) -> None:
		Var = tk.StringVar(value=Selected)
		Values = self._ChildNames(ParentPath) + [self.STOP_VALUE]
		Combo = ttk.Combobox(self.ExistingComboFrame, textvariable=Var, values=Values, state="readonly")
		Combo.grid(row=len(self._ExistingCombos), column=0, sticky="ew", pady=(0, 4))
		Combo.bind("<<ComboboxSelected>>", self._OnExistingSelectionChanged)
		self._ExistingCombos.append(Combo)
		self._ExistingSelections.append(Var)

	def _OnExistingSelectionChanged(self, _Event: tk.Event) -> None:
		self._SetExistingPath(self._ExistingPathParts())

	def _Move(self) -> None:
		self.Result = self._ExistingPathParts() + self.NewPathEditor.GetPathParts()
		self.destroy()

	def _Cancel(self) -> None:
		self.Result = None
		self.destroy()





class RenameElementDialog(tk.Toplevel):
	def __init__(self, Parent: tk.Toplevel, CurrentName: str) -> None:
		super().__init__(Parent)
		self.Result: Optional[str] = None
		self.title("Rename XLSX element")
		self.transient(Parent)
		self.grab_set()
		self.resizable(False, False)
		ttk.Label(self, text="New element text").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 4))
		self.Entry = ttk.Entry(self, width=60)
		self.Entry.insert(0, CurrentName)
		self.Entry.grid(row=1, column=0, sticky="ew", padx=10)
		Buttons = ttk.Frame(self)
		Buttons.grid(row=2, column=0, sticky="e", padx=10, pady=10)
		ttk.Button(Buttons, text="Cancel", command=self._Cancel).grid(row=0, column=0, padx=(0, 6))
		ttk.Button(Buttons, text="Save", command=self._Save).grid(row=0, column=1)
		self.wait_window(self)

	def _Save(self) -> None:
		Value = self.Entry.get().strip()
		if Value:
			self.Result = Value
			self.destroy()

	def _Cancel(self) -> None:
		self.Result = None
		self.destroy()


class AddElementDialog(tk.Toplevel):

	WINDOW_WIDTH = 620
	TITLE_SECTION_HEIGHT = 34
	ACTION_SECTION_HEIGHT = 82
	CANDIDATE_SECTION_HEIGHT = 132
	NEW_SECTION_HEIGHT = 318
	VALUES_SECTION_HEIGHT = 82
	BUTTON_SECTION_HEIGHT = 48
	WINDOW_VERTICAL_PADDING = 32
	CANDIDATE_LIST_ROWS = 5
	PATH_LIST_ROWS = 8
	EXISTING_MODE_HEIGHT = TITLE_SECTION_HEIGHT + ACTION_SECTION_HEIGHT + CANDIDATE_SECTION_HEIGHT + VALUES_SECTION_HEIGHT + BUTTON_SECTION_HEIGHT + WINDOW_VERTICAL_PADDING
	NEW_MODE_HEIGHT = EXISTING_MODE_HEIGHT + NEW_SECTION_HEIGHT

	def __init__(
		self,
		Parent: tk.Toplevel,
		Name: str,
		StructuredResult: dgm_database.ElementSearchResult,
		InitialMode: str = "auto",
		InitialValues: Optional[dgm_database.DgmValues] = None,
	) -> None:
		super().__init__(Parent)
		self.Result: Optional[GuiAddElementResult] = None
		self.Name = Name
		self.AllCandidateRows = self._BuildCandidateRows(StructuredResult)
		self.CandidateRows: List[Optional[dgm_database.PartialElementMatch]] = []

		DefaultMode = "existing" if self._ExistingRows() else "new"
		if InitialMode in ("existing", "new"):
			DefaultMode = InitialMode

		self.title("Add missing element")
		self.transient(Parent)
		self.grab_set()
		self.geometry(self._ModeGeometry(DefaultMode))
		self.columnconfigure(0, weight=1)
		self.rowconfigure(3, weight=1)
		tk.Label(self, text=f"Element: {Name}").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 4))

		ModeFrame = ttk.LabelFrame(self, text="Action", height=self.ACTION_SECTION_HEIGHT)
		ModeFrame.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))
		ModeFrame.grid_propagate(False)
		self.DialogMode = tk.StringVar(value=DefaultMode)
		ttk.Radiobutton(ModeFrame, text="Add DGM values to existing database node", variable=self.DialogMode, value="existing", command=self._UpdateModeState).grid(row=0, column=0, sticky="w", padx=6, pady=4)
		ttk.Radiobutton(ModeFrame, text="Add new database element", variable=self.DialogMode, value="new", command=self._UpdateModeState).grid(row=1, column=0, sticky="w", padx=6, pady=(0, 4))

		self.ExistingFrame = ttk.LabelFrame(self, text="Database candidates", height=self.CANDIDATE_SECTION_HEIGHT)
		self.ExistingFrame.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 8))
		self.ExistingFrame.grid_propagate(False)
		self.CandidateList = tk.Listbox(self.ExistingFrame, height=self.CANDIDATE_LIST_ROWS, exportselection=False)
		self.CandidateList.grid(row=0, column=0, sticky="ew", padx=6, pady=6)
		self.CandidateList.bind("<<ListboxSelect>>", self._OnCandidateSelect)
		self.ExistingFrame.columnconfigure(0, weight=1)

		self.NewFrame = ttk.Frame(self, height=self.NEW_SECTION_HEIGHT)
		self.NewFrame.grid(row=3, column=0, sticky="nsew", padx=10)
		self.NewFrame.grid_propagate(False)
		self.NewFrame.columnconfigure(0, weight=1)
		self.NewFrame.rowconfigure(1, weight=1)
		ttk.Label(self.NewFrame, text="New structured node chain (one new node per line)").grid(row=0, column=0, sticky="w", pady=(0, 4))
		self.NewPathEditor = NewNodePathEditor(
			self.NewFrame,
			ListRows=self.PATH_LIST_ROWS,
			ShowSplitButton=True,
			SplitFallbackParts=self._DefaultSplit(self.Name),
		)
		self.NewPathEditor.grid(row=1, column=0, sticky="nsew")
		AddModeFrame = ttk.LabelFrame(self.NewFrame, text="New element type")
		AddModeFrame.grid(row=2, column=0, sticky="ew", pady=(8, 0))
		AddModeFrame.columnconfigure(2, weight=1)
		self.AddMode = tk.StringVar(value="node")
		ttk.Radiobutton(AddModeFrame, text="Exact structured node with DGM", variable=self.AddMode, value="node", command=self._UpdateDgmValueState).grid(row=0, column=0, sticky="w", padx=(6, 12), pady=4)
		ttk.Radiobutton(AddModeFrame, text="Regex leaf node", variable=self.AddMode, value="regex", command=self._UpdateDgmValueState).grid(row=0, column=1, sticky="w", padx=(0, 12), pady=4)
		ttk.Radiobutton(AddModeFrame, text="Regular node without DGM", variable=self.AddMode, value="node_no_dgm", command=self._UpdateDgmValueState).grid(row=0, column=2, sticky="w", pady=4)

		Values = ttk.LabelFrame(self, text="DGM values, g", height=self.VALUES_SECTION_HEIGHT)
		Values.grid(row=4, column=0, sticky="ew", padx=10, pady=8)
		Values.grid_propagate(False)
		self.ValueEntries: Dict[str, ttk.Entry] = {}
		self.ValueLabels: Dict[str, ttk.Label] = {}
		for Index, (MetalKey, MetalName) in enumerate(dgm_database.METALS):
			Label = ttk.Label(Values, text=MetalName)
			Label.grid(row=0, column=Index, sticky="w")
			self.ValueLabels[MetalKey] = Label
			Entry = ttk.Entry(Values, width=12)
			InitialValue = InitialValues.GetMetalValue(MetalKey) if InitialValues is not None else dgm_database.ReadDecimal("0")
			Entry.insert(0, dgm_database.DecimalToText(InitialValue))
			Entry.grid(row=1, column=Index, padx=(0, 6))
			self.ValueEntries[MetalKey] = Entry
		Buttons = ttk.Frame(self, height=self.BUTTON_SECTION_HEIGHT)
		Buttons.grid(row=5, column=0, sticky="e", padx=10, pady=(0, 10))
		ttk.Button(Buttons, text="Cancel", command=self._Cancel).grid(row=0, column=0, padx=(0, 6))
		ttk.Button(Buttons, text="Add", command=self._Save).grid(row=0, column=1)
		self._UpdateModeState()
		self._UpdateDgmValueState()
		self.wait_window(self)

	def _UpdateDgmValueState(self) -> None:
		DisableValues = self.DialogMode.get() == "new" and self.AddMode.get() == "node_no_dgm"
		State = "disabled" if DisableValues else "normal"
		for Entry in self.ValueEntries.values():
			Entry.configure(state=State)
		for Label in self.ValueLabels.values():
			Label.configure(state=State)

	def _BuildCandidateRows(self, StructuredResult: dgm_database.ElementSearchResult) -> List[dgm_database.PartialElementMatch]:
		Rows: List[dgm_database.PartialElementMatch] = []
		SeenNodes = set()

		def AddRecord(Record: dgm_database.ElementRecord, Remainder: str = "") -> None:
			if Record.Node.tag != "node" or id(Record.Node) in SeenNodes:
				return
			SeenNodes.add(id(Record.Node))
			Rows.append(dgm_database.PartialElementMatch(Record=Record, Remainder=Remainder))

		if StructuredResult.Record is not None:
			AddRecord(StructuredResult.Record)
			if (
				StructuredResult.Record.Node.tag == dgm_database.REGEX_LEAF_TAG
				and StructuredResult.Record.ConsumedText == ""
				and StructuredResult.Record.Parent is not None
			):
				AddRecord(StructuredResult.Record.Parent)
		for Match in StructuredResult.PartialMatches:
			if Match.Record.Node.tag != "node" or id(Match.Record.Node) in SeenNodes:
				continue
			SeenNodes.add(id(Match.Record.Node))
			Rows.append(Match)
		return Rows

	def _ExistingRows(self) -> List[dgm_database.PartialElementMatch]:
		return [Match for Match in self.AllCandidateRows if not Match.Record.HasDgm]

	def GetElementName(self) -> str:
		if self.Result is not None and self.Result.Mode == "regex":
			return self.Result.RegexText or self.Name
		if self.Result is not None and self.Result.PathParts:
			return "".join(self.Result.PathParts)
		return self.Name

	def _DefaultSplit(self, Name: str) -> List[str]:
		return [Part for Part in Name.split("/") if Part] or [Name]

	def _ModeGeometry(self, Mode: str) -> str:
		Height = self.EXISTING_MODE_HEIGHT if Mode == "existing" else self.NEW_MODE_HEIGHT
		return f"{self.WINDOW_WIDTH}x{Height}"

	def _UpdateModeState(self) -> None:
		Mode = self.DialogMode.get()
		self.geometry(self._ModeGeometry(Mode))
		self.ExistingFrame.configure(text="Parent database node" if Mode == "new" else "Database node to update")
		self._PopulateCandidateList(Mode)
		self.ExistingFrame.grid()
		if Mode == "existing":
			self.NewFrame.grid_remove()
		else:
			self.NewFrame.grid()
			self._ApplySelectedParentToPathEditor()
		self._UpdateDgmValueState()

	def _PopulateCandidateList(self, Mode: str) -> None:
		self.CandidateList.delete(0, tk.END)
		if Mode == "new":
			self.CandidateRows = [None] + self.AllCandidateRows
		else:
			self.CandidateRows = list(self._ExistingRows())

		if not self.CandidateRows:
			self.CandidateList.insert(tk.END, "No database candidates found.")
			return

		for Candidate in self.CandidateRows:
			self.CandidateList.insert(tk.END, self._CandidateDisplayName(Candidate))

		if Mode == "new":
			SelectedIndex = max(range(len(self.CandidateRows)), key=lambda Index: len(self._CandidateConsumedText(self.CandidateRows[Index])))
		else:
			SelectedIndex = 0
		self.CandidateList.selection_set(SelectedIndex)
		self.CandidateList.activate(SelectedIndex)

	def _CandidateDisplayName(self, Candidate: Optional[dgm_database.PartialElementMatch]) -> str:
		if Candidate is None:
			return "<catalog root>"
		Marker = "has DGM" if Candidate.Record.HasDgm else "no DGM"
		return f"{Candidate.DisplayName} ({Marker})"

	def _CandidatePathParts(self, Candidate: Optional[dgm_database.PartialElementMatch]) -> List[str]:
		return [] if Candidate is None else Candidate.Record.PathParts

	def _CandidateConsumedText(self, Candidate: Optional[dgm_database.PartialElementMatch]) -> str:
		if Candidate is None:
			return ""
		return "".join(Record.ConsumedText for Record in Candidate.Record.IterPath())

	def _GetCandidate(self) -> Optional[dgm_database.PartialElementMatch]:
		Selection = self.CandidateList.curselection()
		if not Selection or not self.CandidateRows:
			return None
		Index = int(Selection[0])
		return self.CandidateRows[Index] if Index < len(self.CandidateRows) else None

	def _OnCandidateSelect(self, _Event: tk.Event) -> None:
		if self.DialogMode.get() == "new":
			self._ApplySelectedParentToPathEditor()

	def _ApplySelectedParentToPathEditor(self) -> None:
		Candidate = self._GetCandidate()
		Parts = self._NewPathPartsForCandidate(Candidate)
		self.NewPathEditor.SetPathParts(Parts)

	def _NewPathPartsForCandidate(self, Candidate: Optional[dgm_database.PartialElementMatch]) -> List[str]:
		if Candidate is None:
			return self._DefaultSplit(self.Name)
		if Candidate.Remainder:
			return self._DefaultSplit(Candidate.Remainder)
		ConsumedText = self._CandidateConsumedText(Candidate)
		Remainder = self._RemoveConsumedPrefix(self.Name, ConsumedText)
		return self._DefaultSplit(Remainder) if Remainder else self._DefaultSplit(self.Name)

	def _RemoveConsumedPrefix(self, Name: str, ConsumedText: str) -> str:
		if ConsumedText and Name.casefold().startswith(ConsumedText.casefold()):
			return Name[len(ConsumedText):]
		return Name

	def _Save(self) -> None:
		if self.DialogMode.get() == "new" and self.AddMode.get() == "node_no_dgm":
			Values = dgm_database.DgmValues(
				GoldG=dgm_database.ReadDecimal("0"),
				SilverG=dgm_database.ReadDecimal("0"),
				PlatinumG=dgm_database.ReadDecimal("0"),
				MpgG=dgm_database.ReadDecimal("0"),
			)
		else:
			try:
				Values = dgm_database.DgmValues(
					GoldG=dgm_database.ReadDecimal(self.ValueEntries["gold"].get()),
					SilverG=dgm_database.ReadDecimal(self.ValueEntries["silver"].get()),
					PlatinumG=dgm_database.ReadDecimal(self.ValueEntries["platinum"].get()),
					MpgG=dgm_database.ReadDecimal(self.ValueEntries["mpg"].get()),
				)
			except decimal.InvalidOperation as Error:
				tkinter.messagebox.showerror(WINDOW_TITLE, f"Invalid DGM value: {Error}", parent=self)
				return
		Candidate = self._GetCandidate()
		if self.DialogMode.get() == "existing":
			if Candidate is None or Candidate.Record.HasDgm:
				tkinter.messagebox.showerror(WINDOW_TITLE, "Select an existing database node without DGM values or choose Add new.", parent=self)
				return
			self.Result = GuiAddElementResult("existing", Values, Candidate.Record.PathParts)
		else:
			NewPathParts = self.NewPathEditor.GetPathParts()
			if not NewPathParts:
				tkinter.messagebox.showerror(WINDOW_TITLE, "Enter at least one new node.", parent=self)
				return
			ParentPathParts = self._CandidatePathParts(Candidate)
			if self.AddMode.get() == "regex":
				RegexText = NewPathParts[-1].strip()
				if not RegexText:
					tkinter.messagebox.showerror(WINDOW_TITLE, "Regex leaf text cannot be empty.", parent=self)
					return
				self.Result = GuiAddElementResult("regex", Values, ParentPathParts + NewPathParts[:-1], RegexText)
			elif self.AddMode.get() == "node_no_dgm":
				self.Result = GuiAddElementResult("node_no_dgm", Values, ParentPathParts + NewPathParts)
			else:
				self.Result = GuiAddElementResult("new", Values, ParentPathParts + NewPathParts)
		self.destroy()

	def _Cancel(self) -> None:
		self.Result = None
		self.destroy()
