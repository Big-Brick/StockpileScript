from __future__ import annotations

import decimal
import tkinter as tk
import tkinter.messagebox
import tkinter.ttk as ttk
import xml.etree.ElementTree as XmlTree
from typing import Dict, List, Optional

import dgm_database
from dgm_gui_common import GuiAddElementResult, WINDOW_TITLE


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
		self._Kind = tk.StringVar(value="regex" if Node.tag == "regex" else "node")
		self._Optional = tk.BooleanVar(value=dgm_database.IsOptionalNode(Node))

		Row = 0
		ttk.Label(self, text="Node type").grid(row=Row, column=0, sticky="w", padx=(10, 6), pady=4)
		TypeFrame = ttk.Frame(self)
		TypeFrame.grid(row=Row, column=1, sticky="w", padx=(0, 10), pady=4)
		ttk.Radiobutton(TypeFrame, text="Regular node", variable=self._Kind, value="node", command=self._UpdateNodeState).grid(row=0, column=0, sticky="w", padx=(0, 10))
		ttk.Radiobutton(TypeFrame, text="Regex node", variable=self._Kind, value="regex", command=self._UpdateNodeState).grid(row=0, column=1, sticky="w")
		Row += 1

		ttk.Label(self, text="Optional").grid(row=Row, column=0, sticky="w", padx=(10, 6), pady=4)
		self._OptionalCheck = ttk.Checkbutton(self, text="Search children even when this node text is absent", variable=self._Optional)
		self._OptionalCheck.grid(row=Row, column=1, sticky="w", padx=(0, 10), pady=4)
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

	def _BuildFields(self, Node: XmlTree.Element) -> list[tuple[str, str, str]]:
		DgmNode = Node.find("dgm")
		SourceNode = DgmNode if DgmNode is not None else Node
		Fields = [
			("text", "Display text", Node.get("text", Node.get("name", ""))),
			("name", "Name", Node.get("name", "")),
			("pattern", "Regex pattern", Node.get("pattern", "")),
		]
		for MetalKey, MetalName in dgm_database.METALS:
			Fields.append((MetalKey, f"{MetalName} g", SourceNode.get(f"{MetalKey}_g", "0")))
		return Fields

	def _UpdateNodeState(self) -> None:
		IsRegex = self._Kind.get() == "regex"
		PatternEntry = self._Entries.get("pattern")
		if PatternEntry is not None:
			PatternEntry.configure(state="normal" if IsRegex else "disabled")
		self._OptionalCheck.configure(state="disabled" if IsRegex else "normal")
		if IsRegex:
			self._Optional.set(False)

	def _Save(self) -> None:
		try:
			for MetalKey, _ in dgm_database.METALS:
				dgm_database.ReadDecimal(self._Entries[MetalKey].get())
		except decimal.InvalidOperation as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, f"Invalid decimal value: {Error}", parent=self)
			return

		self.Result = {Key: Entry.get() for Key, Entry in self._Entries.items()}
		self.Result["kind"] = self._Kind.get()
		self.Result["optional"] = "true" if self._Optional.get() else "false"
		self.destroy()

	def _Cancel(self) -> None:
		self.Result = None
		self.destroy()


class MoveCatalogNodeDialog(tk.Toplevel):
	def __init__(self, Parent: tk.Tk, CurrentParentPath: List[str]) -> None:
		super().__init__(Parent)
		self.Result: Optional[List[str]] = None

		self.title("Move catalog node")
		self.transient(Parent)
		self.grab_set()
		self.resizable(False, False)

		self.columnconfigure(0, weight=1)
		self.rowconfigure(1, weight=1)

		ttk.Label(self, text="New parent path").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 4))

		self.PathList = tk.Listbox(self, height=8, width=48, exportselection=False)
		self.PathList.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 6))

		for Part in CurrentParentPath:
			self.PathList.insert(tk.END, Part)

		EditFrame = ttk.Frame(self)
		EditFrame.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 6))
		EditFrame.columnconfigure(0, weight=1)

		self.PartEntry = ttk.Entry(EditFrame)
		self.PartEntry.grid(row=0, column=0, sticky="ew", padx=(0, 6))

		ttk.Button(EditFrame, text="Add", command=self._AddPart).grid(row=0, column=1, padx=(0, 6))
		ttk.Button(EditFrame, text="Replace", command=self._ReplacePart).grid(row=0, column=2)

		ButtonFrame = ttk.Frame(self)
		ButtonFrame.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 10))

		ttk.Button(ButtonFrame, text="Remove", command=self._RemovePart).grid(row=0, column=0, padx=(0, 6))
		ttk.Button(ButtonFrame, text="Up", command=self._MovePartUp).grid(row=0, column=1, padx=(0, 6))
		ttk.Button(ButtonFrame, text="Down", command=self._MovePartDown).grid(row=0, column=2, padx=(0, 6))

		ttk.Button(ButtonFrame, text="Cancel", command=self._Cancel).grid(row=0, column=3, sticky="e", padx=(20, 6))
		ttk.Button(ButtonFrame, text="Move", command=self._Move).grid(row=0, column=4, sticky="e")

		ttk.Label(
			self,
			text="Each row is one parent node. Empty list means catalog root."
		).grid(row=4, column=0, sticky="w", padx=10, pady=(0, 10))

		self.PathList.bind("<<ListboxSelect>>", self._OnSelect)

		self.bind("<Escape>", lambda _Event: self._Cancel())
		self.bind("<Return>", lambda _Event: self._Move())
		self.protocol("WM_DELETE_WINDOW", self._Cancel)

		self.PartEntry.focus_set()
		self.wait_window(self)

	def _GetSelectedIndex(self) -> Optional[int]:
		Selection = self.PathList.curselection()
		if not Selection:
			return None
		return int(Selection[0])

	def _OnSelect(self, _Event: tk.Event) -> None:
		Index = self._GetSelectedIndex()
		if Index is None:
			return

		self.PartEntry.delete(0, tk.END)
		self.PartEntry.insert(0, self.PathList.get(Index))

	def _AddPart(self) -> None:
		Part = self.PartEntry.get()

		# ВАЖЛИВО: не робимо .strip(), бо trailing space може бути значущим.
		if Part == "":
			return

		self.PathList.insert(tk.END, Part)
		self.PartEntry.delete(0, tk.END)

	def _ReplacePart(self) -> None:
		Index = self._GetSelectedIndex()
		if Index is None:
			return

		Part = self.PartEntry.get()

		# ВАЖЛИВО: не робимо .strip().
		if Part == "":
			return

		self.PathList.delete(Index)
		self.PathList.insert(Index, Part)
		self.PathList.selection_set(Index)

	def _RemovePart(self) -> None:
		Index = self._GetSelectedIndex()
		if Index is None:
			return

		self.PathList.delete(Index)

	def _MovePartUp(self) -> None:
		Index = self._GetSelectedIndex()
		if Index is None or Index <= 0:
			return

		Part = self.PathList.get(Index)
		self.PathList.delete(Index)
		self.PathList.insert(Index - 1, Part)
		self.PathList.selection_set(Index - 1)

	def _MovePartDown(self) -> None:
		Index = self._GetSelectedIndex()
		if Index is None or Index >= self.PathList.size() - 1:
			return

		Part = self.PathList.get(Index)
		self.PathList.delete(Index)
		self.PathList.insert(Index + 1, Part)
		self.PathList.selection_set(Index + 1)

	def _Move(self) -> None:
		self.Result = []

		for Index in range(self.PathList.size()):
			self.Result.append(self.PathList.get(Index))

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
	def __init__(
		self,
		Parent: tk.Toplevel,
		Db: dgm_database.DgmDatabase,
		Name: str,
		InitialPathParts: Optional[List[str]] = None,
		Title: str = "Add missing element",
		InitialMode: str = "auto",
	) -> None:
		super().__init__(Parent)
		self.Result: Optional[GuiAddElementResult] = None
		self.Db = Db
		self.Name = Name
		self.InitialPathParts = InitialPathParts
		StructuredResult = Db.FindStructuredElement(dgm_database.NormalizeText(Name), Name)
		self.ExactCandidate = StructuredResult.ExactMatch if StructuredResult.ExactMatch is not None and StructuredResult.ExactMatch.Node.tag == "node" else None
		self.Candidates = []
		if self.ExactCandidate is not None:
			self.Candidates.append(self.ExactCandidate)
		self.Candidates.extend(Candidate for Candidate in (StructuredResult.PartialMatches or []) if Candidate.Node.tag == "node" and Candidate is not self.ExactCandidate)
		self.title(Title)
		self.transient(Parent)
		self.grab_set()
		self.geometry("620x650")
		self.columnconfigure(0, weight=1)
		self.rowconfigure(3, weight=1)
		tk.Label(self, text=f"Element: {Name}").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 4))

		ModeFrame = ttk.LabelFrame(self, text="Action")
		ModeFrame.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))
		DefaultMode = "existing" if self.ExactCandidate is not None else "new"
		if InitialMode in ("existing", "new"):
			DefaultMode = InitialMode
		self.DialogMode = tk.StringVar(value=DefaultMode)
		ttk.Radiobutton(ModeFrame, text="Add DGM values to existing database node", variable=self.DialogMode, value="existing", command=self._UpdateModeState).grid(row=0, column=0, sticky="w", padx=6, pady=4)
		ttk.Radiobutton(ModeFrame, text="Add new database element", variable=self.DialogMode, value="new", command=self._UpdateModeState).grid(row=1, column=0, sticky="w", padx=6, pady=(0, 4))

		self.ExistingFrame = ttk.LabelFrame(self, text="Database candidates")
		self.ExistingFrame.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 8))
		self.CandidateList = tk.Listbox(self.ExistingFrame, height=5, exportselection=False)
		self.CandidateList.grid(row=0, column=0, sticky="ew", padx=6, pady=6)
		self.ExistingFrame.columnconfigure(0, weight=1)
		for Candidate in self.Candidates:
			Kind = "full match" if Candidate is self.ExactCandidate else "partial match"
			Marker = "has DGM" if Candidate.HasDgm else "no DGM"
			self.CandidateList.insert(tk.END, f"{Candidate.DisplayName} ({Kind}, {Marker})")
		if self.Candidates:
			self.CandidateList.selection_set(0)
		else:
			self.CandidateList.insert(tk.END, "No database candidates found.")

		self.NewFrame = ttk.Frame(self)
		self.NewFrame.grid(row=3, column=0, sticky="nsew", padx=10)
		self.NewFrame.columnconfigure(0, weight=1)
		self.NewFrame.rowconfigure(1, weight=1)
		ttk.Label(self.NewFrame, text="Structured node chain (one node per line)").grid(row=0, column=0, sticky="w", pady=(0, 4))
		self.PathList = tk.Listbox(self.NewFrame, height=8, exportselection=False)
		self.PathList.grid(row=1, column=0, sticky="nsew")
		for Part in (InitialPathParts if InitialPathParts is not None else self._DefaultSplit(Name)):
			self.PathList.insert(tk.END, Part)
		self.PathList.bind("<<ListboxSelect>>", self._OnPathPartSelect)
		Edit = ttk.Frame(self.NewFrame)
		Edit.grid(row=2, column=0, sticky="ew", pady=4)
		Edit.columnconfigure(0, weight=1)
		self.PartEntry = ttk.Entry(Edit)
		self.PartEntry.grid(row=0, column=0, sticky="ew", padx=(0, 6))
		ttk.Button(Edit, text="Split by /", command=self._SplitEntry).grid(row=0, column=1, padx=(0, 6))
		ttk.Button(Edit, text="Add", command=self._AddPart).grid(row=0, column=2, padx=(0, 6))
		ttk.Button(Edit, text="Replace", command=self._ReplacePart).grid(row=0, column=3)
		Controls = ttk.Frame(self.NewFrame)
		Controls.grid(row=3, column=0, sticky="w")
		ttk.Button(Controls, text="Remove", command=self._RemovePart).grid(row=0, column=0, padx=(0, 6))
		ttk.Button(Controls, text="Use existing candidate as parent", command=self._UseCandidateAsParent).grid(row=0, column=1)
		AddModeFrame = ttk.LabelFrame(self.NewFrame, text="New element type")
		AddModeFrame.grid(row=4, column=0, sticky="ew", pady=(8, 0))
		AddModeFrame.columnconfigure(1, weight=1)
		self.AddMode = tk.StringVar(value="node")
		ttk.Radiobutton(AddModeFrame, text="Exact structured node", variable=self.AddMode, value="node").grid(row=0, column=0, sticky="w", padx=(6, 12), pady=4)
		ttk.Radiobutton(AddModeFrame, text="Regex leaf node", variable=self.AddMode, value="regex").grid(row=0, column=1, sticky="w", pady=4)
		ttk.Label(AddModeFrame, text="Regex pattern").grid(row=1, column=0, sticky="w", padx=6, pady=(0, 4))
		self.RegexPatternEntry = ttk.Entry(AddModeFrame)
		self.RegexPatternEntry.grid(row=1, column=1, sticky="ew", padx=(0, 6), pady=(0, 4))
		ttk.Label(AddModeFrame, text="Display text").grid(row=2, column=0, sticky="w", padx=6, pady=(0, 6))
		self.RegexDisplayEntry = ttk.Entry(AddModeFrame)
		self.RegexDisplayEntry.grid(row=2, column=1, sticky="ew", padx=(0, 6), pady=(0, 6))

		Values = ttk.LabelFrame(self, text="DGM values, g")
		Values.grid(row=4, column=0, sticky="ew", padx=10, pady=8)
		self.ValueEntries: Dict[str, ttk.Entry] = {}
		for Index, (MetalKey, MetalName) in enumerate(dgm_database.METALS):
			ttk.Label(Values, text=MetalName).grid(row=0, column=Index, sticky="w")
			Entry = ttk.Entry(Values, width=12)
			Entry.insert(0, "0")
			Entry.grid(row=1, column=Index, padx=(0, 6))
			self.ValueEntries[MetalKey] = Entry
		Buttons = ttk.Frame(self)
		Buttons.grid(row=5, column=0, sticky="e", padx=10, pady=(0, 10))
		ttk.Button(Buttons, text="Cancel", command=self._Cancel).grid(row=0, column=0, padx=(0, 6))
		ttk.Button(Buttons, text="Add", command=self._Save).grid(row=0, column=1)
		self._UpdateModeState()
		self.wait_window(self)

	def GetElementName(self) -> str:
		if self.Result is not None and self.Result.Mode == "regex":
			return self.Result.DisplayText or self.Result.Pattern or self.Name
		if self.Result is not None and self.Result.PathParts:
			return "".join(self.Result.PathParts)
		return self.Name

	def _DefaultSplit(self, Name: str) -> List[str]:
		return [Part for Part in Name.split("/") if Part] or [Name]

	def _UpdateModeState(self) -> None:
		self.ExistingFrame.grid()
		if self.DialogMode.get() == "existing":
			self.NewFrame.grid_remove()
		else:
			self.NewFrame.grid()

	def _GetCandidate(self) -> Optional[dgm_database.PartialElementMatch]:
		Selection = self.CandidateList.curselection()
		if not Selection or not self.Candidates:
			return None
		Index = int(Selection[0])
		return self.Candidates[Index] if Index < len(self.Candidates) else None

	def _OnPathPartSelect(self, _Event: tk.Event) -> None:
		Selection = self.PathList.curselection()
		if not Selection:
			return
		self.PartEntry.delete(0, tk.END)
		self.PartEntry.insert(0, self.PathList.get(int(Selection[0])))

	def _UseCandidateAsParent(self) -> None:
		Candidate = self._GetCandidate()
		if Candidate is None:
			return
		self.PathList.delete(0, tk.END)
		PathParts = self.Db.GetNodePathParts(Candidate.Node)
		LeafRemainder = self._GetNameRemainderAfterCandidatePath(PathParts)
		for Part in PathParts + [LeafRemainder]:
			self.PathList.insert(tk.END, Part)

	def _GetNameRemainderAfterCandidatePath(self, PathParts: List[str]) -> str:
		CandidatePrefix = "".join(PathParts)
		if CandidatePrefix and self.Name.casefold().startswith(CandidatePrefix.casefold()):
			Remainder = self.Name[len(CandidatePrefix):]
			if Remainder:
				return Remainder
		return self.Name

	def _SplitEntry(self) -> None:
		Parts = [Part for Part in self.PartEntry.get().split("/") if Part]
		if not Parts:
			Parts = self._DefaultSplit(self.Name)
		Selection = self.PathList.curselection()
		InsertIndex = tk.END
		if Selection:
			InsertIndex = int(Selection[0])
			self.PathList.delete(InsertIndex)
		for Part in Parts:
			self.PathList.insert(InsertIndex, Part)
			if InsertIndex != tk.END:
				InsertIndex += 1
		self.PartEntry.delete(0, tk.END)

	def _AddPart(self) -> None:
		if self.PartEntry.get():
			self.PathList.insert(tk.END, self.PartEntry.get())
			self.PartEntry.delete(0, tk.END)

	def _ReplacePart(self) -> None:
		Selection = self.PathList.curselection()
		if Selection and self.PartEntry.get():
			Index = int(Selection[0])
			self.PathList.delete(Index)
			self.PathList.insert(Index, self.PartEntry.get())

	def _RemovePart(self) -> None:
		Selection = self.PathList.curselection()
		if Selection:
			self.PathList.delete(int(Selection[0]))

	def _Save(self) -> None:
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
		if self.DialogMode.get() == "existing":
			Candidate = self._GetCandidate()
			if Candidate is None or Candidate is not self.ExactCandidate:
				tkinter.messagebox.showerror(WINDOW_TITLE, "Select the full existing match or choose Add new.", parent=self)
				return
			self.Result = GuiAddElementResult("existing", Values, self.Db.GetNodePathParts(Candidate.Node))
		else:
			PathParts = [self.PathList.get(Index) for Index in range(self.PathList.size())]
			if not PathParts:
				tkinter.messagebox.showerror(WINDOW_TITLE, "Enter at least one node.", parent=self)
				return
			if self.AddMode.get() == "regex":
				Pattern = self.RegexPatternEntry.get().strip() or PathParts[-1]
				DisplayText = self.RegexDisplayEntry.get().strip()
				if not Pattern:
					tkinter.messagebox.showerror(WINDOW_TITLE, "Regex pattern cannot be empty.", parent=self)
					return
				self.Result = GuiAddElementResult("regex", Values, PathParts[:-1], Pattern, DisplayText or Pattern)
			else:
				self.Result = GuiAddElementResult("new", Values, PathParts)
		self.destroy()

	def _Cancel(self) -> None:
		self.Result = None
		self.destroy()
