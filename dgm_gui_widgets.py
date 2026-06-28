from __future__ import annotations

import decimal
from typing import Callable, Dict, List, Optional
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox
import tkinter.simpledialog
import xml.etree.ElementTree as XmlTree

import dgm_database
from dgm_gui_common import GuiAddElementResult, WINDOW_TITLE
from dgm_gui_dialogs import AddElementDialog, CatalogNodeEditDialog, MoveCatalogNodeDialog


class NewNodePathEditor(ttk.Frame):
	EMPTY_MARKER = "○"
	FILLED_MARKER = "●"

	def __init__(
		self,
		Parent: tk.Misc,
		ListRows: int = 8,
		ListWidth: int = 54,
		ShowSplitButton: bool = False,
		SplitFallbackParts: Optional[List[str]] = None,
	) -> None:
		super().__init__(Parent)
		self._SelectedPathIndex = 0
		self._UpdatingEntry = False
		self._SplitFallbackParts = SplitFallbackParts or []

		self.columnconfigure(0, weight=1)
		self.rowconfigure(0, weight=1)

		self.PathList = tk.Listbox(self, height=ListRows, width=ListWidth, exportselection=False, activestyle="dotbox")
		self.PathList.grid(row=0, column=0, sticky="nsew")
		self.PathList.bind("<<ListboxSelect>>", self._OnPathPartSelect)
		self.PathList.bind("<FocusOut>", lambda _Event: self.EnsureSelection())

		EditFrame = ttk.Frame(self)
		EditFrame.grid(row=1, column=0, sticky="ew", pady=(4, 0))
		EditFrame.columnconfigure(0, weight=1)
		self.PartEntry = ttk.Entry(EditFrame)
		self.PartEntry.grid(row=0, column=0, sticky="ew", padx=(0, 6))
		self.PartEntry.bind("<KeyRelease>", self._OnEntryChanged)
		self.PartEntry.bind("<FocusIn>", lambda _Event: self.EnsureSelection())

		ButtonColumn = 1
		if ShowSplitButton:
			ttk.Button(EditFrame, text="Split by /", command=self.SplitEntry).grid(row=0, column=ButtonColumn, padx=(0, 6))
			ButtonColumn += 1
		ttk.Button(EditFrame, text="Add empty row", command=self.AddEmptyRow).grid(row=0, column=ButtonColumn)

		Controls = ttk.Frame(self)
		Controls.grid(row=2, column=0, sticky="w", pady=(4, 0))
		ttk.Button(Controls, text="Remove", command=self.RemoveSelectedRow).grid(row=0, column=0, padx=(0, 6))
		ttk.Button(Controls, text="Up", command=self.MoveSelectedRowUp).grid(row=0, column=1, padx=(0, 6))
		ttk.Button(Controls, text="Down", command=self.MoveSelectedRowDown).grid(row=0, column=2, padx=(0, 6))

		self.SetPathParts([])

	def FocusEntry(self) -> None:
		self.PartEntry.focus_set()

	def SetSplitFallbackParts(self, Parts: List[str]) -> None:
		self._SplitFallbackParts = Parts

	def GetPathParts(self, IncludeEmpty: bool = False) -> List[str]:
		Parts = self._RawPathParts()
		if IncludeEmpty:
			return Parts
		return [Part for Part in Parts if Part != ""]

	def SetPathParts(self, Parts: List[str]) -> None:
		self.PathList.delete(0, tk.END)
		for Part in Parts or [""]:
			self.PathList.insert(tk.END, self._DisplayPart(Part))
		self._SelectPathIndex(min(self._SelectedPathIndex, self.PathList.size() - 1))

	def EnsureSelection(self) -> None:
		self._SelectPathIndex(self._SelectedPathIndex)

	def SplitEntry(self) -> None:
		Parts = [Part for Part in self.PartEntry.get().split("/") if Part]
		if not Parts:
			Parts = list(self._SplitFallbackParts)
		if not Parts:
			Parts = [""]

		self.EnsureSelection()
		AllParts = self._RawPathParts()
		AllParts[self._SelectedPathIndex:self._SelectedPathIndex + 1] = Parts
		self.SetPathParts(AllParts)
		self._SelectPathIndex(min(self._SelectedPathIndex + len(Parts) - 1, self.PathList.size() - 1))

	def AddEmptyRow(self) -> None:
		self.PathList.insert(tk.END, self._DisplayPart(""))
		self._SelectPathIndex(self.PathList.size() - 1)

	def RemoveSelectedRow(self) -> None:
		self.EnsureSelection()
		self.PathList.delete(self._SelectedPathIndex)
		self._SelectPathIndex(min(self._SelectedPathIndex, self.PathList.size() - 1))

	def MoveSelectedRowUp(self) -> None:
		self._MoveSelectedRow(-1)

	def MoveSelectedRowDown(self) -> None:
		self._MoveSelectedRow(1)

	def _RawPathParts(self) -> List[str]:
		return [self.PathList.get(Index)[2:] for Index in range(self.PathList.size())]

	def _DisplayPart(self, Part: str) -> str:
		return f"{self.FILLED_MARKER if Part else self.EMPTY_MARKER} {Part}"

	def _SelectPathIndex(self, Index: int) -> None:
		self._SetListSelection(Index)
		self._UpdatingEntry = True
		self.PartEntry.delete(0, tk.END)
		self.PartEntry.insert(0, self._RawPathParts()[self._SelectedPathIndex])
		self._UpdatingEntry = False

	def _SetListSelection(self, Index: int) -> None:
		if self.PathList.size() == 0:
			self.PathList.insert(tk.END, self._DisplayPart(""))
		self._SelectedPathIndex = max(0, min(Index, self.PathList.size() - 1))
		self.PathList.selection_clear(0, tk.END)
		self.PathList.selection_set(self._SelectedPathIndex)
		self.PathList.activate(self._SelectedPathIndex)
		self.PathList.see(self._SelectedPathIndex)

	def _OnPathPartSelect(self, _Event: tk.Event) -> None:
		Selection = self.PathList.curselection()
		if Selection:
			self._SelectPathIndex(int(Selection[0]))
		else:
			self.EnsureSelection()

	def _OnEntryChanged(self, _Event: tk.Event) -> None:
		if self._UpdatingEntry:
			return

		CurrentText = self.PartEntry.get()
		self._SetListSelection(self._SelectedPathIndex)
		Parts = self._RawPathParts()
		Parts[self._SelectedPathIndex] = CurrentText
		self.PathList.delete(self._SelectedPathIndex)
		self.PathList.insert(self._SelectedPathIndex, self._DisplayPart(CurrentText))
		self._SetListSelection(self._SelectedPathIndex)

	def _MoveSelectedRow(self, Delta: int) -> None:
		self.EnsureSelection()
		NewIndex = self._SelectedPathIndex + Delta
		if NewIndex < 0 or NewIndex >= self.PathList.size():
			return
		Parts = self._RawPathParts()
		Parts[self._SelectedPathIndex], Parts[NewIndex] = Parts[NewIndex], Parts[self._SelectedPathIndex]
		self._SelectedPathIndex = NewIndex
		self.SetPathParts(Parts)



class CatalogTreeEditorWidget(ttk.Frame):
	def __init__(self, Parent: tk.Misc, Database: dgm_database.DgmDatabase, RootNode: Optional[XmlTree.Element] = None, RootTitle: Optional[str] = None, OnDatabaseChanged: Optional[Callable[[], None]] = None) -> None:
		super().__init__(Parent)
		self.Database = Database
		self.RootNode = RootNode
		self.RootTitle = RootTitle
		self.OnDatabaseChanged = OnDatabaseChanged
		self.CatalogItems: Dict[str, XmlTree.Element] = {}
		self.CatalogItemPaths: Dict[str, List[str]] = {}
		self._BuildLayout()
		self.PopulateCatalogTree(PreserveState=False)

	def PopulateCatalogTree(self, PreserveState: bool = True) -> None:
		self._PopulateCatalogTree(PreserveState)

	def _NotifyChanged(self) -> None:
		if self.OnDatabaseChanged is not None:
			self.OnDatabaseChanged()

	def _BuildLayout(self) -> None:
		self.columnconfigure(0, weight=1)
		self.rowconfigure(1, weight=1)
		Parent = self

		HeaderFrame = ttk.Frame(Parent)
		HeaderFrame.grid(row=0, column=0, sticky="ew", pady=(0, 4))
		HeaderFrame.columnconfigure(1, weight=1)

		ttk.Label(HeaderFrame, text="Component database", style="Heading.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 10))
		self.SearchEntry = ttk.Entry(HeaderFrame)
		self.SearchEntry.grid(row=0, column=1, sticky="ew", padx=(0, 6))
		ttk.Button(HeaderFrame, text="Search", command=self._SearchCatalog).grid(row=0, column=2)
		self.SearchStatusLabel = ttk.Label(Parent, text="Enter a component name and click Search to test exact and partial matches.")
		self.SearchStatusLabel.grid(row=3, column=0, sticky="ew", pady=(4, 0))
		self.SearchEntry.bind("<Return>", lambda _Event: self._SearchCatalog())

		Columns = ("kind", "optional", "name", "gold", "silver", "platinum", "mpg")
		self.CatalogTree = ttk.Treeview(Parent, columns=Columns, show="tree headings", selectmode="browse")
		self.CatalogTree.heading("#0", text="Component / structure")
		self.CatalogTree.heading("kind", text="Type")
		self.CatalogTree.heading("optional", text="Optional")
		self.CatalogTree.heading("name", text="Name")
		self.CatalogTree.heading("gold", text="Gold g")
		self.CatalogTree.heading("silver", text="Silver g")
		self.CatalogTree.heading("platinum", text="Platinum g")
		self.CatalogTree.heading("mpg", text="MPG g")

		self.CatalogTree.column("#0", width=260, minwidth=180, stretch=True)
		self.CatalogTree.column("kind", width=85, minwidth=65, anchor="center", stretch=False)
		self.CatalogTree.column("optional", width=75, minwidth=65, anchor="center", stretch=False)
		self.CatalogTree.column("name", width=180, minwidth=120, stretch=True)
		for Column in ("gold", "silver", "platinum", "mpg"):
			self.CatalogTree.column(Column, width=90, minwidth=70, anchor="e", stretch=False)

		VerticalScrollbar = ttk.Scrollbar(Parent, orient=tk.VERTICAL, command=self.CatalogTree.yview)
		HorizontalScrollbar = ttk.Scrollbar(Parent, orient=tk.HORIZONTAL, command=self.CatalogTree.xview)
		self.CatalogTree.configure(yscrollcommand=VerticalScrollbar.set, xscrollcommand=HorizontalScrollbar.set)

		self.CatalogTree.grid(row=1, column=0, sticky="nsew")
		VerticalScrollbar.grid(row=1, column=1, sticky="ns")
		HorizontalScrollbar.grid(row=2, column=0, sticky="ew")

		self.CatalogContextMenu = tk.Menu(self, tearoff=False)
		self.CatalogContextMenu.add_command(label="Add node...", command=self._AddCatalogNode)
		self.CatalogContextMenu.add_command(label="Modify database node...", command=self._ModifySelectedCatalogNode)
		self.CatalogContextMenu.add_command(label="Move node to another parent...", command=self._MoveSelectedCatalogNode)
		self.CatalogContextMenu.add_command(label="Move all children to another parent...", command=self._MoveSelectedCatalogNodeChildren)
		self.CatalogContextMenu.add_command(label="Move trailing characters to children...", command=self._MoveTrailingCharactersToChildren)
		self.CatalogContextMenu.add_separator()
		self.CatalogContextMenu.add_command(label="Remove node", command=self._RemoveSelectedCatalogNode)
		self.CatalogTree.bind("<Button-3>", self._ShowCatalogContextMenu)
		self.CatalogTree.bind("<Control-Button-1>", self._ShowCatalogContextMenu)

	def _PopulateCatalogTree(self, PreserveState: bool = True) -> None:
		ExpandedNodeIds, SelectedNodeId = self._GetCatalogTreeState() if PreserveState else (set(), None)
		self.CatalogItems.clear()
		self.CatalogItemPaths.clear()
		for Item in self.CatalogTree.get_children():
			self.CatalogTree.delete(Item)

		RootNode = self.RootNode if self.RootNode is not None else self.Database.CatalogNode
		RootText = self.RootTitle if self.RootTitle is not None else "catalog"
		CatalogRoot = self.CatalogTree.insert("", "end", text=RootText, values=("root", "", "", "", "", "", ""), open=True)
		self.CatalogItemPaths[CatalogRoot] = []
		if self.RootNode is not None:
			self.CatalogItems[CatalogRoot] = self.RootNode
		for Child in list(RootNode):
			self._InsertCatalogNode(CatalogRoot, Child, [])

		self._SortCatalogTreeChildren()
		if PreserveState:
			self._RestoreCatalogTreeState(ExpandedNodeIds, SelectedNodeId)

	def _GetCatalogTreeState(self) -> tuple[set[int], Optional[int]]:
		ExpandedNodeIds: set[int] = set()
		for ItemId, Node in self.CatalogItems.items():
			if self.CatalogTree.item(ItemId, "open"):
				ExpandedNodeIds.add(id(Node))

		SelectedNodeId: Optional[int] = None
		SelectedNode = self._GetSelectedCatalogNode()
		if SelectedNode is not None:
			SelectedNodeId = id(SelectedNode)
		return ExpandedNodeIds, SelectedNodeId

	def _RestoreCatalogTreeState(self, ExpandedNodeIds: set[int], SelectedNodeId: Optional[int]) -> None:
		SelectedItem = ""
		for ItemId, Node in self.CatalogItems.items():
			NodeId = id(Node)
			if NodeId in ExpandedNodeIds:
				self.CatalogTree.item(ItemId, open=True)
			if SelectedNodeId is not None and NodeId == SelectedNodeId:
				SelectedItem = ItemId
		for RootItem in self.CatalogTree.get_children(""):
			self.CatalogTree.item(RootItem, open=True)
		if SelectedItem:
			self.CatalogTree.selection_set(SelectedItem)
			self.CatalogTree.focus(SelectedItem)
			self.CatalogTree.see(SelectedItem)

	def _SortCatalogTreeChildren(self, ParentItemId: str = "") -> None:
		Children = list(self.CatalogTree.get_children(ParentItemId))

		Children.sort(key=self._CatalogTreeSortKey)

		for Index, ChildItemId in enumerate(Children):
			self.CatalogTree.move(ChildItemId, ParentItemId, Index)
			self._SortCatalogTreeChildren(ChildItemId)

	def _CatalogTreeSortKey(self, ItemId: str) -> tuple[str]:
		Text = self.CatalogTree.item(ItemId, "text")
		return (Text.casefold(),)

	def _InsertCatalogNode(self, ParentItem: str, Node: XmlTree.Element, ParentPathParts: List[str]) -> None:
		if Node.tag not in dgm_database.CATALOG_ELEMENT_TAGS:
			return

		PathText = Node.get("text", Node.get("name", Node.tag))
		PathParts = ParentPathParts + [PathText]
		Text = '|' + PathText + '|'
		Item = self.CatalogTree.insert(
			ParentItem,
			"end",
			text=Text,
			values=(
				dgm_database.REGEX_LEAF_TAG if Node.tag == dgm_database.REGEX_LEAF_TAG else "node",
				"yes" if dgm_database.IsOptionalNode(Node) else "",
				Node.get("name", ""),
				*self._ReadDgmColumns(Node),
			),
			open=False,
		)
		self.CatalogItems[Item] = Node
		self.CatalogItemPaths[Item] = PathParts
		for Child in list(Node):
			self._InsertCatalogNode(Item, Child, PathParts)

	def _ReadDgmColumns(self, Node: XmlTree.Element) -> tuple[str, str, str, str]:
		DgmNode = Node.find("dgm")
		SourceNode = DgmNode if DgmNode is not None else Node
		if not self._HasDgmValues(SourceNode):
			return ("", "", "", "")
		return tuple(SourceNode.get(f"{MetalKey}_g", "0") for MetalKey, _ in dgm_database.METALS)  # type: ignore[return-value]

	def _HasDgmValues(self, Node: XmlTree.Element) -> bool:
		return any(Node.get(f"{MetalKey}_g") is not None for MetalKey, _ in dgm_database.METALS)


	def _ShowCatalogContextMenu(self, Event: tk.Event) -> str:
		ItemId = self.CatalogTree.identify_row(Event.y)
		if ItemId == "" or ItemId not in self.CatalogItems:
			return "break"

		self.CatalogTree.selection_set(ItemId)
		self.CatalogTree.focus(ItemId)
		self.CatalogContextMenu.tk_popup(Event.x_root, Event.y_root)
		return "break"

	def _GetSelectedCatalogItemId(self) -> str:
		ItemId = self.CatalogTree.focus()
		if not ItemId:
			Selection = self.CatalogTree.selection()
			ItemId = Selection[0] if Selection else ""
		return ItemId

	def _GetSelectedCatalogNode(self) -> Optional[XmlTree.Element]:
		return self.CatalogItems.get(self._GetSelectedCatalogItemId())

	def _GetSelectedCatalogPathParts(self) -> List[str]:
		return self.CatalogItemPaths.get(self._GetSelectedCatalogItemId(), [])

	def _GetSelectedCatalogSearchResult(self) -> dgm_database.ElementSearchResult:
		ItemId = self._GetSelectedCatalogItemId()
		if not ItemId or ItemId not in self.CatalogItems:
			return dgm_database.ElementSearchResult(None)

		NodeChain: List[XmlTree.Element] = []
		CurrentItem = ItemId
		while CurrentItem in self.CatalogItems:
			NodeChain.append(self.CatalogItems[CurrentItem])
			CurrentItem = self.CatalogTree.parent(CurrentItem)
		NodeChain.reverse()

		RootRecord = dgm_database.ElementRecord(
			self.Database,
			self.Database.CatalogNode,
			"catalog",
			dgm_database.DgmValues(decimal.Decimal("0"), decimal.Decimal("0"), decimal.Decimal("0"), decimal.Decimal("0")),
			False,
			None,
			"",
			"",
			False,
			True,
		)
		ParentRecord = RootRecord
		for Node in NodeChain:
			PathText = Node.get("text", Node.get("name", Node.tag))
			DisplayText = dgm_database.FormatOptionalPathText(PathText) if dgm_database.IsOptionalNode(Node) else PathText
			ParentRecord = self.Database.MakeRecord(
				Node,
				DisplayText,
				ParentRecord,
				PathText,
				PathText,
				ParentRecord.MatchedByRegex or Node.tag == dgm_database.REGEX_LEAF_TAG,
			)
		return dgm_database.ElementSearchResult(ParentRecord)

	def _ApplyAddElementResult(self, Name: str, Result: GuiAddElementResult) -> None:
		if Result.Mode == "existing":
			self.Database.AddDgmToExistingPath(Name, Result.Values, Result.PathParts)
		elif Result.Mode == "regex":
			self.Database.AddRegexElement(Name, Result.Values, Result.PathParts, Result.RegexText)
		elif Result.Mode == "node_no_dgm":
			self.Database.AddRegularNodePath(Result.PathParts)
		else:
			self.Database.AddElement(Name, Result.Values, Result.PathParts)

	def _AddCatalogNode(self) -> None:
		ParentNode = self._GetSelectedCatalogNode()
		if ParentNode is None:
			return

		DefaultName = "New node"
		Dialog = AddElementDialog(
			self,
			DefaultName,
			self._GetSelectedCatalogSearchResult(),
			InitialMode="new",
		)
		if Dialog.Result is None:
			return

		try:
			Name = Dialog.GetElementName()
			self._ApplyAddElementResult(Name, Dialog.Result)
			self.Database.Save()
			self._NotifyChanged()
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, str(Error), parent=self)
			return

		self._PopulateCatalogTree(PreserveState=True)

	def _ModifySelectedCatalogNode(self) -> None:
		Node = self._GetSelectedCatalogNode()
		if Node is None:
			return

		Dialog = CatalogNodeEditDialog(self, Node)
		if Dialog.Result is None:
			return

		try:
			self._ApplyCatalogNodeChanges(Node, Dialog.Result)
			self.Database.Save()
			self._NotifyChanged()
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, str(Error), parent=self)
			return

		self._PopulateCatalogTree(PreserveState=True)

	def _ApplyCatalogNodeChanges(self, Node: XmlTree.Element, Values: Dict[str, str]) -> None:
		Text = Values["text"]
		Name = Values["name"]
		NodeKind = Values["kind"]
		if not Text:
			raise RuntimeError("Display text is required")
		if NodeKind not in dgm_database.CATALOG_ELEMENT_TAGS:
			raise RuntimeError("Unsupported catalog node type")

		Node.tag = NodeKind
		Node.set("text", Text)
		if Name:
			Node.set("name", Name)
		elif "name" in Node.attrib:
			del Node.attrib["name"]

		dgm_database.SetOptionalNode(Node, NodeKind == "node" and Values.get("optional") == "true")

		if NodeKind == dgm_database.REGEX_LEAF_TAG and not Text:
			raise RuntimeError("Regex leaf text is required")
		Node.attrib.pop("pattern", None)

		if Values.get("has_dgm") != "true":
			for DgmNode in list(Node.findall("dgm")):
				Node.remove(DgmNode)
			for MetalKey, _ in dgm_database.METALS:
				Node.attrib.pop(f"{MetalKey}_g", None)
			return

		DgmNode = self.Database.EnsureChild(Node, "dgm")
		for MetalKey, _ in dgm_database.METALS:
			Value = dgm_database.ReadDecimal(Values[MetalKey])
			DgmNode.set(f"{MetalKey}_g", dgm_database.DecimalToText(Value))

	def _MoveSelectedCatalogNode(self) -> None:
		Node = self._GetSelectedCatalogNode()
		if Node is None:
			return

		CurrentParent = self.Database.FindCatalogParentOfNode(Node) or self.Database.CatalogNode
		Dialog = MoveCatalogNodeDialog(self, self.Database, CurrentParent)
		if Dialog.Result is None:
			return

		try:
			ExistingParent, NewPathParts = Dialog.Result
			self.Database.MoveCatalogNodeToParent(Node, ExistingParent, NewPathParts)
			self.Database.Save()
			self._NotifyChanged()
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, str(Error), parent=self)
			return

		self._PopulateCatalogTree(PreserveState=True)

	def _MoveSelectedCatalogNodeChildren(self) -> None:
		Node = self._GetSelectedCatalogNode()
		if Node is None:
			return

		CurrentParent = self.Database.FindCatalogParentOfNode(Node) or self.Database.CatalogNode
		Dialog = MoveCatalogNodeDialog(self, self.Database, CurrentParent)
		if Dialog.Result is None:
			return

		try:
			ExistingParent, NewPathParts = Dialog.Result
			self.Database.MoveCatalogChildrenToParent(Node, ExistingParent, NewPathParts)
			self.Database.Save()
			self._NotifyChanged()
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, str(Error), parent=self)
			return

		self._PopulateCatalogTree(PreserveState=True)

	def _MoveTrailingCharactersToChildren(self) -> None:
		Node = self._GetSelectedCatalogNode()
		if Node is None:
			return

		Count = tkinter.simpledialog.askinteger(
			WINDOW_TITLE,
			"How many trailing characters should be removed from the selected node and prepended to each child?",
			parent=self,
			minvalue=1,
		)
		if Count is None:
			return

		try:
			self.Database.MoveTrailingTextToChildren(Node, Count)
			self.Database.Save()
			self._NotifyChanged()
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, str(Error), parent=self)
			return

		self._PopulateCatalogTree(PreserveState=True)

	def _RemoveSelectedCatalogNode(self) -> None:
		Node = self._GetSelectedCatalogNode()
		if Node is None:
			return

		Warnings: List[str] = []
		if any(Child.tag in dgm_database.CATALOG_ELEMENT_TAGS for Child in list(Node)):
			Warnings.append("it has child nodes")
		if self.Database.CatalogNodeHasNonZeroDgmValues(Node):
			Warnings.append("it has non-zero DGM values")

		if Warnings:
			Name = Node.get("text", Node.get("name", Node.tag))
			Message = f"Remove '{Name}' and its subtree? This node is not empty because " + " and ".join(Warnings) + "."
			if not tkinter.messagebox.askyesno(WINDOW_TITLE, Message, parent=self):
				return

		try:
			self.Database.RemoveCatalogNode(Node)
			self.Database.Save()
			self._NotifyChanged()
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, str(Error), parent=self)
			return

		self._PopulateCatalogTree(PreserveState=True)


	def _SearchCatalog(self) -> None:
		SearchText = self.SearchEntry.get().strip()
		if not SearchText:
			self.SearchStatusLabel.configure(text="Enter a component name to search.")
			return

		Result = self.Database.FindElement(SearchText)
		Messages: List[str] = []
		if Result.Record is not None:
			MatchType = "regex leaf" if Result.MatchedByRegex else "exact"
			if Result.Record.HasDgm:
				Marker = "non-zero DGM" if Result.Record.HasNonZeroDgm else "zero DGM"
				Messages.append(f"Found {MatchType} match with {Marker}: {Result.Record.DisplayName}")
			else:
				Messages.append(f"Found {MatchType} match without DGM: {Result.Record.DisplayName}")
		else:
			Messages.append("No exact match found.")

		PartialMatches = Result.PartialMatches
		if PartialMatches:
			PartialText = "; ".join(
				f"{Match.DisplayName}{' (has DGM)' if Match.HasDgm else ''}"
				for Match in PartialMatches
			)
			Messages.append(f"Partial matches: {PartialText}")
		else:
			Messages.append("No partial matches.")

		self.SearchStatusLabel.configure(text=" ".join(Messages))
