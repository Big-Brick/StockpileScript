from __future__ import annotations

import decimal
from typing import Dict, List, Optional
import tkinter as tk
import tkinter.messagebox
import tkinter.ttk as ttk
import xml.etree.ElementTree as XmlTree

import dgm_database
from dgm_gui_common import GuiAddElementResult, WINDOW_TITLE
from dgm_gui_dialogs import AddElementDialog, CatalogNodeEditDialog, MoveCatalogNodeDialog
from dgm_gui_xlsx_processor import XlsxProcessingMixin


class DgmDatabaseViewer(tk.Toplevel, XlsxProcessingMixin):
	def __init__(self, Parent: tk.Tk) -> None:
		super().__init__(Parent)
		self.ParentApplication = Parent
		self.DatabasePath = Parent.DatabasePath  # type: ignore[attr-defined]
		self.Database = Parent.Database  # type: ignore[attr-defined]
		self.CatalogItems: Dict[str, XmlTree.Element] = {}
		self.CatalogItemPaths: Dict[str, List[str]] = {}

		self.title(f"{WINDOW_TITLE} - {self.DatabasePath.name}")
		self.geometry("1100x700")
		self.minsize(850, 500)

		self._ConfigureStyle()
		self._BuildLayout()
		self._PopulateDatabaseViews()

	def _ConfigureStyle(self) -> None:
		Style = ttk.Style(self)
		if "clam" in Style.theme_names():
			Style.theme_use("clam")
		Style.configure("Treeview", rowheight=24)
		Style.configure("Heading.TLabel", font=("TkDefaultFont", 10, "bold"))

	def _BuildLayout(self) -> None:
		self.columnconfigure(0, weight=1)
		self.rowconfigure(1, weight=1)

		Header = ttk.Frame(self, padding=(10, 8, 10, 4))
		Header.grid(row=0, column=0, sticky="ew")
		Header.columnconfigure(1, weight=1)

		ttk.Label(Header, text="Database:", style="Heading.TLabel").grid(row=0, column=0, sticky="w")
		ttk.Label(Header, text=str(self.DatabasePath)).grid(row=0, column=1, sticky="w", padx=(6, 0))

		Content = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
		Content.grid(row=1, column=0, sticky="nsew", padx=10, pady=(4, 10))

		CatalogPane = ttk.Frame(Content)
		IgnoredPane = ttk.Frame(Content)
		Content.add(CatalogPane, weight=4)
		Content.add(IgnoredPane, weight=1)

		self._BuildCatalogPane(CatalogPane)
		self._BuildIgnoredPane(IgnoredPane)

	def _BuildCatalogPane(self, Parent: ttk.Frame) -> None:
		Parent.columnconfigure(0, weight=1)
		Parent.rowconfigure(1, weight=1)

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
		self.CatalogContextMenu.add_separator()
		self.CatalogContextMenu.add_command(label="Remove node", command=self._RemoveSelectedCatalogNode)
		self.CatalogTree.bind("<Button-3>", self._ShowCatalogContextMenu)
		self.CatalogTree.bind("<Control-Button-1>", self._ShowCatalogContextMenu)

	def _BuildIgnoredPane(self, Parent: ttk.Frame) -> None:
		Parent.columnconfigure(0, weight=1)
		Parent.rowconfigure(1, weight=1)

		ttk.Label(Parent, text="Ignored items", style="Heading.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 4))

		Frame = ttk.Frame(Parent)
		Frame.grid(row=1, column=0, sticky="nsew")
		Frame.columnconfigure(0, weight=1)
		Frame.rowconfigure(0, weight=1)

		self.IgnoredList = tk.Listbox(Frame, activestyle="dotbox", exportselection=False)
		VerticalScrollbar = ttk.Scrollbar(Frame, orient=tk.VERTICAL, command=self.IgnoredList.yview)
		HorizontalScrollbar = ttk.Scrollbar(Frame, orient=tk.HORIZONTAL, command=self.IgnoredList.xview)
		self.IgnoredList.configure(yscrollcommand=VerticalScrollbar.set, xscrollcommand=HorizontalScrollbar.set)

		self.IgnoredList.grid(row=0, column=0, sticky="nsew")
		VerticalScrollbar.grid(row=0, column=1, sticky="ns")
		HorizontalScrollbar.grid(row=1, column=0, sticky="ew")

		self.IgnoredCountLabel = ttk.Label(Parent, text="0 ignored items")
		self.IgnoredCountLabel.grid(row=2, column=0, sticky="w", pady=(4, 0))

	def _PopulateDatabaseViews(self) -> None:
		self._PopulateCatalogTree()
		self._PopulateIgnoredList()

	def _PopulateCatalogTree(self) -> None:
		self.CatalogItems.clear()
		self.CatalogItemPaths.clear()
		for Item in self.CatalogTree.get_children():
			self.CatalogTree.delete(Item)

		CatalogRoot = self.CatalogTree.insert("", "end", text="catalog", values=("root", "", "", "", "", "", ""), open=True)
		self.CatalogItemPaths[CatalogRoot] = []
		for Child in list(self.Database.CatalogNode):
			self._InsertCatalogNode(CatalogRoot, Child, [])

		self._SortCatalogTreeChildren()

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
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, str(Error), parent=self)
			return

		self._PopulateCatalogTree()

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
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, str(Error), parent=self)
			return

		self._PopulateCatalogTree()

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

		Dialog = MoveCatalogNodeDialog(self, self._GetSelectedCatalogPathParts()[:-1])
		if Dialog.Result is None:
			return

		try:
			self.Database.MoveCatalogNode(Node, Dialog.Result)
			self.Database.Save()
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, str(Error), parent=self)
			return

		self._PopulateCatalogTree()

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
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, str(Error), parent=self)
			return

		self._PopulateCatalogTree()


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

	def _PopulateIgnoredList(self) -> None:
		self.IgnoredList.delete(0, tk.END)
		IgnoredValues = sorted(
			(Node.get("value", "") for Node in self.Database.IgnoredNode.findall("text")),
			key=dgm_database.NormalizeText,
		)
		for Value in IgnoredValues:
			self.IgnoredList.insert(tk.END, Value)
		self.IgnoredCountLabel.configure(text=f"{len(IgnoredValues)} ignored items")
