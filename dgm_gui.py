#!/usr/bin/env python3
"""
GUI editor for DGM XML databases.

Scope:
- visualize and edit the structured component catalog in a tree view;
- move catalog nodes between parents;
- visualize ignored item texts in a list view.

Usage:
    python dgm_gui.py [database.xml]

If the database argument is omitted, a file selection dialog is shown.
"""

from __future__ import annotations

import decimal
import sys
import tkinter as tk
import tkinter.filedialog
import tkinter.messagebox
import tkinter.ttk as ttk
import xml.etree.ElementTree as XmlTree
from pathlib import Path
from typing import Dict, Optional

import dgm_database


WINDOW_TITLE = "DGM Database Editor"


class DgmDatabaseViewer(tk.Tk):
	def __init__(self, DatabasePath: Path) -> None:
		super().__init__()
		self.DatabasePath = DatabasePath
		self.Database = dgm_database.OpenDatabase(DatabasePath)
		self.CatalogItems: Dict[str, XmlTree.Element] = {}

		self.title(f"{WINDOW_TITLE} - {DatabasePath.name}")
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
		ttk.Label(Header, text="Catalog editor").grid(row=0, column=2, sticky="e")

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

		ttk.Label(Parent, text="Component database", style="Heading.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 4))

		Columns = ("kind", "name", "gold", "silver", "platinum", "mpg", "pattern")
		self.CatalogTree = ttk.Treeview(Parent, columns=Columns, show="tree headings", selectmode="browse")
		self.CatalogTree.heading("#0", text="Component / structure")
		self.CatalogTree.heading("kind", text="Type")
		self.CatalogTree.heading("name", text="Name")
		self.CatalogTree.heading("gold", text="Gold g")
		self.CatalogTree.heading("silver", text="Silver g")
		self.CatalogTree.heading("platinum", text="Platinum g")
		self.CatalogTree.heading("mpg", text="MPG g")
		self.CatalogTree.heading("pattern", text="Regex pattern")

		self.CatalogTree.column("#0", width=260, minwidth=180, stretch=True)
		self.CatalogTree.column("kind", width=85, minwidth=65, anchor="center", stretch=False)
		self.CatalogTree.column("name", width=180, minwidth=120, stretch=True)
		for Column in ("gold", "silver", "platinum", "mpg"):
			self.CatalogTree.column(Column, width=90, minwidth=70, anchor="e", stretch=False)
		self.CatalogTree.column("pattern", width=180, minwidth=120, stretch=True)

		VerticalScrollbar = ttk.Scrollbar(Parent, orient=tk.VERTICAL, command=self.CatalogTree.yview)
		HorizontalScrollbar = ttk.Scrollbar(Parent, orient=tk.HORIZONTAL, command=self.CatalogTree.xview)
		self.CatalogTree.configure(yscrollcommand=VerticalScrollbar.set, xscrollcommand=HorizontalScrollbar.set)

		self.CatalogTree.grid(row=1, column=0, sticky="nsew")
		VerticalScrollbar.grid(row=1, column=1, sticky="ns")
		HorizontalScrollbar.grid(row=2, column=0, sticky="ew")

		self.CatalogContextMenu = tk.Menu(self, tearoff=False)
		self.CatalogContextMenu.add_command(label="Modify database node...", command=self._ModifySelectedCatalogNode)
		self.CatalogContextMenu.add_command(label="Move node to another parent...", command=self._MoveSelectedCatalogNode)
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
		for Item in self.CatalogTree.get_children():
			self.CatalogTree.delete(Item)

		CatalogRoot = self.CatalogTree.insert("", "end", text="catalog", values=("root", "", "", "", "", "", ""), open=True)
		for Child in list(self.Database.CatalogNode):
			self._InsertCatalogNode(CatalogRoot, Child)

		LegacyElements = self.Database.LegacyElementsNode.findall("element")
		if LegacyElements:
			LegacyRoot = self.CatalogTree.insert("", "end", text="legacy elements", values=("legacy", "", "", "", "", "", ""), open=True)
			for ElementNode in LegacyElements:
				self._InsertLegacyElement(LegacyRoot, ElementNode)

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

	def _InsertCatalogNode(self, ParentItem: str, Node: XmlTree.Element) -> None:
		if Node.tag not in ("node", "regex"):
			return

		Text = '|' + Node.get("text", Node.get("name", Node.tag)) + '|'
		Item = self.CatalogTree.insert(
			ParentItem,
			"end",
			text=Text,
			values=(
				"regex" if Node.tag == "regex" else "node",
				Node.get("name", ""),
				*self._ReadDgmColumns(Node),
				Node.get("pattern", ""),
			),
			open=False,
		)
		self.CatalogItems[Item] = Node
		for Child in list(Node):
			self._InsertCatalogNode(Item, Child)

	def _InsertLegacyElement(self, ParentItem: str, Node: XmlTree.Element) -> None:
		Name = Node.get("name", "")
		self.CatalogTree.insert(
			ParentItem,
			"end",
			text=Name,
			values=("legacy", Name, *self._ReadDgmColumns(Node), ""),
		)

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

	def _GetSelectedCatalogNode(self) -> Optional[XmlTree.Element]:
		ItemId = self.CatalogTree.focus()
		if not ItemId:
			Selection = self.CatalogTree.selection()
			ItemId = Selection[0] if Selection else ""
		return self.CatalogItems.get(ItemId)

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
		Text = Values["text"].strip()
		Name = Values["name"].strip()
		if not Text:
			raise RuntimeError("Display text is required")

		Node.set("text", Text)
		if Name:
			Node.set("name", Name)
		elif "name" in Node.attrib:
			del Node.attrib["name"]

		if Node.tag == "regex":
			Pattern = Values["pattern"].strip()
			if not Pattern:
				raise RuntimeError("Regex pattern is required for regex nodes")
			Node.set("pattern", Pattern)

		DgmNode = self.Database.EnsureChild(Node, "dgm")
		for MetalKey, _ in dgm_database.METALS:
			Value = dgm_database.ReadDecimal(Values[MetalKey])
			DgmNode.set(f"{MetalKey}_g", dgm_database.DecimalToText(Value))

	def _MoveSelectedCatalogNode(self) -> None:
		Node = self._GetSelectedCatalogNode()
		if Node is None:
			return

		CurrentPath = self.Database.GetNodePathParts(Node)
		CurrentParentPath = "/".join(CurrentPath[:-1])
		Dialog = MoveCatalogNodeDialog(self, CurrentParentPath)
		if Dialog.Result is None:
			return

		try:
			NewParentParts = [Part.strip() for Part in Dialog.Result.split("/") if Part.strip()]
			self.Database.MoveCatalogNode(Node, NewParentParts)
			self.Database.Save()
		except Exception as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, str(Error), parent=self)
			return

		self._PopulateCatalogTree()

	def _PopulateIgnoredList(self) -> None:
		self.IgnoredList.delete(0, tk.END)
		IgnoredValues = sorted(
			(Node.get("value", "") for Node in self.Database.IgnoredNode.findall("text")),
			key=dgm_database.NormalizeText,
		)
		for Value in IgnoredValues:
			self.IgnoredList.insert(tk.END, Value)
		self.IgnoredCountLabel.configure(text=f"{len(IgnoredValues)} ignored items")


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
		Row = 0
		for Key, Label, Value in self._BuildFields(Node):
			ttk.Label(self, text=Label).grid(row=Row, column=0, sticky="w", padx=(10, 6), pady=4)
			Entry = ttk.Entry(self, width=42)
			Entry.insert(0, Value)
			Entry.grid(row=Row, column=1, sticky="ew", padx=(0, 10), pady=4)
			self._Entries[Key] = Entry
			Row += 1

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
		]
		if Node.tag == "regex":
			Fields.append(("pattern", "Regex pattern", Node.get("pattern", "")))
		for MetalKey, MetalName in dgm_database.METALS:
			Fields.append((MetalKey, f"{MetalName} g", SourceNode.get(f"{MetalKey}_g", "0")))
		return Fields

	def _Save(self) -> None:
		try:
			for MetalKey, _ in dgm_database.METALS:
				dgm_database.ReadDecimal(self._Entries[MetalKey].get())
		except decimal.InvalidOperation as Error:
			tkinter.messagebox.showerror(WINDOW_TITLE, f"Invalid decimal value: {Error}", parent=self)
			return

		self.Result = {Key: Entry.get() for Key, Entry in self._Entries.items()}
		self.destroy()

	def _Cancel(self) -> None:
		self.Result = None
		self.destroy()


class MoveCatalogNodeDialog(tk.Toplevel):
	def __init__(self, Parent: tk.Tk, CurrentParentPath: str) -> None:
		super().__init__(Parent)
		self.Result: Optional[str] = None
		self.title("Move catalog node")
		self.transient(Parent)
		self.grab_set()
		self.resizable(False, False)
		self.columnconfigure(1, weight=1)

		ttk.Label(self, text="New parent path").grid(row=0, column=0, sticky="w", padx=(10, 6), pady=(10, 4))
		self.PathEntry = ttk.Entry(self, width=48)
		self.PathEntry.insert(0, CurrentParentPath)
		self.PathEntry.grid(row=0, column=1, sticky="ew", padx=(0, 10), pady=(10, 4))
		ttk.Label(self, text="Use / between parent nodes. Leave empty for catalog root.").grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=(0, 8))

		ButtonFrame = ttk.Frame(self)
		ButtonFrame.grid(row=2, column=0, columnspan=2, sticky="e", padx=10, pady=(0, 10))
		ttk.Button(ButtonFrame, text="Cancel", command=self._Cancel).grid(row=0, column=0, padx=(0, 6))
		ttk.Button(ButtonFrame, text="Move", command=self._Move).grid(row=0, column=1)

		self.bind("<Escape>", lambda _Event: self._Cancel())
		self.bind("<Return>", lambda _Event: self._Move())
		self.protocol("WM_DELETE_WINDOW", self._Cancel)
		self.PathEntry.focus_set()
		self.wait_window(self)

	def _Move(self) -> None:
		self.Result = self.PathEntry.get().strip()
		self.destroy()

	def _Cancel(self) -> None:
		self.Result = None
		self.destroy()


def SelectDatabaseWithDialog() -> Optional[Path]:
	Root = tk.Tk()
	Root.withdraw()
	Root.update()
	try:
		SelectedFile = tkinter.filedialog.askopenfilename(
			title="Select DGM XML database",
			filetypes=(("XML database", "*.xml"), ("All files", "*.*")),
		)
	finally:
		Root.destroy()

	if not SelectedFile:
		return None
	return Path(SelectedFile).expanduser().resolve()


def ResolveDatabaseArgument(Argument: Optional[str]) -> Optional[Path]:
	if Argument is not None:
		return Path(Argument).expanduser().resolve()
	return SelectDatabaseWithDialog()


def Main() -> int:
	if len(sys.argv) not in (1, 2):
		print("Usage: python dgm_gui.py [database.xml]", file=sys.stderr)
		return 2

	DatabasePath = ResolveDatabaseArgument(sys.argv[1] if len(sys.argv) == 2 else None)
	if DatabasePath is None:
		print("No database selected.", file=sys.stderr)
		return 2
	if not DatabasePath.exists() or not DatabasePath.is_file():
		print(f"Database file does not exist: {DatabasePath}", file=sys.stderr)
		return 2

	try:
		Application = DgmDatabaseViewer(DatabasePath)
	except Exception as Error:
		try:
			tkinter.messagebox.showerror(WINDOW_TITLE, str(Error))
		except Exception:
			pass
		print(str(Error), file=sys.stderr)
		return 1

	Application.mainloop()
	return 0


if __name__ == "__main__":
	sys.exit(Main())
