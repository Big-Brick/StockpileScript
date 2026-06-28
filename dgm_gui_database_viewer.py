from __future__ import annotations

import tkinter as tk
import tkinter.ttk as ttk

import dgm_database
from dgm_gui_common import WINDOW_TITLE
from dgm_gui_catalog_editor import CatalogEditorPanel
from dgm_gui_xlsx_processor import XlsxProcessingMixin


class DgmDatabaseViewer(tk.Toplevel, XlsxProcessingMixin):
	def __init__(self, Parent: tk.Tk) -> None:
		super().__init__(Parent)
		self.ParentApplication = Parent
		self.DatabasePath = Parent.DatabasePath  # type: ignore[attr-defined]
		self.Database = Parent.Database  # type: ignore[attr-defined]

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
		Parent.rowconfigure(0, weight=1)
		self.CatalogWidget = CatalogEditorPanel(Parent, self.Database, OnDatabaseChanged=self._PopulateIgnoredList)
		self.CatalogWidget.grid(row=0, column=0, sticky="nsew")

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
		self.CatalogWidget.PopulateCatalogTree(PreserveState=True)
		self._PopulateIgnoredList()

	def _PopulateIgnoredList(self) -> None:
		self.IgnoredList.delete(0, tk.END)
		IgnoredValues = sorted(
			(Node.get("value", "") for Node in self.Database.IgnoredNode.findall("text")),
			key=dgm_database.NormalizeText,
		)
		for Value in IgnoredValues:
			self.IgnoredList.insert(tk.END, Value)
		self.IgnoredCountLabel.configure(text=f"{len(IgnoredValues)} ignored items")
