from __future__ import annotations

from pathlib import Path
from typing import Optional
import tkinter as tk
import tkinter.ttk as ttk

import dgm_database
from dgm_gui_database_viewer import DgmDatabaseViewer
from dgm_gui_xlsx_processor import XlsxProcessingMixin
from dgm_gui_xlsx_preprocessor import XlsxPreprocessingMixin
from dgm_gui_missing_elements import MissingElementsMixin
from dgm_gui_xlsx_conflicts import XlsxConflictsMixin
from dgm_gui_xlsx_postprocessor import XlsxPostprocessingMixin


class DgmMainWindow(tk.Tk, XlsxProcessingMixin, XlsxPreprocessingMixin, MissingElementsMixin, XlsxConflictsMixin, XlsxPostprocessingMixin):
	def __init__(self, DatabasePath: Path) -> None:
		super().__init__()
		self.DatabasePath = DatabasePath
		self.Database = dgm_database.OpenDatabase(DatabasePath)
		self.DatabaseEditor: Optional[DgmDatabaseViewer] = None
		self.ProcessSubfolders = tk.BooleanVar(value=False)
		self.AutoFillXlsxWithoutReview = tk.BooleanVar(value=False)

		self.title(f"DGM Inventory Tools - {DatabasePath.name}")
		self.geometry("360x610")
		self.minsize(320, 560)

		self._ConfigureStyle()
		self._BuildLayout()

	def _ConfigureStyle(self) -> None:
		Style = ttk.Style(self)
		if "clam" in Style.theme_names():
			Style.theme_use("clam")
		Style.configure("Heading.TLabel", font=("TkDefaultFont", 10, "bold"))

	def _BuildLayout(self) -> None:
		self.columnconfigure(0, weight=1)
		ttk.Label(self, text="DGM inventory tools", style="Heading.TLabel").grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 8))
		ttk.Button(self, text="Open database editor", command=self._OpenDatabaseEditor).grid(row=1, column=0, sticky="ew", padx=16, pady=4)
		ttk.Button(self, text="Preprocess .xlsx file", command=self._SelectAndPreprocessXlsxFile).grid(row=2, column=0, sticky="ew", padx=16, pady=4)
		ttk.Button(self, text="Preprocess folder", command=self._SelectAndPreprocessXlsxFolder).grid(row=3, column=0, sticky="ew", padx=16, pady=4)
		ttk.Checkbutton(self, text="Process selected folder subfolders", variable=self.ProcessSubfolders).grid(row=4, column=0, sticky="w", padx=16, pady=4)
		ttk.Checkbutton(self, text="Fill without review window", variable=self.AutoFillXlsxWithoutReview).grid(row=5, column=0, sticky="w", padx=16, pady=4)
		ttk.Button(self, text="Fill .xlsx file", command=self._SelectAndProcessXlsxFile).grid(row=6, column=0, sticky="ew", padx=16, pady=4)
		ttk.Button(self, text="Clean .xlsx DGM cells", command=self._SelectAndCleanXlsxFile).grid(row=7, column=0, sticky="ew", padx=16, pady=4)
		ttk.Button(self, text="Fill folder", command=self._SelectAndProcessXlsxFolder).grid(row=8, column=0, sticky="ew", padx=16, pady=4)
		ttk.Button(self, text="Review XLSX/database conflicts in file", command=self._SelectAndReviewXlsxConflictsFile).grid(row=9, column=0, sticky="ew", padx=16, pady=4)
		ttk.Button(self, text="Review XLSX/database conflicts in folder", command=self._SelectAndReviewXlsxConflictsFolder).grid(row=10, column=0, sticky="ew", padx=16, pady=4)
		ttk.Button(self, text="List missing elements in .xlsx file", command=self._SelectAndListMissingXlsxFile).grid(row=11, column=0, sticky="ew", padx=16, pady=4)
		ttk.Button(self, text="List missing elements in folder", command=self._SelectAndListMissingXlsxFolder).grid(row=12, column=0, sticky="ew", padx=16, pady=4)
		ttk.Button(self, text="Postprocess .xlsx file", command=self._SelectAndPostprocessXlsxFile).grid(row=13, column=0, sticky="ew", padx=16, pady=4)
		ttk.Button(self, text="Postprocess folder", command=self._SelectAndPostprocessXlsxFolder).grid(row=14, column=0, sticky="ew", padx=16, pady=4)
		ttk.Button(self, text="Postprocess registry", command=self._SelectAndPostprocessRegistry).grid(row=15, column=0, sticky="ew", padx=16, pady=(4, 16))

	def _OpenDatabaseEditor(self) -> None:
		if self.DatabaseEditor is not None and self.DatabaseEditor.winfo_exists():
			self.DatabaseEditor.lift()
			self.DatabaseEditor.focus_force()
			return
		self.DatabaseEditor = DgmDatabaseViewer(self)
		self.DatabaseEditor.protocol("WM_DELETE_WINDOW", self._CloseDatabaseEditor)

	def _CloseDatabaseEditor(self) -> None:
		if self.DatabaseEditor is not None and self.DatabaseEditor.winfo_exists():
			self.DatabaseEditor.destroy()
		self.DatabaseEditor = None

	def _PopulateDatabaseViews(self) -> None:
		if self.DatabaseEditor is not None and self.DatabaseEditor.winfo_exists():
			self.DatabaseEditor._PopulateDatabaseViews()

	def _PopulateIgnoredList(self) -> None:
		if self.DatabaseEditor is not None and self.DatabaseEditor.winfo_exists():
			self.DatabaseEditor._PopulateIgnoredList()
