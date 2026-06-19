#!/usr/bin/env python3
"""GUI editor and XLSX tooling for DGM XML databases."""

from __future__ import annotations

import sys
import tkinter as tk
import tkinter.filedialog
import tkinter.messagebox
from pathlib import Path
from typing import Optional

from dgm_gui_common import WINDOW_TITLE
from dgm_gui_main_window import DgmMainWindow


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
		Application = DgmMainWindow(DatabasePath)
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
