from __future__ import annotations

from typing import List, Optional
import tkinter as tk
import tkinter.ttk as ttk


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
