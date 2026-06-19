#!/usr/bin/env python3
"""
XML database support for DGM inventory processing.

The database stores:
- spreadsheet column settings;
- ignored non-element texts;
- structured radio-element records;
- optional regex records for equivalent sibling variants.
"""

from __future__ import annotations

import copy
import decimal
import re
import string
import xml.etree.ElementTree as XmlTree
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

try:
	import openpyxl.utils
except ImportError:
	openpyxl = None  # type: ignore[assignment]


METALS: Tuple[Tuple[str, str], ...] = (
	("gold", "Gold"),
	("silver", "Silver"),
	("platinum", "Platinum"),
	("mpg", "MPG/palladium"),
)

DEFAULT_COLUMNS = {
	"name": "B",
	"quantity": "C",
	"gold": "D",
	"silver": "E",
	"platinum": "F",
	"mpg": "G",
	"total_gold": "H",
	"total_silver": "I",
	"total_platinum": "J",
	"total_mpg": "K",
}

DATABASE_VERSION = "2"
DATABASE_CODE_VERSION = "2026-06-19.1"
COLUMN_LETTERS = set(string.ascii_uppercase)


@dataclass
class Columns:
	Name: str
	Quantity: str
	PerElement: Dict[str, str]
	Total: Dict[str, str]


@dataclass
class DgmValues:
	GoldG: decimal.Decimal
	SilverG: decimal.Decimal
	PlatinumG: decimal.Decimal
	MpgG: decimal.Decimal

	def GetMetalValue(self, MetalKey: str) -> decimal.Decimal:
		if MetalKey == "gold":
			return self.GoldG
		if MetalKey == "silver":
			return self.SilverG
		if MetalKey == "platinum":
			return self.PlatinumG
		if MetalKey == "mpg":
			return self.MpgG
		raise ValueError(f"Unknown metal key: {MetalKey}")

	def IsZero(self) -> bool:
		return all(self.GetMetalValue(MetalKey) == 0 for MetalKey, _ in METALS)

	def SetMetalValue(self, MetalKey: str, ValueG: decimal.Decimal) -> None:
		if MetalKey == "gold":
			self.GoldG = ValueG
			return
		if MetalKey == "silver":
			self.SilverG = ValueG
			return
		if MetalKey == "platinum":
			self.PlatinumG = ValueG
			return
		if MetalKey == "mpg":
			self.MpgG = ValueG
			return
		raise ValueError(f"Unknown metal key: {MetalKey}")


@dataclass
class PartialElementMatch:
	Node: XmlTree.Element
	DisplayName: str
	MatchedByRegex: bool = False
	HasDgm: bool = False
	Remainder: str = ""


@dataclass
class ElementSearchResult:
	Record: Optional["ElementRecord"]
	MatchedByRegex: bool = False
	PartialMatches: Optional[List[PartialElementMatch]] = None

	def __post_init__(self) -> None:
		if self.PartialMatches is None:
			self.PartialMatches = []


@dataclass
class SiblingInfo:
	Index: int
	Kind: str
	Text: str
	Pattern: str
	HasDgm: bool


@dataclass
class ExistingPathInfo:
	Index: int
	PathParts: List[str]
	DisplayPath: str
	DisplayName: str
	HasDgm: bool


class ElementRecord:
	def __init__(self, Database: "DgmDatabase", Node: XmlTree.Element, DisplayName: str, Values: DgmValues, HasDgm: bool = True) -> None:
		self.Database = Database
		self.Node = Node
		self.DisplayName = DisplayName
		self.Values = Values
		self.HasDgm = HasDgm

	@property
	def HasNonZeroDgm(self) -> bool:
		return self.HasDgm and not self.Values.IsZero()

	@property
	def IsUsableForInventoryFill(self) -> bool:
		return self.HasDgm

	def GetMetalValue(self, MetalKey: str) -> decimal.Decimal:
		return self.Values.GetMetalValue(MetalKey)

	def SetMetalValue(self, MetalKey: str, ValueG: decimal.Decimal) -> None:
		self.Values.SetMetalValue(MetalKey, ValueG)
		self.WriteValuesToXml()

	def SetValues(self, Values: DgmValues) -> None:
		self.Values = Values
		self.WriteValuesToXml()

	def WriteValuesToXml(self) -> None:
		DgmNode = self.Database.EnsureChild(self.Node, "dgm")
		DgmNode.set("gold_g", DecimalToText(self.Values.GoldG))
		DgmNode.set("silver_g", DecimalToText(self.Values.SilverG))
		DgmNode.set("platinum_g", DecimalToText(self.Values.PlatinumG))
		DgmNode.set("mpg_g", DecimalToText(self.Values.MpgG))


class DgmDatabase:
	def __init__(self, PathToDatabase: Path) -> None:
		self.Path = PathToDatabase
		self.Tree: XmlTree.ElementTree
		self.Root: XmlTree.Element
		self.SettingsNode: XmlTree.Element
		self.ColumnsNode: XmlTree.Element
		self.CatalogNode: XmlTree.Element
		self.IgnoredNode: XmlTree.Element
		self.Columns: Columns

		if self.Path.exists():
			self.Load()
		else:
			self.CreateEmpty()
			self.Save()

	def CreateEmpty(self) -> None:
		self.Root = XmlTree.Element("dgm_database", {"version": DATABASE_VERSION})
		self.SettingsNode = XmlTree.SubElement(self.Root, "settings")
		self.ColumnsNode = XmlTree.SubElement(self.SettingsNode, "columns", copy.deepcopy(DEFAULT_COLUMNS))
		self.CatalogNode = XmlTree.SubElement(self.Root, "catalog")
		self.IgnoredNode = XmlTree.SubElement(self.Root, "ignored")
		self.Tree = XmlTree.ElementTree(self.Root)
		self.Columns = self.ReadColumns()

	def Load(self) -> None:
		try:
			self.Tree = XmlTree.parse(self.Path)
		except XmlTree.ParseError as Error:
			raise RuntimeError(f"Cannot parse XML database '{self.Path}': {Error}") from Error

		self.Root = self.Tree.getroot()
		if self.Root.tag != "dgm_database":
			raise RuntimeError(f"Unexpected XML root in '{self.Path}': {self.Root.tag}")

		self.Root.set("version", self.Root.get("version", DATABASE_VERSION))
		self.SettingsNode = self.EnsureChild(self.Root, "settings")
		self.ColumnsNode = self.EnsureChild(self.SettingsNode, "columns", DEFAULT_COLUMNS)
		self.CatalogNode = self.EnsureChild(self.Root, "catalog")
		for ElementsNode in self.Root.findall("elements"):
			self.Root.remove(ElementsNode)
		self.IgnoredNode = self.EnsureChild(self.Root, "ignored")
		self.Columns = self.ReadColumns()

	def EnsureChild(self, Parent: XmlTree.Element, Tag: str, Attributes: Optional[Dict[str, str]] = None) -> XmlTree.Element:
		Existing = Parent.find(Tag)
		if Existing is not None:
			if Attributes is not None:
				for Key, Value in Attributes.items():
					Existing.set(Key, Existing.get(Key, Value))
			return Existing
		return XmlTree.SubElement(Parent, Tag, Attributes or {})

	def ReadColumns(self) -> Columns:
		def ReadColumn(Attribute: str) -> str:
			Column = self.ColumnsNode.get(Attribute, DEFAULT_COLUMNS[Attribute]).strip().upper()
			try:
				if openpyxl is not None:
					openpyxl.utils.column_index_from_string(Column)  # type: ignore[union-attr]
				elif not IsSpreadsheetColumnName(Column):
					raise ValueError("column name must contain only letters")
			except Exception as Error:
				raise RuntimeError(f"Invalid column '{Column}' in database setting '{Attribute}'") from Error
			return Column

		PerElement = {}
		Total = {}
		for MetalKey, _ in METALS:
			PerElement[MetalKey] = ReadColumn(MetalKey)
			Total[MetalKey] = ReadColumn(f"total_{MetalKey}")

		return Columns(
			Name=ReadColumn("name"),
			Quantity=ReadColumn("quantity"),
			PerElement=PerElement,
			Total=Total,
		)

	def FindElement(self, Name: str) -> ElementSearchResult:
		NormalizedName = NormalizeText(Name)
		if not NormalizedName:
			return ElementSearchResult(None, False)

		StructuredResult = self.FindStructuredElement(NormalizedName, Name)
		if StructuredResult.Record is not None:
			return StructuredResult

		return StructuredResult

	def FindStructuredElement(self, NormalizedName: str, OriginalName: str) -> ElementSearchResult:
		States: List[Tuple[XmlTree.Element, str, bool, List[str]]] = [(self.CatalogNode, NormalizedName, False, [])]
		BestRecord: Optional[ElementRecord] = None
		BestRegexState = False
		PartialMatchesByNode: Dict[int, PartialElementMatch] = {}
		DeepestPartialLength = 0

		while States:
			Parent, Remaining, MatchedRegex, PathTexts = States.pop()
			MatchedLength = len(NormalizedName) - len(Remaining)
			DgmNode = Parent.find("dgm")

			if PathTexts and Remaining != "" and MatchedLength > 0:
				if MatchedLength > DeepestPartialLength:
					PartialMatchesByNode.clear()
					DeepestPartialLength = MatchedLength
				if MatchedLength == DeepestPartialLength:
					PartialMatchesByNode[id(Parent)] = PartialElementMatch(
						Node=Parent,
						DisplayName=" => ".join(PathTexts),
						MatchedByRegex=MatchedRegex,
						HasDgm=DgmNode is not None,
						Remainder=self._GetOriginalRemainder(NormalizedName, OriginalName, MatchedLength),
					)

			if Remaining == "" and PathTexts:
				EmptyRegexRecord = self._FindEmptyRegexChildRecord(Parent, PathTexts, OriginalName)
				if EmptyRegexRecord is not None:
					BestRecord = EmptyRegexRecord
					BestRegexState = True
					break
				BestRecord = self.MakeRecord(Parent, OriginalName or "".join(PathTexts), DgmNode)
				BestRegexState = MatchedRegex
				break

			for Child in list(Parent):
				if Child.tag == "node":
					ChildText = Child.get("text", "")
					ChildKey = NormalizeStructureText(ChildText)
					if ChildKey and Remaining.startswith(ChildKey):
						States.append((Child, Remaining[len(ChildKey):], MatchedRegex, PathTexts + [ChildText]))
					elif IsOptionalNode(Child):
						States.append((Child, Remaining, MatchedRegex, PathTexts + [FormatOptionalPathText(ChildText)]))
				elif Child.tag == "regex":
					Pattern = Child.get("pattern", "")
					if not Pattern:
						continue
					try:
						Match = re.match(Pattern, Remaining, flags=re.IGNORECASE)
					except re.error:
						continue
					if Match is None:
						continue
					MatchedPart = Match.group(0)
					if MatchedPart == "" and Remaining != "":
						continue
					States.append((Child, Remaining[len(MatchedPart):], True, PathTexts + [MatchedPart or Child.get("text", Pattern)]))
		return ElementSearchResult(BestRecord, BestRegexState, list(PartialMatchesByNode.values()))

	def _GetOriginalRemainder(self, NormalizedName: str, OriginalName: str, MatchedLength: int) -> str:
		if MatchedLength <= 0:
			return OriginalName
		if MatchedLength >= len(NormalizedName):
			return ""
		if len(OriginalName) >= MatchedLength and NormalizeText(OriginalName[:MatchedLength]) == NormalizedName[:MatchedLength]:
			return OriginalName[MatchedLength:]
		return NormalizedName[MatchedLength:]

	def _FindEmptyRegexChildRecord(self, Parent: XmlTree.Element, PathTexts: List[str], OriginalName: str) -> Optional[ElementRecord]:
		for Child in list(Parent):
			if Child.tag != "regex":
				continue
			Pattern = Child.get("pattern", "")
			DgmNode = Child.find("dgm")
			if not Pattern or DgmNode is None:
				continue
			try:
				Match = re.match(Pattern, "", flags=re.IGNORECASE)
			except re.error:
				continue
			if Match is None or Match.group(0) != "":
				continue
			return self.MakeRecord(Child, OriginalName or "".join(PathTexts), DgmNode)
		return None

	def MakeRecord(self, Node: XmlTree.Element, DisplayName: str, DgmNode: Optional[XmlTree.Element]) -> ElementRecord:
		if DgmNode is None:
			return ElementRecord(
				self,
				Node,
				DisplayName,
				DgmValues(decimal.Decimal("0"), decimal.Decimal("0"), decimal.Decimal("0"), decimal.Decimal("0")),
				False,
			)
		return ElementRecord(self, Node, DisplayName, self.ReadDgmValues(DgmNode), True)

	def ReadDgmValues(self, DgmNode: XmlTree.Element) -> DgmValues:
		return DgmValues(
			GoldG=ReadDecimal(DgmNode.get("gold_g", "0")),
			SilverG=ReadDecimal(DgmNode.get("silver_g", "0")),
			PlatinumG=ReadDecimal(DgmNode.get("platinum_g", "0")),
			MpgG=ReadDecimal(DgmNode.get("mpg_g", "0")),
		)

	def AddElement(self, Name: str, Values: DgmValues, PathParts: Sequence[str]) -> ElementRecord:
		CleanParts = [Part for Part in PathParts if Part != ""]
		if not CleanParts:
			CleanParts = [Name]

		Parent = self.CatalogNode
		for Part in CleanParts:
			Parent = self.EnsureRegularNode(Parent, Part)

		Parent.set("name", Parent.get("name", Name))
		Record = ElementRecord(self, Parent, Name, Values)
		Record.WriteValuesToXml()
		return Record

	def AddRegexElement(self, Name: str, Values: DgmValues, ParentPathParts: Sequence[str], Pattern: str, DisplayText: str = "") -> ElementRecord:
		Parent = self.FindOrCreateParent(ParentPathParts)
		RegexNode = XmlTree.SubElement(
			Parent,
			"regex",
			{
				"pattern": Pattern,
				"text": DisplayText or Pattern,
				"name": Name,
			},
		)
		Record = ElementRecord(self, RegexNode, Name, Values)
		Record.WriteValuesToXml()
		return Record

	def ConvertSiblingToRegex(self, ParentPathParts: Sequence[str], SiblingIndex: int, Pattern: str, DisplayText: str = "") -> ElementRecord:
		Parent = self.FindParent(ParentPathParts)
		if Parent is None:
			raise RuntimeError("Parent path does not exist")

		Children = list(Parent)
		if SiblingIndex < 1 or SiblingIndex > len(Children):
			raise RuntimeError("Sibling index is out of range")

		OldNode = Children[SiblingIndex - 1]
		if OldNode.tag not in ("node", "regex"):
			raise RuntimeError("Selected sibling is not an element node")

		OldDgmNode = OldNode.find("dgm")
		if OldDgmNode is None:
			raise RuntimeError("Selected sibling does not have DGM values")

		InsertPosition = Children.index(OldNode)
		Values = DgmValues(
			GoldG=ReadDecimal(OldDgmNode.get("gold_g", "0")),
			SilverG=ReadDecimal(OldDgmNode.get("silver_g", "0")),
			PlatinumG=ReadDecimal(OldDgmNode.get("platinum_g", "0")),
			MpgG=ReadDecimal(OldDgmNode.get("mpg_g", "0")),
		)

		Parent.remove(OldNode)
		NewNode = XmlTree.Element(
			"regex",
			{
				"pattern": Pattern,
				"text": DisplayText or Pattern,
				"name": OldNode.get("name", OldNode.get("text", Pattern)),
			},
		)
		NewDgmNode = XmlTree.SubElement(NewNode, "dgm")
		NewDgmNode.set("gold_g", DecimalToText(Values.GoldG))
		NewDgmNode.set("silver_g", DecimalToText(Values.SilverG))
		NewDgmNode.set("platinum_g", DecimalToText(Values.PlatinumG))
		NewDgmNode.set("mpg_g", DecimalToText(Values.MpgG))

		Parent.insert(InsertPosition, NewNode)
		return ElementRecord(self, NewNode, NewNode.get("name", Pattern), Values)

	def EnsureRegularNode(self, Parent: XmlTree.Element, Text: str) -> XmlTree.Element:
		Key = NormalizeStructureText(Text)
		for Child in Parent.findall("node"):
			if NormalizeStructureText(Child.get("text", "")) == Key:
				return Child
		return XmlTree.SubElement(Parent, "node", {"text": Text})

	def FindOrCreateParent(self, PathParts: Sequence[str]) -> XmlTree.Element:
		Parent = self.CatalogNode
		for Part in PathParts:
			if Part == "":
				continue
			Parent = self.EnsureRegularNode(Parent, Part)
		return Parent

	def FindParent(self, PathParts: Sequence[str]) -> Optional[XmlTree.Element]:
		Parent = self.CatalogNode
		for Part in PathParts:
			if Part == "":
				continue
			Key = NormalizeStructureText(Part)
			NextNode = None
			for Child in Parent.findall("node"):
				if NormalizeStructureText(Child.get("text", "")) == Key:
					NextNode = Child
					break
			if NextNode is None:
				return None
			Parent = NextNode
		return Parent

	def AddDgmToExistingPath(self, Name: str, Values: DgmValues, PathParts: Sequence[str]) -> ElementRecord:
		Node = self.FindParent(PathParts)
		if Node is None or Node is self.CatalogNode:
			raise RuntimeError("The selected structured path does not exist")
		if Node.tag != "node":
			raise RuntimeError("The selected structured path is not a regular node")

		Node.set("name", Node.get("name", Name))
		Record = ElementRecord(self, Node, Name, Values)
		Record.WriteValuesToXml()
		return Record

	def GetRegularPathInfo(self, PathParts: Sequence[str]) -> Optional[ExistingPathInfo]:
		CleanParts = [Part for Part in (Part.strip() for Part in PathParts) if Part]
		if not CleanParts:
			return None

		Node = self.FindParent(CleanParts)
		if Node is None or Node is self.CatalogNode or Node.tag != "node":
			return None

		DisplayName = Node.get("name", "".join(CleanParts))
		return ExistingPathInfo(
			Index=1,
			PathParts=CleanParts,
			DisplayPath="/".join(CleanParts),
			DisplayName=DisplayName,
			HasDgm=Node.find("dgm") is not None,
		)

	def GetSiblingInfos(self, ParentPathParts: Sequence[str]) -> List[SiblingInfo]:
		Parent = self.FindParent(ParentPathParts)
		if Parent is None:
			return []

		Infos: List[SiblingInfo] = []
		for Index, Child in enumerate(list(Parent), start=1):
			if Child.tag == "node":
				Infos.append(SiblingInfo(Index, "node", Child.get("text", ""), "", Child.find("dgm") is not None))
			elif Child.tag == "regex":
				Infos.append(SiblingInfo(Index, "regex", Child.get("text", ""), Child.get("pattern", ""), Child.find("dgm") is not None))
		return Infos

	def GetNodePathParts(self, Node: XmlTree.Element) -> List[str]:
		PathParts: List[str] = []
		Found = False

		def Walk(Parent: XmlTree.Element, Parts: List[str]) -> None:
			nonlocal PathParts, Found
			if Found:
				return
			for Child in list(Parent):
				if Child.tag not in ("node", "regex"):
					continue
				ChildParts = Parts + [Child.get("text", Child.get("name", Child.tag))]
				if Child is Node:
					PathParts = ChildParts
					Found = True
					return
				Walk(Child, ChildParts)

		Walk(self.CatalogNode, [])
		return PathParts

	def FindCatalogParentOfNode(self, Node: XmlTree.Element) -> Optional[XmlTree.Element]:
		if Node is self.CatalogNode:
			return None

		def Walk(Parent: XmlTree.Element) -> Optional[XmlTree.Element]:
			for Child in list(Parent):
				if Child is Node:
					return Parent
				if Child.tag in ("node", "regex"):
					FoundParent = Walk(Child)
					if FoundParent is not None:
						return FoundParent
			return None

		return Walk(self.CatalogNode)

	def IsCatalogDescendant(self, CandidateParent: XmlTree.Element, Node: XmlTree.Element) -> bool:
		for Child in list(Node):
			if Child is CandidateParent:
				return True
			if Child.tag in ("node", "regex") and self.IsCatalogDescendant(CandidateParent, Child):
				return True
		return False

	def MoveCatalogNode(self, Node: XmlTree.Element, NewParentPathParts: Sequence[str]) -> None:
		if Node is self.CatalogNode or Node.tag not in ("node", "regex"):
			raise RuntimeError("Only catalog node and regex entries can be moved")

		OldParent = self.FindCatalogParentOfNode(Node)
		if OldParent is None:
			raise RuntimeError("Selected catalog entry is not attached to the database")

		NewParent = self.FindOrCreateParent(NewParentPathParts)
		if NewParent is Node or self.IsCatalogDescendant(NewParent, Node):
			raise RuntimeError("Cannot move a catalog entry under itself or one of its children")

		OldParent.remove(Node)
		NewParent.append(Node)

	def CatalogNodeHasNonZeroDgmValues(self, Node: XmlTree.Element) -> bool:
		NodesToCheck = [Node]
		while NodesToCheck:
			Current = NodesToCheck.pop()
			DgmNode = Current.find("dgm")
			SourceNode = DgmNode if DgmNode is not None else Current
			for MetalKey, _ in METALS:
				try:
					if ReadDecimal(SourceNode.get(f"{MetalKey}_g", "0")) != 0:
						return True
				except decimal.InvalidOperation:
					return True
			NodesToCheck.extend(Child for Child in list(Current) if Child.tag in ("node", "regex"))
		return False

	def RemoveCatalogNode(self, Node: XmlTree.Element) -> None:
		if Node is self.CatalogNode or Node.tag not in ("node", "regex"):
			raise RuntimeError("Only catalog node and regex entries can be removed")

		Parent = self.FindCatalogParentOfNode(Node)
		if Parent is None:
			raise RuntimeError("Selected catalog entry is not attached to the database")

		Parent.remove(Node)

	def IsIgnoredText(self, Text: str) -> bool:
		Key = NormalizeText(Text)
		if not Key:
			return True
		for Node in self.IgnoredNode.findall("text"):
			StoredKey = Node.get("key") or NormalizeText(Node.get("value", ""))
			if NormalizeText(StoredKey) == Key:
				return True
		return False

	def AddIgnoredText(self, Text: str) -> None:
		Key = NormalizeText(Text)
		if not Key or self.IsIgnoredText(Text):
			return
		XmlTree.SubElement(self.IgnoredNode, "text", {"key": Key, "value": Text})

	def Save(self) -> None:
		self.Path.parent.mkdir(parents=True, exist_ok=True)
		try:
			XmlTree.indent(self.Tree, space="\t", level=0)
		except AttributeError:
			pass
		self.Tree.write(self.Path, encoding="utf-8", xml_declaration=True)


def OpenDatabase(PathToDatabase: Path | str) -> DgmDatabase:
	return DgmDatabase(Path(PathToDatabase))


def NormalizeText(Value: str) -> str:
	return " ".join(str(Value).strip().split()).casefold()

def NormalizeStructureText(Value: str) -> str:
	return str(Value or "").casefold()

def IsOptionalNode(Node: XmlTree.Element) -> bool:
	return Node.tag == "node" and NormalizeText(Node.get("optional", "")) in ("1", "true", "yes", "on")

def SetOptionalNode(Node: XmlTree.Element, IsOptional: bool) -> None:
	if IsOptional:
		Node.set("optional", "true")
	elif "optional" in Node.attrib:
		del Node.attrib["optional"]

def FormatOptionalPathText(Text: str) -> str:
	return f"{Text} (optional)" if Text else "(optional)"

def ReadDecimal(Value: str) -> decimal.Decimal:
	Value = (Value or "0").strip().replace(",", ".")
	if Value == "":
		Value = "0"
	return decimal.Decimal(Value)


def DecimalToText(Value: decimal.Decimal) -> str:
	if Value == 0:
		return "0"
	return format(Value.normalize(), "f")


def IsSpreadsheetColumnName(Value: str) -> bool:
	return bool(Value) and all(Character in COLUMN_LETTERS for Character in Value)
