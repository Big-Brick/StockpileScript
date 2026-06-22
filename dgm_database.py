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
from dataclasses import dataclass, field
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
	Record: "ElementRecord"
	Remainder: str = ""

	@property
	def Node(self) -> XmlTree.Element:
		return self.Record.Node

	@property
	def DisplayName(self) -> str:
		return self.Record.DisplayName

	@property
	def MatchedByRegex(self) -> bool:
		return self.Record.MatchedByRegex

	@property
	def HasDgm(self) -> bool:
		return self.Record.HasDgm


@dataclass
class ElementSearchResult:
	Record: Optional["ElementRecord"]
	MatchedByRegex: bool = False
	PartialMatches: List[PartialElementMatch] = field(default_factory=list)


@dataclass
class SiblingInfo:
	Index: int
	Kind: str
	Text: str
	Pattern: str
	HasDgm: bool


class ElementRecord:
	def __init__(
		self,
		Database: "DgmDatabase",
		Node: XmlTree.Element,
		DisplayText: str,
		Values: DgmValues,
		HasDgm: bool = True,
		Parent: Optional["ElementRecord"] = None,
		PathText: str = "",
		ConsumedText: str = "",
		MatchedByRegex: bool = False,
		IsRoot: bool = False,
	) -> None:
		self.Database = Database
		self.Node = Node
		self.DisplayText = DisplayText
		self.Values = Values
		self.HasDgm = HasDgm
		self.Parent = Parent
		self.PathText = PathText or DisplayText
		self.ConsumedText = ConsumedText
		self.MatchedByRegex = MatchedByRegex
		self.IsRoot = IsRoot

	@property
	def DisplayName(self) -> str:
		return self.DisplayPath

	@property
	def HasNonZeroDgm(self) -> bool:
		return self.HasDgm and not self.Values.IsZero()

	@property
	def IsUsableForInventoryFill(self) -> bool:
		return self.HasDgm

	def IterPath(self) -> List["ElementRecord"]:
		Path: List[ElementRecord] = []
		Current: Optional[ElementRecord] = self
		while Current is not None:
			if not Current.IsRoot:
				Path.append(Current)
			Current = Current.Parent
		Path.reverse()
		return Path

	@property
	def PathParts(self) -> List[str]:
		return [Record.PathText for Record in self.IterPath()]

	@property
	def DisplayPathParts(self) -> List[str]:
		return [Record.DisplayText for Record in self.IterPath()]

	@property
	def DisplayPath(self) -> str:
		return " => ".join(self.DisplayPathParts)

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
		RootRecord = ElementRecord(
			self,
			self.CatalogNode,
			"catalog",
			DgmValues(decimal.Decimal("0"), decimal.Decimal("0"), decimal.Decimal("0"), decimal.Decimal("0")),
			False,
			None,
			"",
			"",
			False,
			True,
		)
		States: List[Tuple[XmlTree.Element, str, ElementRecord]] = [(self.CatalogNode, NormalizedName, RootRecord)]
		BestRecord: Optional[ElementRecord] = None
		BestRegexState = False
		PartialMatchesByNode: Dict[int, PartialElementMatch] = {}
		DeepestPartialLength = 0

		while States:
			Parent, Remaining, ParentRecord = States.pop()
			MatchedLength = len(NormalizedName) - len(Remaining)

			if not ParentRecord.IsRoot and Remaining != "" and MatchedLength > 0:
				if MatchedLength > DeepestPartialLength:
					PartialMatchesByNode.clear()
					DeepestPartialLength = MatchedLength
				if MatchedLength == DeepestPartialLength:
					PartialMatchesByNode[id(Parent)] = PartialElementMatch(
						Record=ParentRecord,
						Remainder=self._GetOriginalRemainder(NormalizedName, OriginalName, MatchedLength),
					)

			if Remaining == "" and not ParentRecord.IsRoot:
				EmptyRegexRecord = self._FindEmptyRegexChildRecord(ParentRecord)
				if EmptyRegexRecord is not None:
					BestRecord = EmptyRegexRecord
					BestRegexState = True
					break
				BestRecord = ParentRecord
				BestRegexState = ParentRecord.MatchedByRegex
				break

			for Child in list(Parent):
				if Child.tag == "node":
					ChildText = Child.get("text", "")
					ChildKey = NormalizeStructureText(ChildText)
					if ChildKey and Remaining.startswith(ChildKey):
						ChildRecord = self.MakeRecord(
							Child,
							ChildText,
							ParentRecord,
							ChildText,
							ChildText,
							ParentRecord.MatchedByRegex,
						)
						States.append((Child, Remaining[len(ChildKey):], ChildRecord))
					elif IsOptionalNode(Child):
						ChildRecord = self.MakeRecord(
							Child,
							FormatOptionalPathText(ChildText),
							ParentRecord,
							ChildText,
							"",
							ParentRecord.MatchedByRegex,
						)
						States.append((Child, Remaining, ChildRecord))
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
					PathText = Child.get("text", Child.get("name", Pattern))
					ChildRecord = self.MakeRecord(
						Child,
						MatchedPart or PathText,
						ParentRecord,
						PathText,
						MatchedPart,
						True,
					)
					States.append((Child, Remaining[len(MatchedPart):], ChildRecord))
		return ElementSearchResult(BestRecord, BestRegexState, list(PartialMatchesByNode.values()))

	def _GetOriginalRemainder(self, NormalizedName: str, OriginalName: str, MatchedLength: int) -> str:
		if MatchedLength <= 0:
			return OriginalName
		if MatchedLength >= len(NormalizedName):
			return ""
		if len(OriginalName) >= MatchedLength and NormalizeText(OriginalName[:MatchedLength]) == NormalizedName[:MatchedLength]:
			return OriginalName[MatchedLength:]
		return NormalizedName[MatchedLength:]

	def _FindEmptyRegexChildRecord(self, ParentRecord: ElementRecord) -> Optional[ElementRecord]:
		for Child in list(ParentRecord.Node):
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
			PathText = Child.get("text", Child.get("name", Pattern))
			return self.MakeRecord(Child, PathText, ParentRecord, PathText, "", True)
		return None

	def MakeRecord(
		self,
		Node: XmlTree.Element,
		DisplayText: str,
		Parent: Optional[ElementRecord] = None,
		PathText: str = "",
		ConsumedText: str = "",
		MatchedByRegex: bool = False,
	) -> ElementRecord:
		DgmNode = Node.find("dgm")
		if DgmNode is None:
			return ElementRecord(
				self,
				Node,
				DisplayText,
				DgmValues(decimal.Decimal("0"), decimal.Decimal("0"), decimal.Decimal("0"), decimal.Decimal("0")),
				False,
				Parent,
				PathText,
				ConsumedText,
				MatchedByRegex,
			)
		return ElementRecord(self, Node, DisplayText, self.ReadDgmValues(DgmNode), True, Parent, PathText, ConsumedText, MatchedByRegex)

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
