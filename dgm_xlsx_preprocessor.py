from __future__ import annotations

import re
import xml.etree.ElementTree as XmlTree
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import dgm_database
import dgm_xlsx_common

try:
	import openpyxl
except ImportError:
	openpyxl = None  # type: ignore[assignment]


RULES_VERSION = "1"
DEFAULT_RULES_FILENAME = "preprocess_rules.xml"


@dataclass
class PreprocessRegexRule:
	Id: str
	Pattern: str
	Replacement: str
	Description: str = ""
	Enabled: bool = True


@dataclass
class PreprocessElementType:
	Id: str
	Canonical: str
	Aliases: List[str] = field(default_factory=list)
	Prefixes: List[str] = field(default_factory=list)
	DatabaseCheck: bool = True
	XmlNode: Optional[XmlTree.Element] = None


@dataclass
class PreprocessRules:
	Path: Path
	Tree: XmlTree.ElementTree
	Root: XmlTree.Element
	IgnoreCase: bool = True
	CollapseWhitespace: bool = True
	TrimWhitespace: bool = True
	AutoUpdateCache: bool = True
	IgnoredTextRules: List[PreprocessRegexRule] = field(default_factory=list)
	StageOneRules: List[PreprocessRegexRule] = field(default_factory=list)
	StageThreeRules: List[PreprocessRegexRule] = field(default_factory=list)
	ElementTypes: List[PreprocessElementType] = field(default_factory=list)


@dataclass
class PreprocessChange:
	Row: int
	OriginalText: str
	NewText: str
	StageNotes: List[str]
	DatabaseVerified: bool = False
	Ambiguous: bool = False
	StageId: str = ""
	ElementType: str = ""
	SafetyLevel: str = ""
	FilePath: Path = Path()
	SheetName: str = ""
	RulesPath: Path = Path()
	StageName: str = "All stages"


@dataclass
class PreprocessStage:
	Id: str
	Name: str


PREPROCESS_STAGES: Tuple[PreprocessStage, ...] = (
	PreprocessStage("language", "Language normalization"),
	PreprocessStage("prefix", "Add element prefixes"),
	PreprocessStage("technical", "Technical normalization"),
)

PREFIX_SAFETY_ORDER = {"safe": 0, "partial": 1, "ambiguous": 2, "unidentified": 3}
PREFIX_SAFETY_LABELS = {
	"safe": "Safe exact database match",
	"partial": "Partially safe partial database match",
	"ambiguous": "Ambiguous element type",
	"unidentified": "Unidentified by rules",
}


class XlsxPreprocessor:
	def __init__(self, Database: dgm_database.DgmDatabase, RulesPath: Path) -> None:
		self.Database = Database
		self.Rules = LoadPreprocessRules(RulesPath)
		self.RegexFlags = re.IGNORECASE if self.Rules.IgnoreCase else 0
		self.CacheChanged = False

	def PreprocessWorkbook(self, FilePath: Path, StageId: str = "all") -> List[PreprocessChange]:
		if openpyxl is None:
			raise RuntimeError("Missing dependency: openpyxl")

		Workbook = openpyxl.load_workbook(FilePath, data_only=False)
		Sheet = Workbook.active
		Changes: List[PreprocessChange] = []
		ConsecutiveIgnoredRows = 0
		Row = 1
		MaxRow = Sheet.max_row or 1

		while Row <= MaxRow:
			RawName = Sheet[f"{self.Database.Columns.Name}{Row}"].value
			if not dgm_xlsx_common.CellHasUsableText(RawName):
				ConsecutiveIgnoredRows += 1
			else:
				Original = str(RawName)
				if self.IsIgnoredText(Original):
					ConsecutiveIgnoredRows += 1
				else:
					Change = self.PreprocessText(Row, Original, StageId)
					self._AttachChangeMetadata(Change, FilePath, Sheet.title, StageId)
					if self._ShouldOfferChange(Change, StageId):
						Changes.append(Change)
					ConsecutiveIgnoredRows = 0

			if ConsecutiveIgnoredRows >= dgm_xlsx_common.STOP_AFTER_CONSECUTIVE_IGNORED_ROWS:
				break
			Row += 1

		return Changes

	def _AttachChangeMetadata(self, Change: PreprocessChange, FilePath: Path, SheetName: str, StageId: str) -> PreprocessChange:
		Change.FilePath = FilePath
		Change.SheetName = SheetName
		Change.RulesPath = self.Rules.Path
		Change.StageId = StageId
		Change.StageName = self.GetStageName(StageId)
		return Change

	def IsIgnoredText(self, Text: str) -> bool:
		Current = self._CleanWhitespace(Text)
		for Rule in self.Rules.IgnoredTextRules:
			if not Rule.Enabled or not Rule.Pattern:
				continue
			try:
				if re.search(Rule.Pattern, Current, flags=self.RegexFlags):
					return True
			except re.error:
				continue
		return False

	def ApplyChanges(self, Changes: Sequence[PreprocessChange]) -> None:
		if openpyxl is None:
			raise RuntimeError("Missing dependency: openpyxl")
		if not Changes:
			return

		ChangesBySheet: Dict[Tuple[Path, str], List[PreprocessChange]] = {}
		for Change in Changes:
			ChangesBySheet.setdefault((Change.FilePath, Change.SheetName), []).append(Change)

		for (FilePath, SheetName), SheetChanges in ChangesBySheet.items():
			Workbook = openpyxl.load_workbook(FilePath, data_only=False)
			Sheet = Workbook[SheetName]
			for Change in SheetChanges:
				Sheet[f"{self.Database.Columns.Name}{Change.Row}"].value = Change.NewText
			Workbook.save(FilePath)
		if self.CacheChanged:
			SavePreprocessRules(self.Rules)

	def PreprocessText(self, Row: int, Text: str, StageId: str = "all") -> PreprocessChange:
		if StageId == "language":
			return self._PreprocessLanguage(Row, Text)
		if StageId == "prefix":
			return self._PreprocessPrefix(Row, Text)
		if StageId == "technical":
			return self._PreprocessTechnical(Row, Text)

		Notes: List[str] = []
		Original = Text
		Current = self._CleanWhitespace(Text)
		if Current != Text:
			Notes.append("Cleaned whitespace")
		if self.IsIgnoredText(Current):
			Notes.append("Ignored block/header row")
			return PreprocessChange(Row, Original, Current, Notes, False, False)

		Current = self._ApplyRules(Current, self.Rules.StageOneRules, "Stage 1", Notes)
		Current, TypeVerified, TypeAmbiguous = self._NormalizeType(Current, Notes)
		Current = self._ApplyRules(Current, self.Rules.StageThreeRules, "Stage 3", Notes)
		CaseNormalized = NormalizeDesignationCase(Current)
		if CaseNormalized != Current:
			Notes.append("Stage 3: normalized designation letter casing")
			Current = CaseNormalized
		Current = self._CleanWhitespace(Current)
		Current = self._ApplyExactDatabaseCasing(Current, Notes)

		DatabaseRecord = self.Database.FindElement(Current).Record
		DatabaseVerified = TypeVerified or (DatabaseRecord is not None and DatabaseRecord.HasDgm)
		if DatabaseVerified and not TypeVerified:
			Notes.append("Verified final text against database")
		if not Notes and Original == Current:
			Notes.append("No change")
		return PreprocessChange(Row, Original, Current, Notes, DatabaseVerified, TypeAmbiguous, "all")

	def _PreprocessLanguage(self, Row: int, Text: str) -> PreprocessChange:
		Notes: List[str] = []
		Original = Text
		Current = self._CleanWhitespace(Text)
		if Current != Text:
			Notes.append("Cleaned whitespace")
		Current = self._ApplyRules(Current, self.Rules.StageOneRules, "Stage 1", Notes)
		return PreprocessChange(Row, Original, Current, Notes, True, False, "language")

	def _PreprocessTechnical(self, Row: int, Text: str) -> PreprocessChange:
		Notes: List[str] = []
		Original = Text
		Current = self._CleanWhitespace(Text)
		if Current != Text:
			Notes.append("Cleaned whitespace")
		Current = self._ApplyRules(Current, self.Rules.StageThreeRules, "Stage 3", Notes)
		CaseNormalized = NormalizeDesignationCase(Current)
		if CaseNormalized != Current:
			Notes.append("Stage 3: normalized designation letter casing")
			Current = CaseNormalized
		Current = self._CleanWhitespace(Current)
		Current = self._ApplyExactDatabaseCasing(Current, Notes)
		DatabaseRecord = self.Database.FindElement(Current).Record
		Verified = DatabaseRecord is not None and DatabaseRecord.HasDgm
		return PreprocessChange(Row, Original, Current, Notes, Verified, False, "technical")

	def _ApplyExactDatabaseCasing(self, Text: str, Notes: List[str]) -> str:
		Result = self.Database.FindElement(Text)
		Record = Result.Record
		if Record is None or not Record.HasDgm or Result.MatchedByRegex:
			return Text
		DatabaseText = BuildExactDatabaseText(Record)
		if not DatabaseText:
			return Text
		if Text != DatabaseText and Text.casefold() == DatabaseText.casefold() and len(Text) == len(DatabaseText):
			Notes.append("Stage 3: matched exact database letter casing")
			return DatabaseText
		return Text


	def _PreprocessPrefix(self, Row: int, Text: str) -> PreprocessChange:
		Original = Text
		Current = self._CleanWhitespace(Text)
		ExistingType = self._FindExplicitType(Current)
		ExistingResult = self.Database.FindElement(Current)
		ExistingRecord = ExistingResult.Record
		if (
			ExistingType is not None
			and ExistingRecord is not None
			and ExistingRecord.HasDgm
			and not ExistingResult.MatchedByRegex
			and self._RecordMatchesElementType(ExistingRecord, ExistingType[0])
		):
			return PreprocessChange(Row, Original, Current, ["Stage 2: existing prefixed text verified by exact database match"], True, False, "prefix", ExistingType[0].Canonical, "safe")
		elif ExistingType is not None:
			ElementType, Remainder = ExistingType
			Candidate = self._CleanWhitespace(f"{ElementType.Canonical} {Remainder}")
			Safety = self._ClassifyPrefixCandidate(Candidate, ElementType)
			if Safety in ("safe", "partial") and Candidate == Current:
				return PreprocessChange(Row, Original, Current, [f"Stage 2: existing prefix accepted ({PREFIX_SAFETY_LABELS[Safety]})"], Safety == "safe", False, "prefix", ElementType.Canonical, Safety)
			return PreprocessChange(Row, Original, Candidate, [f"Stage 2: normalized explicit type as {ElementType.Canonical}"], Safety == "safe", Safety == "ambiguous", "prefix", ElementType.Canonical, Safety)

		PrefixMatchesByCache: List[Tuple[PreprocessElementType, str, str]] = []
		Token = ExtractLeadingPrefix(Current)
		if Token:
			for ElementType in self.Rules.ElementTypes:
				if any(PrefixMatches(Token, Prefix) for Prefix in ElementType.Prefixes):
					Candidate = self._MakeTypedCandidate(ElementType, Current)
					Safety = self._ClassifyPrefixCandidate(Candidate, ElementType)
					if Safety in ("safe", "partial"):
						PrefixMatchesByCache.append((ElementType, Candidate, Safety))
			if len(PrefixMatchesByCache) == 1:
				ElementType, Candidate, Safety = PrefixMatchesByCache[0]
				return PreprocessChange(Row, Original, Candidate, [f"Stage 2: inferred {ElementType.Canonical} from cached prefix {Token}"], Safety == "safe", False, "prefix", ElementType.Canonical, Safety)
			if len(PrefixMatchesByCache) > 1:
				Types = ", ".join(Match[0].Canonical for Match in PrefixMatchesByCache)
				return PreprocessChange(Row, Original, Current, [f"Stage 2: ambiguous cached prefix {Token}: {Types}"], False, True, "prefix", "", "ambiguous")

		ExactMatches: List[Tuple[PreprocessElementType, str]] = []
		PartialMatches: List[Tuple[PreprocessElementType, str]] = []
		for ElementType in self.Rules.ElementTypes:
			Candidate = self._MakeTypedCandidate(ElementType, Current)
			Safety = self._ClassifyPrefixCandidate(Candidate, ElementType)
			if Safety == "safe":
				ExactMatches.append((ElementType, Candidate))
			elif Safety == "partial":
				PartialMatches.append((ElementType, Candidate))

		if len(ExactMatches) == 1:
			ElementType, Candidate = ExactMatches[0]
			self._RememberPrefix(ElementType, Current, [])
			return PreprocessChange(Row, Original, Candidate, [f"Stage 2: inferred {ElementType.Canonical} by exact database match"], True, False, "prefix", ElementType.Canonical, "safe")
		if len(ExactMatches) > 1:
			Types = ", ".join(Match[0].Canonical for Match in ExactMatches)
			return PreprocessChange(Row, Original, Current, [f"Stage 2: ambiguous exact database matches: {Types}"], False, True, "prefix", "", "ambiguous")
		if len(PartialMatches) == 1:
			ElementType, Candidate = PartialMatches[0]
			return PreprocessChange(Row, Original, Candidate, [f"Stage 2: inferred {ElementType.Canonical} by partial database match"], False, False, "prefix", ElementType.Canonical, "partial")
		if len(PartialMatches) > 1:
			Types = ", ".join(Match[0].Canonical for Match in PartialMatches)
			return PreprocessChange(Row, Original, Current, [f"Stage 2: ambiguous partial database matches: {Types}"], False, True, "prefix", "", "ambiguous")
		return PreprocessChange(Row, Original, Current, ["Stage 2: element type was not identified by rules"], False, True, "prefix", "", "unidentified")

	def _ShouldOfferChange(self, Change: PreprocessChange, StageId: str) -> bool:
		if StageId == "prefix":
			if Change.NewText == Change.OriginalText and Change.ElementType:
				return False
			return Change.SafetyLevel not in ("safe", "partial") or Change.NewText != Change.OriginalText
		return Change.NewText != Change.OriginalText

	def _ClassifyPrefixCandidate(self, Candidate: str, ElementType: PreprocessElementType) -> str:
		Result = self.Database.FindElement(Candidate)
		if (
			Result.Record is not None
			and Result.Record.HasDgm
			and not Result.MatchedByRegex
			and self._RecordMatchesElementType(Result.Record, ElementType)
		):
			return "safe"
		for Match in Result.PartialMatches:
			if not self._RecordMatchesElementType(Match.Record, ElementType):
				continue
			NonOptionalNodes = sum(1 for Record in Match.Record.IterPath() if not dgm_database.IsOptionalNode(Record.Node))
			if NonOptionalNodes >= 2:
				return "partial"
		return "unidentified"

	def _RecordMatchesElementType(self, Record: dgm_database.ElementRecord, ElementType: PreprocessElementType) -> bool:
		Canonical = dgm_database.NormalizeText(ElementType.Canonical)
		DisplayName = dgm_database.NormalizeText(Record.DisplayName)
		return DisplayName == Canonical or DisplayName.startswith(Canonical + " ")

	def GetStageName(self, StageId: str) -> str:
		for Stage in PREPROCESS_STAGES:
			if Stage.Id == StageId:
				return Stage.Name
		return "All stages"


	def _CleanWhitespace(self, Text: str) -> str:
		Result = Text
		if self.Rules.TrimWhitespace:
			Result = Result.strip()
		if self.Rules.CollapseWhitespace:
			Result = " ".join(Result.split())
		return Result

	def _ApplyRules(self, Text: str, Rules: Sequence[PreprocessRegexRule], StageName: str, Notes: List[str]) -> str:
		Current = Text
		for Rule in Rules:
			if not Rule.Enabled:
				continue
			try:
				Updated = re.sub(Rule.Pattern, ConvertReplacement(Rule.Replacement), Current, flags=self.RegexFlags)
			except re.error as Error:
				Notes.append(f"{StageName}: skipped invalid rule {Rule.Id}: {Error}")
				continue
			if Updated != Current:
				Notes.append(f"{StageName}: applied {Rule.Id}")
				Current = Updated
		return Current

	def _NormalizeType(self, Text: str, Notes: List[str]) -> Tuple[str, bool, bool]:
		Explicit = self._FindExplicitType(Text)
		if Explicit is not None:
			ElementType, Remainder = Explicit
			Updated = self._CleanWhitespace(f"{ElementType.Canonical} {Remainder}")
			if Updated != Text:
				Notes.append(f"Stage 2: moved/capitalized type as {ElementType.Canonical}")
			return Updated, self._DatabaseAccepts(ElementType, Updated), False

		Token = ExtractLeadingPrefix(Text)
		if Token:
			for ElementType in self.Rules.ElementTypes:
				if any(PrefixMatches(Token, Prefix) for Prefix in ElementType.Prefixes):
					Candidate = self._MakeTypedCandidate(ElementType, Text)
					if self._DatabaseAccepts(ElementType, Candidate):
						Notes.append(f"Stage 2: inferred {ElementType.Canonical} from cached prefix {Token}")
						return Candidate, True, False

		Matches: List[Tuple[PreprocessElementType, str]] = []
		for ElementType in self.Rules.ElementTypes:
			Candidate = self._MakeTypedCandidate(ElementType, Text)
			if self._DatabaseAccepts(ElementType, Candidate):
				Matches.append((ElementType, Candidate))

		if len(Matches) == 1:
			ElementType, Candidate = Matches[0]
			Notes.append(f"Stage 2: inferred {ElementType.Canonical} by database search")
			self._RememberPrefix(ElementType, Text, Notes)
			return Candidate, True, False
		if len(Matches) > 1:
			Notes.append("Stage 2: ambiguous database type inference")
			return Text, False, True
		return Text, False, False

	def _FindExplicitType(self, Text: str) -> Optional[Tuple[PreprocessElementType, str]]:
		for ElementType in self.Rules.ElementTypes:
			Aliases = [ElementType.Canonical] + ElementType.Aliases
			for Alias in sorted(Aliases, key=len, reverse=True):
				Pattern = re.compile(rf"(^|\s)({re.escape(Alias)})(\s|$)", flags=self.RegexFlags)
				Match = Pattern.search(Text)
				if Match is None:
					continue
				Remainder = self._CleanWhitespace((Text[:Match.start(2)] + " " + Text[Match.end(2):]).strip())
				return ElementType, Remainder
		return None

	def _MakeTypedCandidate(self, ElementType: PreprocessElementType, Text: str) -> str:
		Candidate = self._CleanWhitespace(f"{ElementType.Canonical} {Text}")
		Candidate = self._ApplyRules(Candidate, self.Rules.StageThreeRules, "Stage 3", [])
		return self._CleanWhitespace(NormalizeDesignationCase(Candidate))

	def _DatabaseAccepts(self, ElementType: PreprocessElementType, Candidate: str) -> bool:
		if not ElementType.DatabaseCheck:
			return True
		Result = self.Database.FindElement(Candidate)
		Record = Result.Record
		return (
			Record is not None
			and Record.HasDgm
			and not Result.MatchedByRegex
			and self._RecordMatchesElementType(Record, ElementType)
		)

	def _RememberPrefix(self, ElementType: PreprocessElementType, Text: str, Notes: List[str]) -> None:
		if not self.Rules.AutoUpdateCache or ElementType.XmlNode is None:
			return
		Prefix = ExtractLeadingPrefix(Text)
		if not Prefix or any(PrefixMatches(Prefix, Existing) for Existing in ElementType.Prefixes):
			return
		CacheNode = ElementType.XmlNode.find("prefix_cache")
		if CacheNode is None:
			CacheNode = XmlTree.SubElement(ElementType.XmlNode, "prefix_cache")
		XmlTree.SubElement(CacheNode, "prefix").text = Prefix
		ElementType.Prefixes.append(Prefix)
		self.CacheChanged = True
		Notes.append(f"Stage 2: cached new prefix {Prefix}")


def BuildExactDatabaseText(Record: dgm_database.ElementRecord) -> str:
	Parts: List[str] = []
	for PathRecord in Record.IterPath():
		if not PathRecord.ConsumedText:
			continue
		if PathRecord.MatchedByRegex:
			Parts.append(PathRecord.ConsumedText)
		else:
			Parts.append(PathRecord.DisplayText)
	return "".join(Parts)


def LoadPreprocessRules(PathToRules: Path) -> PreprocessRules:
	if not PathToRules.exists():
		CreateDefaultPreprocessRules(PathToRules)

	try:
		Tree = XmlTree.parse(PathToRules)
	except XmlTree.ParseError as Error:
		raise RuntimeError(f"Cannot parse preprocessing XML '{PathToRules}': {Error}") from Error

	Root = Tree.getroot()
	if Root.tag != "xlsx_preprocess_rules":
		raise RuntimeError(f"Unexpected preprocessing XML root in '{PathToRules}': {Root.tag}")

	Rules = PreprocessRules(Path=PathToRules, Tree=Tree, Root=Root)
	Settings = Root.find("settings")
	if Settings is not None:
		RegexFlags = Settings.find("regex_flags")
		Whitespace = Settings.find("whitespace")
		Cache = Settings.find("cache")
		if RegexFlags is not None:
			Rules.IgnoreCase = ReadBool(RegexFlags.get("ignore_case"), True)
		if Whitespace is not None:
			Rules.CollapseWhitespace = ReadBool(Whitespace.get("collapse"), True)
			Rules.TrimWhitespace = ReadBool(Whitespace.get("trim"), True)
		if Cache is not None:
			Rules.AutoUpdateCache = ReadBool(Cache.get("auto_update"), True)

	IgnoredTextsNode = Root.find("ignored_texts")
	if IgnoredTextsNode is not None:
		Rules.IgnoredTextRules.extend(ParseIgnoreRule(Node) for Node in IgnoredTextsNode.findall("pattern"))

	for Stage in Root.findall("stage"):
		StageName = Stage.get("name", "")
		ParsedRules = [ParseRegexRule(Node) for Node in Stage.findall("rule")]
		if StageName == "language_normalization" or Stage.get("id") == "1":
			Rules.StageOneRules.extend(ParsedRules)
		elif StageName == "technical_normalization" or Stage.get("id") == "3":
			Rules.StageThreeRules.extend(ParsedRules)

	ElementTypesNode = Root.find("element_types")
	if ElementTypesNode is not None:
		for TypeNode in ElementTypesNode.findall("type"):
			Canonical = TypeNode.get("canonical", "").strip()
			TypeId = TypeNode.get("id", Canonical).strip()
			if not Canonical:
				continue
			AliasesNode = TypeNode.find("aliases")
			PrefixNode = TypeNode.find("prefix_cache")
			DatabaseCheckNode = TypeNode.find("database_check")
			Aliases = [Child.text.strip() for Child in (AliasesNode.findall("alias") if AliasesNode is not None else []) if Child.text and Child.text.strip()]
			Prefixes = [Child.text.strip() for Child in (PrefixNode.findall("prefix") if PrefixNode is not None else []) if Child.text and Child.text.strip()]
			Rules.ElementTypes.append(
				PreprocessElementType(
					Id=TypeId,
					Canonical=Canonical,
					Aliases=Aliases,
					Prefixes=Prefixes,
					DatabaseCheck=ReadBool(DatabaseCheckNode.get("enabled") if DatabaseCheckNode is not None else None, True),
					XmlNode=TypeNode,
				)
			)

	return Rules


def ParseIgnoreRule(Node: XmlTree.Element) -> PreprocessRegexRule:
	DescriptionNode = Node.find("description")
	Description = Node.get("description", "")
	if DescriptionNode is not None and DescriptionNode.text is not None:
		Description = DescriptionNode.text
	return PreprocessRegexRule(
		Id=Node.get("id", "unnamed_ignore_pattern"),
		Pattern=(Node.get("regex") or Node.text or "").strip(),
		Replacement="",
		Description=Description,
		Enabled=ReadBool(Node.get("enabled"), True),
	)


def ParseRegexRule(Node: XmlTree.Element) -> PreprocessRegexRule:
	PatternNode = Node.find("pattern")
	ReplacementNode = Node.find("replacement")
	DescriptionNode = Node.find("description")
	return PreprocessRegexRule(
		Id=Node.get("id", "unnamed_rule"),
		Pattern=(PatternNode.text if PatternNode is not None and PatternNode.text is not None else ""),
		Replacement=(ReplacementNode.text if ReplacementNode is not None and ReplacementNode.text is not None else ""),
		Description=(DescriptionNode.text if DescriptionNode is not None and DescriptionNode.text is not None else ""),
		Enabled=ReadBool(Node.get("enabled"), True),
	)


def SavePreprocessRules(Rules: PreprocessRules) -> None:
	IndentXml(Rules.Root)
	Rules.Tree.write(Rules.Path, encoding="utf-8", xml_declaration=True)


def CreateDefaultPreprocessRules(PathToRules: Path) -> None:
	PathToRules.parent.mkdir(parents=True, exist_ok=True)
	Root = XmlTree.Element("xlsx_preprocess_rules", {"version": RULES_VERSION})
	Settings = XmlTree.SubElement(Root, "settings")
	XmlTree.SubElement(Settings, "regex_flags", {"ignore_case": "true", "unicode": "true"})
	XmlTree.SubElement(Settings, "whitespace", {"collapse": "true", "trim": "true"})
	XmlTree.SubElement(Settings, "cache", {"auto_update": "true"})

	IgnoredTexts = XmlTree.SubElement(Root, "ignored_texts")
	AddIgnorePattern(IgnoredTexts, "board_contains_placed", "Board/block header listing placed components", r"\b(плата|блок|модуль|вузол)\b.*\b(на\s+н(і|и)й|де|в\s+якому|у\s+якому)\b.*\b(розміщен[іи]|встановлен[іи]|змонтован[іи]|наявн[іи]|присутн[іи]|розташован[іи])\b")
	AddIgnorePattern(IgnoredTexts, "board_contains_short", "Short board header: на ній:", r"\b(плата|блок|модуль|вузол)\b.*\bна\s+н(і|и)й\b\s*:?$")
	AddIgnorePattern(IgnoredTexts, "board_missing_elements", "Board/block header listing missing components", r"\b(плата|блок|модуль|вузол)\b.*\b(відсутн[іи]|немає|відсутнє|без)\b")
	AddIgnorePattern(IgnoredTexts, "board_except_elements", "Board/block header listing present except missing components", r"\b(плата|блок|модуль|вузол)\b.*\b(в\s+наявності|наявн[іи]|присутн[іи])\b.*\b(за\s+виключенням|за\s+винятком|крім|окрім)\b")
	AddIgnorePattern(IgnoredTexts, "generic_placed_marker", "Generic phrase marker for placed/listed contents rows", r"\b(на\s+н(і|и)й|на\s+якій|де)\b\s*:?\s*$|\b(на\s+н(і|и)й|на\s+якій|де)\b.*\b(розміщен[іи]|розташован[іи]|встановлен[іи]|змонтован[іи])\b")
	AddIgnorePattern(IgnoredTexts, "generic_missing_marker", "Generic phrase marker for missing/excluded contents rows", r"\b(відсутн[іи]|немає)\b\s*:?$|\b(в\s+наявності|наявн[іи])\b.*\b(за\s+виключенням|за\s+винятком|крім|окрім)\b")
	AddIgnorePattern(IgnoredTexts, "generic_military_rank", "Rows containing Ukrainian military rank words", r"\b(рекрут\w*|солдат\w*|сержант\w*|старшин\w*|прапорщик\w*|лейтенант\w*|капітан\w*|майор\w*|підполковник\w*|полковник\w*|генерал\w*)\b")
	AddIgnorePattern(IgnoredTexts, "generic_commission", "Rows containing commission wording", r"\bкомісі\w*\b")
	AddIgnorePattern(IgnoredTexts, "generic_completeness", "Rows containing completeness/complectation wording", r"\b(комплектн\w*|комплектаці\w*)\b")
	AddIgnorePattern(IgnoredTexts, "generic_placed_word", "Rows containing placed/located wording", r"\bрозміщен\w*\b")
	AddIgnorePattern(IgnoredTexts, "generic_available_word", "Rows containing available/present wording", r"\bнаявн\w*\b")
	AddIgnorePattern(IgnoredTexts, "generic_property_word", "Rows containing property wording", r"\bмайн\w*\b")
	AddIgnorePattern(IgnoredTexts, "generic_counting_word", "Rows containing count/calculation wording", r"\bпідрахун\w*\b")

	StageOne = XmlTree.SubElement(Root, "stage", {"id": "1", "name": "language_normalization"})
	AddRule(StageOne, "ru_relay_res", "Russian РЭС relay spelling to Ukrainian РЕС", r"\bРЭС\b", "РЕС")
	AddRule(StageOne, "ru_diode_to_uk_diode", "Russian диод to Ukrainian діод", r"\bдиод\b", "діод")

	Types = XmlTree.SubElement(Root, "element_types")
	AddElementType(Types, "capacitor", "Конденсатор", ["конденсатор", "конд", "кондер"], ["КМ", "К10", "К50"])
	AddElementType(Types, "resistor", "Резистор", ["резистор", "сопротивление", "опір"], ["МЛТ", "СП", "С2"])
	AddElementType(Types, "diode", "Діод", ["діод", "диод"], ["Д", "КД"])
	AddElementType(Types, "relay", "Реле", ["реле"], ["РЕС", "РП"])

	StageThree = XmlTree.SubElement(Root, "stage", {"id": "3", "name": "technical_normalization"})
	AddRule(StageThree, "km_series_hyphenation", "Normalize КМ capacitor names like КМ 5б Н90 to КМ-5б-Н90", r"\bКМ[\s\-]*(\d+[а-яА-Яa-zA-Z]?)[\s\-]*(Н\d+)\b", "КМ-$1-$2")
	AddRule(StageThree, "res_relay_series_hyphenation", "Normalize РЕС relay names with missing hyphen", r"\bРЕС[\s\-]*(\d+)\b", "РЕС-$1")

	Tree = XmlTree.ElementTree(Root)
	IndentXml(Root)
	Tree.write(PathToRules, encoding="utf-8", xml_declaration=True)


def AddIgnorePattern(Parent: XmlTree.Element, PatternId: str, Description: str, Pattern: str) -> None:
	PatternNode = XmlTree.SubElement(Parent, "pattern", {"id": PatternId, "enabled": "true", "description": Description})
	PatternNode.text = Pattern


def AddRule(Parent: XmlTree.Element, RuleId: str, Description: str, Pattern: str, Replacement: str) -> None:
	Rule = XmlTree.SubElement(Parent, "rule", {"id": RuleId, "enabled": "true"})
	XmlTree.SubElement(Rule, "description").text = Description
	XmlTree.SubElement(Rule, "pattern").text = Pattern
	XmlTree.SubElement(Rule, "replacement").text = Replacement


def AddElementType(Parent: XmlTree.Element, TypeId: str, Canonical: str, Aliases: Sequence[str], Prefixes: Sequence[str]) -> None:
	TypeNode = XmlTree.SubElement(Parent, "type", {"id": TypeId, "canonical": Canonical})
	AliasesNode = XmlTree.SubElement(TypeNode, "aliases")
	for Alias in Aliases:
		XmlTree.SubElement(AliasesNode, "alias").text = Alias
	CacheNode = XmlTree.SubElement(TypeNode, "prefix_cache")
	for Prefix in Prefixes:
		XmlTree.SubElement(CacheNode, "prefix").text = Prefix
	XmlTree.SubElement(TypeNode, "database_check", {"enabled": "true"})


def ReadBool(Value: Optional[str], Default: bool) -> bool:
	if Value is None:
		return Default
	return Value.strip().casefold() in ("1", "true", "yes", "on")


def ConvertReplacement(Replacement: str) -> str:
	return re.sub(r"\$(\d+)", r"\\\1", Replacement)


def ExtractLeadingPrefix(Text: str) -> str:
	Clean = Text.strip()
	Match = re.match(r"^([A-Za-zА-Яа-яІіЇїЄєҐґ]+\d*)", Clean)
	return Match.group(1).upper() if Match else ""


def PrefixMatches(Token: str, Prefix: str) -> bool:
	TokenFolded = Token.casefold()
	PrefixFolded = Prefix.casefold()
	if TokenFolded == PrefixFolded:
		return True
	if not TokenFolded.startswith(PrefixFolded):
		return False
	Remainder = Token[len(Prefix):]
	return bool(Remainder) and not Remainder[0].isalpha()


def IndentXml(Element: XmlTree.Element, Level: int = 0) -> None:
	Indent = "\n" + Level * "\t"
	if len(Element):
		if not Element.text or not Element.text.strip():
			Element.text = Indent + "\t"
		for Child in Element:
			IndentXml(Child, Level + 1)
		if not Child.tail or not Child.tail.strip():
			Child.tail = Indent
	if Level and (not Element.tail or not Element.tail.strip()):
		Element.tail = Indent


def NormalizeDesignationCase(Text: str) -> str:
	def UpperLettersBeforeDigits(Match: re.Match[str]) -> str:
		return Match.group(1).upper() + Match.group(2)

	return re.sub(r"(?<!\d)([A-Za-zА-Яа-яІіЇїЄєҐґ]+)(\d+)", UpperLettersBeforeDigits, Text)
