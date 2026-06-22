#!/usr/bin/env python3
from __future__ import annotations

import argparse
import shutil
import sys
import xml.etree.ElementTree as XmlTree
from pathlib import Path

import dgm_database

OLD_REGEX_TAG = "regex"
CATALOG_CHILD_TAGS = {"node", OLD_REGEX_TAG, dgm_database.REGEX_LEAF_TAG}


def NormalizePattern(Pattern: str) -> str:
	if Pattern.startswith("^"):
		Pattern = Pattern[1:]
	if Pattern.endswith("$"):
		Pattern = Pattern[:-1]
	return Pattern


def Migrate(DatabasePath: Path) -> int:
	Tree = XmlTree.parse(DatabasePath)
	Root = Tree.getroot()
	Catalog = Root.find("catalog")
	if Catalog is None:
		raise RuntimeError("Database does not contain a <catalog> node")

	Migrated = 0
	MissingPattern = 0
	DifferingText = 0
	UnexpectedChildren = 0

	for Node in Catalog.iter(OLD_REGEX_TAG):
		Migrated += 1
		OldPattern = Node.get("pattern", "")
		if not OldPattern:
			MissingPattern += 1
		OldText = Node.get("text", "")
		NewText = NormalizePattern(OldPattern)
		if OldText != NewText:
			DifferingText += 1
		if any(Child.tag in CATALOG_CHILD_TAGS for Child in list(Node)):
			UnexpectedChildren += 1
		Node.tag = dgm_database.REGEX_LEAF_TAG
		Node.set("text", NewText)
		Node.attrib.pop("pattern", None)

	BackupPath = DatabasePath.with_name(DatabasePath.name + ".bak")
	shutil.copy2(DatabasePath, BackupPath)

	try:
		XmlTree.indent(Tree, space="\t", level=0)
	except AttributeError:
		pass
	Tree.write(DatabasePath, encoding="utf-8", xml_declaration=True)

	print(f"Migrated regex nodes: {Migrated}")
	print(f"Missing/empty pattern attributes: {MissingPattern}")
	print(f"Old text differed from migrated text: {DifferingText}")
	print(f"Regex nodes with catalog children: {UnexpectedChildren}")
	print(f"Backup written: {BackupPath}")
	if UnexpectedChildren:
		return 1
	return 0


def Main() -> int:
	Parser = argparse.ArgumentParser(description="Migrate catalog <regex> nodes to <regex-leaf> nodes.")
	Parser.add_argument("database", nargs="?", default="database.xml", help="Path to database XML file. Defaults to database.xml.")
	Args = Parser.parse_args()
	DatabasePath = Path(Args.database).expanduser().resolve()
	if not DatabasePath.exists():
		print(f"Database file does not exist: {DatabasePath}", file=sys.stderr)
		return 2
	try:
		return Migrate(DatabasePath)
	except Exception as Error:
		print(str(Error), file=sys.stderr)
		return 1


if __name__ == "__main__":
	sys.exit(Main())
