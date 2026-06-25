# XLSX Preprocessing Module Design

## Goal

Add a GUI-opened preprocessing module that normalizes element names inside `.xlsx` inventory files before the existing DGM database lookup/fill workflow.

The module should transform inconsistent human-written names such as:

- `Конденсатор`
- `конденсатор`
- `КМ-5б-Н90`
- `КМ 5б Н90`
- `Реле РЭС`

into standardized Ukrainian, database-friendly names such as:

- `Конденсатор КМ-5б-Н90`
- `Реле РЕС`

The normalization logic should be driven by a separate XML rules file, so future rule expansion does not require code edits.

## Recommended Module Shape

### Python module

Suggested name:

```text
dgm_xlsx_preprocessor.py
```

Primary responsibilities:

1. Load preprocessing XML rules.
2. Normalize one element name string.
3. Optionally preprocess an entire `.xlsx` file.
4. Track changes for review before saving.
5. Update cache entries in the XML when database-confirmed type detection succeeds.

### GUI integration

The main window should later receive a new button:

```text
Preprocess .xlsx file
```

Optional later additions:

```text
Preprocess folder
Preprocess then fill .xlsx file
```

The preprocessor should be independent from the existing fill logic but may reuse:

- `self.Database`
- `self.Database.Columns.Name`
- `self.Database.FindElement(...)`

## Processing Pipeline

Each raw cell value from the configured name column should pass through these stages:

```text
raw cell text
  ↓
basic whitespace cleanup
  ↓
stage 1: language normalization
  ↓
stage 2: element type placement/detection
  ↓
stage 3: technical spelling normalization
  ↓
final normalized text
```

The module should preserve non-string or empty cells unchanged.

## XML Rules File

Suggested file name:

```text
preprocess_rules.xml
```

Suggested root:

```xml
<xlsx_preprocess_rules version="1">
    ...
</xlsx_preprocess_rules>
```

The file should be separate from the main DGM database XML, because it contains transformation logic rather than DGM metal values.

## Full XML Skeleton

```xml
<?xml version="1.0" encoding="utf-8"?>
<xlsx_preprocess_rules version="1">

    <settings>
        <regex_flags ignore_case="true" unicode="true"/>
        <whitespace collapse="true" trim="true"/>
        <cache auto_update="true"/>
    </settings>

    <stage id="1" name="language_normalization">
        <rule id="ru_relay_res" enabled="true">
            <description>Russian РЭС relay spelling to Ukrainian РЕС</description>
            <pattern>\bРЭС\b</pattern>
            <replacement>РЕС</replacement>
        </rule>
    </stage>

    <element_types>
        <type id="capacitor" canonical="Конденсатор">
            <aliases>
                <alias>конденсатор</alias>
                <alias>Кондер</alias>
                <alias>конд</alias>
            </aliases>

            <prefix_cache>
                <prefix>КМ</prefix>
                <prefix>К10</prefix>
                <prefix>К50</prefix>
            </prefix_cache>

            <database_check enabled="true"/>
        </type>

        <type id="resistor" canonical="Резистор">
            <aliases>
                <alias>резистор</alias>
                <alias>сопротивление</alias>
                <alias>опір</alias>
            </aliases>

            <prefix_cache>
                <prefix>МЛТ</prefix>
                <prefix>СП</prefix>
                <prefix>С2</prefix>
            </prefix_cache>

            <database_check enabled="true"/>
        </type>

        <type id="diode" canonical="Діод">
            <aliases>
                <alias>діод</alias>
                <alias>диод</alias>
            </aliases>

            <prefix_cache>
                <prefix>Д</prefix>
                <prefix>КД</prefix>
                <prefix>ДГ</prefix>
            </prefix_cache>

            <database_check enabled="true"/>
        </type>
    </element_types>

    <stage id="3" name="technical_normalization">
        <rule id="km_series_hyphenation" enabled="true">
            <description>Normalize КМ ceramic capacitor spelling</description>
            <pattern>\bКМ[\s\-]*(\d+[а-яА-Яa-zA-Z]?)[\s\-]*(Н\d+)\b</pattern>
            <replacement>КМ-$1-$2</replacement>
        </rule>
    </stage>

</xlsx_preprocess_rules>
```

## XML Sections

### Settings

```xml
<settings>
    <regex_flags ignore_case="true" unicode="true"/>
    <whitespace collapse="true" trim="true"/>
    <cache auto_update="true"/>
</settings>
```

#### `regex_flags`

Controls default regex behavior.

- `ignore_case="true"` means regexes should normally use `re.IGNORECASE`.
- `unicode="true"` is mostly documentary in Python 3, but should be kept for clarity.

#### `whitespace`

- `trim="true"` removes leading/trailing whitespace.
- `collapse="true"` converts repeated whitespace to one space.

Example:

```text
"  КМ   5б   Н90  " → "КМ 5б Н90"
```

#### `cache`

If `auto_update="true"`, the program may update `<prefix_cache>` entries when stage 2 discovers a new reliable prefix through database lookup.

### Stage 1: Language normalization

Purpose: convert Russian or mixed Russian/Ukrainian words and markings into Ukrainian-standard spellings before element type detection.

Example:

```xml
<stage id="1" name="language_normalization">
    <rule id="ru_relay_res" enabled="true">
        <description>Russian РЭС relay spelling to Ukrainian РЕС</description>
        <pattern>\bРЭС\b</pattern>
        <replacement>РЕС</replacement>
    </rule>
</stage>
```

Each rule should have:

```xml
<rule id="unique_rule_id" enabled="true">
    <description>Human-readable explanation</description>
    <pattern>regex pattern</pattern>
    <replacement>replacement text</replacement>
</rule>
```

Rules:

- Rules are applied in XML order.
- Disabled rules are skipped.
- Regex replacement should use Python `re.sub`.
- Backreferences should support `$1`, `$2`, and `$3` by converting them to Python-compatible syntax internally.

### Element types

Purpose: define canonical element type names, known aliases, and cache prefixes used in stage 2.

Example:

```xml
<element_types>
    <type id="capacitor" canonical="Конденсатор">
        <aliases>
            <alias>конденсатор</alias>
            <alias>Кондер</alias>
            <alias>конд</alias>
        </aliases>

        <prefix_cache>
            <prefix>КМ</prefix>
            <prefix>К10</prefix>
            <prefix>К50</prefix>
        </prefix_cache>

        <database_check enabled="true"/>
    </type>
</element_types>
```

The `id` attribute is a stable internal identifier. The `canonical` attribute is the exact standardized first word to place at the beginning of the element name.

Aliases represent known spellings of the type word and should be treated case-insensitively.

Examples:

```text
"конденсатор КМ 5б Н90" → "Конденсатор КМ 5б Н90"
"КМ 5б Н90 конденсатор" → "Конденсатор КМ 5б Н90"
```

The prefix cache helps infer missing element types. For example, if `КМ` is cached under `Конденсатор`, stage 2 can test `Конденсатор КМ-5б-Н90` against the database.

## Stage 2: Element Type Placement and Detection

Stage 2 should enforce this structure:

```text
CanonicalType rest-of-name
```

Examples:

```text
"конденсатор КМ 5б Н90" → "Конденсатор КМ 5б Н90"
"КМ 5б Н90 конденсатор" → "Конденсатор КМ 5б Н90"
"КМ 5б Н90" → "Конденсатор КМ 5б Н90", if database confirms
```

### Step 1: Find explicit type alias

For every `<type>`:

- Check whether the string starts with the canonical name or one of its aliases.
- Check whether the string ends with the canonical name or one of its aliases.
- Optionally check whether the alias appears as a standalone word inside the string.

If found:

1. Remove the alias from its old position.
2. Trim/collapse whitespace.
3. Prepend the canonical type.
4. Capitalize exactly as `canonical`.

### Step 2: Use prefix cache when type is missing

If no explicit type alias is found:

1. Extract the first meaningful token from the string.
2. Compare it against all `<prefix>` values.
3. If a prefix matches, build a candidate: `{canonical} {current_text}`.
4. Verify candidate with `Database.FindElement(candidate)`.
5. If found, accept candidate.

### Step 3: Extensive database search

If cache lookup fails:

1. Try every known type as `{canonical} {current_text}`.
2. Query the database for each candidate.
3. If exactly one candidate matches, accept it.
4. If multiple candidates match, do not modify automatically; report ambiguity.
5. If no candidate matches, leave text unchanged.

### Step 4: Cache update

If extensive database search finds a type, update that type's `<prefix_cache>`.

Rules:

- Only update cache if `auto_update="true"`.
- Do not add duplicate prefixes.
- Prefix extraction should be deterministic.
- Prefer a conservative prefix:
  - leading Cyrillic/Latin letters;
  - optionally followed by digits if that is part of the series name;
  - stop before separator, whitespace, or subtype details.

Examples:

```text
"КМ-5б-Н90" → "КМ"
"К10-17" → "К10"
"МЛТ-0,25" → "МЛТ"
"СП5-2" → "СП5"
```

## Stage 3: Technical Normalization

Purpose: normalize formatting inside the element designation.

Example:

```xml
<stage id="3" name="technical_normalization">
    <rule id="km_series_hyphenation" enabled="true">
        <description>Normalize КМ ceramic capacitor spelling</description>
        <pattern>\bКМ[\s\-]*(\d+[а-яА-Яa-zA-Z]?)[\s\-]*(Н\d+)\b</pattern>
        <replacement>КМ-$1-$2</replacement>
    </rule>
</stage>
```

Example transformation:

```text
"Конденсатор КМ 5б Н90" → "Конденсатор КМ-5б-Н90"
```

Stage 3 should eventually handle:

- missing hyphens;
- excessive spaces;
- inconsistent letter casing;
- decimal comma/dot normalization where safe;
- package/designation-specific spellings;
- known Soviet/CIS component series quirks.

## Rule Ordering

Rules must be applied in this order:

1. Built-in whitespace cleanup.
2. XML stage 1.
3. Stage 2 type normalization/detection.
4. XML stage 3.
5. Final whitespace cleanup.
6. Optional database verification/reporting.

Within each XML stage, rules are applied top to bottom.

## Recommended Initial Canonical Type Names

```xml
<type id="capacitor" canonical="Конденсатор"/>
<type id="resistor" canonical="Резистор"/>
<type id="diode" canonical="Діод"/>
<type id="transistor" canonical="Транзистор"/>
<type id="relay" canonical="Реле"/>
<type id="microchip" canonical="Мікросхема"/>
<type id="connector" canonical="Роз'єм"/>
<type id="switch" canonical="Перемикач"/>
<type id="transformer" canonical="Трансформатор"/>
<type id="inductor" canonical="Котушка"/>
```

These can be adjusted later to match the actual database naming style.

## XML Authoring Guidelines

When another model or a human fills the XML, use these principles:

- Prefer precise word-boundary rules.
- Put specific rules before broad rules.
- Avoid destructive normalization.
- Do not remove meaningful hyphens, digits, or designation symbols unless very certain.
- Use canonical Ukrainian type names.
- Include aliases for lowercase, Russian, Ukrainian, abbreviated, and common slang type names.
- Include `prefix_cache` entries for common component series.

Good language rule:

```xml
<pattern>\bРЭС\b</pattern>
<replacement>РЕС</replacement>
```

Risky broad rule:

```xml
<pattern>Э</pattern>
<replacement>Е</replacement>
```

Broad character replacement may corrupt technical markings.

## Suggested Result Objects

The code should eventually return structured information per changed row.

```python
@dataclass
class PreprocessChange:
    FilePath: Path
    SheetName: str
    RulesPath: Path
    Row: int
    OriginalText: str
    NewText: str
    StageNotes: list[str]
    DatabaseVerified: bool
    Ambiguous: bool = False
    StageId: str = "all"
    StageName: str = "All stages"
    ElementType: str = ""
    SafetyLevel: str = ""
```

For full-file preprocessing, return a flat `list[PreprocessChange]`. Each change carries its own file, sheet, row, rules, and stage metadata, so the GUI can group changes across multiple workbooks without keeping a parallel per-file result container.

This supports a GUI review window before saving.

## Save Behavior Recommendation

Do not silently overwrite the workbook without review.

Recommended workflow:

1. User selects `.xlsx`.
2. Preprocessor scans rows.
3. GUI shows:
   - row number;
   - original text;
   - proposed text;
   - reason/stage notes;
   - database verification status.
4. User can:
   - apply all;
   - apply selected;
   - skip selected;
   - cancel.
5. Save changes to the workbook only after confirmation.

Optional safety feature:

```text
filename.xlsx → filename.preprocessed.xlsx
```

or automatic backup:

```text
filename.xlsx.bak
```

## Implementation Notes for Later

### Use current database column settings

The existing project stores the name column here:

```python
self.Database.Columns.Name
```

So preprocessing should read/write cells from that column.

### Reuse database lookup

Use:

```python
self.Database.FindElement(candidate)
```

A match exists when:

```python
SearchResult.Record is not None
```

### Keep normalization separate from DGM values

The preprocessor should only change the element-name column unless the user explicitly chooses a combined “preprocess then fill” workflow later.

### XML parser

Use Python standard library:

```python
xml.etree.ElementTree
```

This is already used in the project.

### Regex engine

Use Python:

```python
re
```

Compile rules once when loading XML.

## Prompt for Generating `preprocess_rules.xml`

Use this prompt with Web ChatGPT or another research assistant:

```text
Create a preprocess_rules.xml file for normalizing Ukrainian/Russian/Soviet electronic component names in XLSX inventories.

Use this XML schema:

[root]
<xlsx_preprocess_rules version="1">

[settings]
<settings>
  <regex_flags ignore_case="true" unicode="true"/>
  <whitespace collapse="true" trim="true"/>
  <cache auto_update="true"/>
</settings>

[stage 1]
<stage id="1" name="language_normalization">
  <rule id="..." enabled="true">
    <description>...</description>
    <pattern>Python-compatible regular expression</pattern>
    <replacement>replacement text, preferably using $1 backreferences</replacement>
  </rule>
</stage>

[element types]
<element_types>
  <type id="..." canonical="Ukrainian canonical type name">
    <aliases>
      <alias>...</alias>
    </aliases>
    <prefix_cache>
      <prefix>...</prefix>
    </prefix_cache>
    <database_check enabled="true"/>
  </type>
</element_types>

[stage 3]
<stage id="3" name="technical_normalization">
  <rule id="..." enabled="true">
    <description>...</description>
    <pattern>Python-compatible regular expression</pattern>
    <replacement>replacement text, preferably using $1 backreferences</replacement>
  </rule>
</stage>

Rules:
- Stage 1 converts Russian or mixed Russian/Ukrainian words into Ukrainian spellings.
- Element type canonical names must be Ukrainian.
- Stage 3 normalizes technical designations, hyphens, spacing, and common Soviet/CIS component series spellings.
- Prefer precise regexes over broad destructive replacements.
- Do not remove meaningful hyphens or digits.
- Include aliases for lowercase, Russian, Ukrainian, abbreviated, and common slang type names.
- Include prefix_cache entries for common component series.
- Output only valid XML.
```
