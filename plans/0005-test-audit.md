# 0005 — Test Audit: Spec 0005/0006 vs Test Coverage

Systematic walk of every spec acceptance criterion against the four test projects.

---

## Legend

- ✅ = Covered
- ❌ = **Gap — no test**
- ⚠️ = Partial coverage
- 🔮 = Planned feature (not yet implemented) — will need tests

Test level recommendations: **R** = Runtime.Tests, **I** = IntegrationTests, **A** = AddinTests

---

## 1. Entry Points (spec 0005 §Entry Points)

| Criterion | Status | Where Tested | Gap / Notes |
|---|---|---|---|
| Quote-prefix backtick formula intercepts and rewrites | ✅ | AddinTests: ScalarExpressionReturnsCorrectValue, ValuePath_* | |
| Multiple backtick expressions in one formula | ❌ | — | **A**: Two backtick exprs in one cell — verify both compile and combine |
| Floating editor opens and applies | ❌ | — | Manual/UI test only — acceptable gap |
| LET formula with backtick bindings | ✅ | AddinTests: LetFormula_TwoRange, LetFormula_VariableReusedInLambda | |
| LET formula with backtick in result position | ⚠️ | Unit: LetFormulaRewriterTests, ReconstructorTests | **A/I**: No integration or addin test for backtick as the LET result expression |

## 2. Parameter Detection (spec 0005 §Expression Language)

| Criterion | Status | Where Tested | Gap / Notes |
|---|---|---|---|
| Free variables become UDF params | ✅ | Unit: InputDetectorTests (7 tests), Integration: FreeVariable_BecomesParameter | |
| Lambda params excluded | ✅ | Unit: Detect_LambdaParams_NotParameters | |
| Method/property names excluded | ✅ | Implicit in InputDetector tests | |
| Range references (A1:B10) become params | ✅ | Unit: Detect_RangeRef_*, Integration: multiple ValuePath tests | |
| Single-cell references as free variables | ⚠️ | Integration: FreeVariable tests use scalars | **I**: No explicit test for B1 as a free variable resolving as a cell ref |
| Statement blocks — var declarations excluded | ❌ | — | **R/I**: Statement block with var — verify var name is not a UDF param |
| No "primary input" concept — all params equal | ✅ | Unit: Emit_UniformPreamble_AllParamsGetWrapping | |

## 3. Type System — ExcelValue.Wrap() (spec 0005 §Type System)

| Criterion | Status | Where Tested | Gap / Notes |
|---|---|---|---|
| Scalar → ExcelScalar | ✅ | Runtime: ExcelValueTests (double, string, bool, null) | |
| object[,] → ExcelArray | ✅ | Runtime: Wrap_Array, Wrap_MultiRowArray | |
| object[,] with headers → ExcelTable | ✅ | Runtime: Wrap_ArrayWithHeaders | |
| 1x1 array → ExcelScalar (unwrap) | ✅ | Runtime: Wrap_SingleCellArray | |
| 1x1 array WITH headers → ExcelTable (no unwrap) | ✅ | Runtime: Wrap_SingleCellArrayWithHeaders | |
| Already-wrapped ExcelValue → same instance | ✅ | Runtime: Wrap_ExcelValue_ReturnsSameInstance | |

## 4. ExcelScalar Operations (spec 0005 §ExcelScalar)

| Criterion | Status | Where Tested | Gap / Notes |
|---|---|---|---|
| .Sum() returns value | ✅ | Runtime: ExcelScalarTests | |
| .Count() returns 1 | ✅ | Runtime: ExcelScalarTests | |
| Comparison operators | ✅ | Runtime: ExcelScalarTests, ExcelValueTests | |
| Arithmetic operators | ❌ | — | **R**: myCell * 2 + 1 — ExcelScalar arithmetic operators not tested |
| .Where, .Any, .All, .First, .FirstOrDefault | ✅ | Runtime: ExcelScalarTests | |
| .Take, .Skip | ✅ | Runtime: ExcelScalarTests | |
| .Average | ✅ | Runtime: ExcelScalarTests | |
| .Aggregate | ✅ | Runtime: ExcelScalarTests | |
| .Scan | ❌ | — | **R**: ExcelScalar.Scan() not tested |
| .SelectMany | ❌ | — | **R**: ExcelScalar.SelectMany() not tested |
| .Map | ❌ | — | **R**: ExcelScalar.Map() not tested |
| .OrderBy / .OrderByDescending | ❌ | — | **R**: Single-element sort (trivial but confirms API exists) |
| .Distinct | ❌ | — | **R**: Single-element distinct |
| Implicit conversion to double/string | ✅ | Runtime: ImplicitConversions_Work | |
| .Min / .Max | ❌ | — | **R**: ExcelScalar.Min() and .Max() not tested |

## 5. ExcelArray Element-Wise Operations (spec 0005 §ExcelArray)

| Criterion | Status | Where Tested | Gap / Notes |
|---|---|---|---|
| .Where (element-wise) | ✅ | Runtime: ExcelArrayTests | |
| .Select (element-wise, flattens to 1 col) | ✅ | Runtime: Select_TransformsCellsElementWise | |
| .SelectMany | ✅ | Runtime: SelectMany_FlattensResults | |
| .Map (preserves 2D shape) | ❌ | — | **R**: .Map(x => x * 2) preserving original dimensions not tested |
| .Any / .All | ✅ | Runtime: ExcelArrayTests | |
| .First / .FirstOrDefault | ✅ | Runtime: ExcelArrayTests | |
| .OrderBy / .OrderByDescending | ✅ | Runtime: ExcelArrayTests | |
| .Take (positive and negative) | ✅ | Runtime: ExcelArrayTests | |
| .Skip (positive and negative) | ✅ | Runtime: ExcelArrayTests | |
| .Distinct | ✅ | Runtime: ExcelArrayTests | |
| .Aggregate | ✅ | Runtime: ExcelArrayTests | |
| .Scan | ✅ | Runtime: ExcelArrayTests | |
| .Count | ✅ | Runtime: ExcelArrayTests | |
| .Sum / .Min / .Max / .Average | ✅ | Runtime: ExcelArrayTests | |
| Empty array edge cases | ✅ | Runtime: EmptyArray_Count/Any/All | |

## 6. ExcelTable & Row-Wise Operations (spec 0005 §ExcelTable, §Row-Wise)

| Criterion | Status | Where Tested | Gap / Notes |
|---|---|---|---|
| .Headers returns column names | ✅ | Runtime: ExcelTableTests | |
| .Rows.Where | ✅ | Runtime + Integration + AddinTests | |
| .Rows.Select | ✅ | Runtime: Rows_Select_MapsRowsToValues | |
| .Rows.Any | ✅ | Runtime: ExcelArrayTests Rows_Any | |
| .Rows.All | ❌ | — | **R**: RowCollection.All() not tested |
| .Rows.First | ✅ | Runtime: Rows_First_ReturnsFirstMatchingRow | |
| .Rows.FirstOrDefault | ❌ | — | **R**: RowCollection.FirstOrDefault() not tested |
| .Rows.OrderBy | ✅ | Runtime: Rows_OrderBy_SortsByColumn | |
| .Rows.OrderByDescending | ❌ | — | **R**: RowCollection.OrderByDescending() not tested |
| .Rows.Take (positive and negative) | ❌ | — | **R**: RowCollection.Take() not tested (element-wise Take tested, not row-wise) |
| .Rows.Skip (positive and negative) | ❌ | — | **R**: RowCollection.Skip() not tested |
| .Rows.Distinct | ❌ | — | **R**: RowCollection.Distinct() not tested |
| .Rows.Count | ✅ | Integration: Sugar_RowCount | |
| .Rows.ToRange | ✅ | Runtime: Rows_ToRange_ConvertsBackToExcelArray | |
| 🔮 .Rows.Aggregate | — | — | Planned (#99) |
| 🔮 .Rows.Scan | — | — | Planned (#99) |
| 🔮 .Rows.GroupBy | — | — | Planned (#101) |

## 7. Column Access on Rows (spec 0005 §Column Access)

| Criterion | Status | Where Tested | Gap / Notes |
|---|---|---|---|
| r["ColumnName"] string bracket access | ✅ | Runtime + Integration + Addin | |
| r[0] / r[-1] numeric index | ✅ | Runtime: RowTests, Integration: Sugar_NegativeIndex | |
| r.ColumnName dot notation rewrite | ✅ | Unit: DotNotationRewriteTests | |
| Dot notation conflict detection | ✅ | Unit: BuildMapping_ConflictDetection | |
| Case-insensitive column lookup | ✅ | Runtime: ColumnAccess_CaseInsensitive | |
| ColumnValue comparison operators | ✅ | Runtime: ColumnValueTests | |
| ColumnValue arithmetic operators | ✅ | Runtime: ColumnValueTests | |
| ColumnValue implicit conversions | ✅ | Runtime: ColumnValueTests | |
| ColumnValue cross-type with ExcelValue | ⚠️ | Integration: MultipleParams_FilterWithThreshold_NoCast | **R**: No direct unit test for ColumnValue > ExcelScalar operator |
| ColumnValue.ToString() | ✅ | Runtime: ColumnValueTests | |

## 8. Cell Formatting Access (spec 0005 §Cell Formatting)

| Criterion | Status | Where Tested | Gap / Notes |
|---|---|---|---|
| .Cells iteration on ExcelArray | ✅ | Runtime: CellEscalationTests | |
| .Cells on ExcelScalar | ✅ | Runtime: CellEscalationTests | |
| .Cell escalation from ColumnValue | ✅ | Runtime: CellEscalationTests | |
| Cell properties (Value, Color, Bold, etc.) | ✅ | Runtime: CellTests | |
| Cell sub-objects (Interior, CellFont) | ✅ | Runtime: CellTests | |
| Without origin → throws | ✅ | Runtime: CellEscalationTests | |
| Without bridge → throws | ✅ | Runtime: CellEscalationTests | |
| IsMacroType auto-detection (.Cell/.Cells) | ✅ | Unit: InputDetectorTests, Integration, AddinTests | |
| End-to-end color filtering in Excel | ✅ | AddinTests: ObjectModelPath_WhereColor_Sum/Count | |
| End-to-end r["Col"].Cell.Bold | ❌ | — | **A**: Filtering rows by r["Status"].Cell.Bold — tests cell escalation through table rows in live Excel |
| Cell.Formula property | ❌ | — | **R**: No test verifying Cell.Formula is populated |
| Cell.Format property | ❌ | — | **R**: No test verifying Cell.Format is populated |
| Cell.Address property | ❌ | — | **R**: No test verifying Cell.Address is populated |
| Cell.Rgb property (distinct from ColorIndex) | ❌ | — | **A**: data.Cells.Where(c => c.Rgb == 255).Sum() — Rgb path untested end-to-end |
| Cell aggregation extensions (.Sum, .Average, .Min, .Max) | ✅ | Runtime: CellExtensionsTests | |
| Cell aggregation on empty collection | ✅ | Runtime: CellExtensionsTests (Sum_Empty, Average_Empty) | |

## 9. Result Handling (spec 0005 §Result Handling)

| Criterion | Status | Where Tested | Gap / Notes |
|---|---|---|---|
| Single value → bare scalar in cell | ✅ | AddinTests: ScalarResult_ReturnsBareValue | |
| Multi-cell → spills as dynamic array | ✅ | AddinTests: ValuePath_Where, ValuePath_Select_Multiply | |
| Row result → spills as single row | ❌ | — | **R/I**: .Rows.First(r => ...) returning a Row should spill as one row |
| ColumnValue result → single value | ❌ | — | **R/I**: Returning a ColumnValue directly should display as single value |
| ResultConverter: null → empty string | ✅ | Runtime: Convert_NullReturnsEmpty | |
| ResultConverter: ExcelValue delegation | ✅ | Runtime: Convert_ExcelValueDelegates | |
| ResultConverter: IExcelRange → object[,] | ✅ | Runtime: IExcelRangeToResult_ConvertsFilteredRows | |
| ResultConverter: IEnumerable\<Row\> → object[,] | ❌ | — | **R**: No test for converting a raw IEnumerable\<Row\> result |
| ResultConverter: IEnumerable\<ColumnValue\> → object[,] | ❌ | — | **R**: No test for converting IEnumerable\<ColumnValue\> |

## 10. Error Handling (spec 0005 §Error Handling)

| Criterion | Status | Where Tested | Gap / Notes |
|---|---|---|---|
| Compile error → cell comment | ✅ | AddinTests: InvalidExpression_ShowsError | |
| Runtime error → #VALUE! | ❌ | — | **A**: Deliberate runtime error (e.g. divide by zero) → verify #VALUE! |

## 11. Pipeline Architecture (spec 0006)

| Criterion | Status | Where Tested | Gap / Notes |
|---|---|---|---|
| Backtick extraction | ✅ | Unit: InterceptionTests (6 tests) | |
| Backtick formula rewriting | ✅ | Unit: InterceptionTests (4 tests) | |
| InputDetector range preprocessing | ✅ | Unit: InputDetectorTests | |
| InputDetector object model detection | ✅ | Unit: InputDetectorTests | |
| InputDetector per-variable header tracking | ✅ | Unit: InputDetectorTests (4 tests) | |
| CodeEmitter uniform preamble | ✅ | Unit: CodeEmitterTests | |
| CodeEmitter UDF naming (hash, preferred, reserved) | ✅ | Unit: CodeEmitterTests | |
| Pipeline caching | ✅ | Unit+Integration | |
| Pipeline [#All] append for table header vars | ✅ | Unit: Pipeline_AppendsAllToHeaderVariableTableParameter | |
| Pipeline does NOT append [#All] to range refs | ✅ | Unit: Pipeline_DoesNotAppendAllToRangeRefHeaderVariable | |
| LET formula parsing | ✅ | Unit: LetFormulaParserTests (extensive) | |
| LET formula rewriting | ✅ | Unit: LetFormulaRewriterTests | |
| LET formula reconstruction (edit round-trip) | ✅ | Unit: LetFormulaReconstructorTests | |
| ALC assembly loading | ✅ | Runtime: AssemblyIdentitySpikeTests | |
| RuntimeHelpers has no ExcelDNA dependency | ✅ | Unit: RuntimeHelpersGuardTests | |
| Runtime has no ExcelDNA dependency | ✅ | Runtime: RuntimeAssembly_HasNoExcelDnaDependency | |

## 12. Intellisense (spec 0005 §Intellisense, spec 0006 §Intellisense)

| Criterion | Status | Where Tested | Gap / Notes |
|---|---|---|---|
| Wrapper type method completions | ✅ | Unit: RoslynCompletionTests | |
| Column name completions (dot and bracket) | ✅ | Unit: RoslynCompletionTests, CompletionScopingTests | |
| Standard C# method completions (string) | ✅ | Unit: StringMethods_Available | |
| Internal type filtering | ✅ | Unit: FiltersInternalTypes | |
| Object method filtering | ✅ | Unit: FiltersObjectMethods | |
| Completion scoping through chains | ✅ | Unit: CompletionScopingTests | |
| Column completion insertion (bracket + dot) | ✅ | Unit: ColumnCompletionInsertionTests | |
| Synthetic document with LET bindings | ✅ | Unit: SyntheticDocumentBuilderTests | |

## 13. Architecture Review Findings (plans/0003)

| Finding | Status | Where Tested | Gap / Notes |
|---|---|---|---|
| F1: Duplicate ToResultDelegate | ⚠️ | — | Design issue, not a test gap per se. No test verifying AddIn and test helper produce identical results |
| F2: Dot notation in LINQ lambdas (RowCollection) | ✅ | Runtime: Rows_HaveDynamicColumnAccess, Integration: MultipleParams_NoCast | Works via RowCollection instance methods |
| F3: ToResult wraps scalars in 1x1 | ✅ | Runtime: ScalarToResult_ReturnsBareValue (fixed — returns bare value) | |
| F5: GetHeadersDelegate contract mismatch | ❌ | — | **I**: No test that header extraction works for both ExcelReference and object[,] inputs |
| F11: RuntimeBridge.GetCell never initialised | ⚠️ | Runtime: CellEscalationTests use mock | **A**: No addin test for r["Col"].Cell.Bold to verify bridge is initialised in production |

---

## Summary: Gaps by Priority

### High Priority — custom code, untested methods

1. **RowCollection missing method tests** — .All(), .FirstOrDefault(), .OrderByDescending(), .Take(), .Skip(), .Distinct() — **R** (6 tests)
2. **ExcelArray.Map()** — preserves 2D shape, distinct from Select — **R** (1 test)
3. **ExcelScalar missing methods** — .Scan(), .SelectMany(), .Map(), .OrderBy/Descending(), .Distinct(), .Min(), .Max() — **R** (7 tests)
4. **ResultConverter for Row and ColumnValue** — IEnumerable\<Row\> and IEnumerable\<ColumnValue\> conversion — **R** (2 tests)
5. **ColumnValue cross-type operators** — ColumnValue > ExcelScalar, ColumnValue == ExcelValue — **R** (1 test)
6. **Row result spilling** — .Rows.First() result renders as single row — **R/I** (1 test)

### Medium Priority — integration/addin paths

7. **Cell escalation through table rows (addin)** — r["Status"].Cell.Bold end-to-end — **A** (1 test)
8. **Cell.Rgb end-to-end** — filtering by RGB color value — **A** (1 test)
9. **Backtick in LET result position** — end-to-end — **I/A** (1 test)
10. **Multiple backtick expressions in one non-LET formula** — **A** (1 test)
11. **Runtime error → #VALUE!** — deliberate runtime error surfaces correctly — **A** (1 test)
12. **Statement block expression** — { var x = 1; return tbl.Sum() + x; } compiles and executes — **I** (1 test)
13. **ExcelScalar arithmetic operators** — myCell * 2 + 1 — **R** (1 test)

### Low Priority — edge cases, properties

14. **Cell.Formula, Cell.Format, Cell.Address** — verify properties populated from COM — **R** (mock) or **A** (live) (3 tests)
15. **Single-cell ref as free variable** — B1 resolves as cell reference — **I** (1 test)
16. **GetHeadersDelegate contract** — works for both ExcelReference and object[,] inputs — **I** (1 test)

### Planned Features — need tests when implemented

17. 🔮 .Rows.Aggregate() — #99
18. 🔮 .Rows.Scan() — #99
19. 🔮 .Rows.GroupBy() — #101

---

**Total gaps: ~30 tests across ~16 distinct scenarios.** The Runtime.Tests layer has the most gaps (RowCollection methods, ExcelScalar methods, Map, ResultConverter edge cases). Integration and Addin layers are reasonably well covered for the happy paths but missing a few important integration scenarios (cell escalation through tables, statement blocks, multiple backticks).
