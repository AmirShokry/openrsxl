# Developing openrsxl

## Layout

```
python/openrsxl/        Python package (openpyxl's object model, ported)
  _native.py            glue between the object model and the Rust engine
  worksheet/_cellstore.py  dict-compatible CellStore (Python side)
  extended/             openrsxl.extended (streaming features: _streaming.py,
                        load_workbook: _load.py)
src/                    Rust engine (pyo3 extension openrsxl._openrsxl)
  xml.rs                strict streaming XML tokenizer (expat semantics)
  sheet.rs              worksheet reader (WorkSheetParser.parse_cell & co)
  store.rs              CellStore: compact cell storage
  ftemplate.rs          formula text de-duplication
  formula.rs            port of openpyxl's formula Tokenizer / Translator
  strings.rs, text.rs   shared strings / inline strings
  dates.rs              from_excel / to_excel
  prescan.rs            extended: sheet parts outside <sheetData>, row
                        attributes, corner cell styles (cell data skipped)
  writer.rs             <sheetData> serialiser (lxml and ElementTree output)
tests/
  upstream/             openpyxl's own test-suite (ported, run from there)
  oracle/compare.py     whole-workbook differential comparison
  oracle/extended.py    openrsxl.extended streaming vs openpyxl full mode
  smoke/                dependency-free smoke test (Pyodide CI job)
  test_*.py             differential, API surface, security and fuzz tests
examples/               runnable documentation (tested by test_examples.py)
tools/                  porting, corpus, benchmarking and verification
```

## Build

```bash
python -m venv .venv && .venv/Scripts/pip install maturin -e .[test]
tools/dev_install.sh          # maturin develop --release (Windows-safe)
```

## Tests

```bash
pytest                                     # differential, API, security, fuzz tests
OPENPYXL_LXML=False pytest                 # same with the ElementTree backend
OPENRSXL_PURE_PYTHON=1 pytest tests/test_extended.py   # Python fallbacks
cd tests/upstream && pytest                # openpyxl's own suite
OPENRSXL_FUZZ_EXAMPLES=3000 pytest tests/test_fuzz.py tests/test_cellstore.py tests/test_api_fuzz.py
python tests/oracle/compare.py some.xlsx [--read-only|--data-only|--rich-text]
python tests/oracle/extended.py some.xlsx [--data-only]   # extended streaming
python tools/digest.py <lib> f.xlsx <dir> [--read-only] [--data-only]   # huge files
tools/verify_files.sh <workdir> big1.xlsx big2.xlsx                   # memory friendly
```

`OPENRSXL_PURE_PYTHON=1` disables the Rust engine entirely (pure Python
reference behaviour).

* `test_reader_edge_cases.py` - ~375 tiny fixtures, one reader edge case
  each, compared with openpyxl in 5 load modes.
* `test_writer_edge_cases.py` - every value type / feature written through
  the API: packages must be byte-identical.
* `test_security.py` - hostile input (XXE, entity expansion, deep nesting,
  integer overflows, path traversal ...) in every mode, extended included.
* `test_cellstore.py`, `test_api_fuzz.py` - random sequences of API calls on
  workbooks loaded by both libraries (state and saved bytes compared).
* `test_fuzz.py` - hypothesis-driven XML / number / formula fuzzing.
* `test_extended.py` - openrsxl.extended against openpyxl's full mode.
* `test_api_surface.py` - every public name and signature of openpyxl.

### Corpus runs

`tools/corpus.py` compares whole files in 10 modes each (full mode,
`data_only`, read-only, `rich_text`, `keep_vba`, `keep_links=False`, the
saved package byte for byte, and openrsxl.extended against full mode); every
file runs in its own process at idle priority, with a timeout and a memory
watchdog:

```bash
python tools/corpus.py <dir or file>... [--jobs 3] [--out results.jsonl]
```

The release corpus (~1 960 files, results in `docs/TESTING.md`) came from
the test-suites of
[XlsxWriter](https://github.com/jmcnamara/XlsxWriter)
(`xlsxwriter/test/comparison/xlsx_files`, files made by Excel),
[Apache POI](https://github.com/apache/poi) (`test-data/spreadsheet`,
including fuzzer crash files and XML bombs),
[ClosedXML](https://github.com/ClosedXML/ClosedXML),
[calamine](https://github.com/tafia/calamine),
[pandas](https://github.com/pandas-dev/pandas) (`pandas/tests/io/data/excel`),
[exceljs](https://github.com/exceljs/exceljs),
[python-calamine](https://github.com/dimastbk/python-calamine) and
`tests/upstream` (sparse git checkouts of the `*.xlsx` / `*.xlsm` files).

## Lint

```bash
pre-commit run --all-files     # the CI "lint" job
```

Rust: `cargo fmt`, `cargo clippy -D warnings`. Python: ruff on everything;
pyupgrade, isort, black and mypy on the code written for openrsxl only -
openpyxl's ported code and test-suite keep openpyxl's formatting so that new
openpyxl releases can be diffed and merged (configuration in
`pyproject.toml` and `.pre-commit-config.yaml`).

## Benchmarks

```bash
python tools/make_bench_files.py benchfiles --rows 100000
python tools/bench_all.py benchfiles/*.xlsx                 # openpyxl vs openrsxl
python tools/bench_tasks.py some.xlsx                       # full mode vs openrsxl.extended
python tools/bench_extended.py some.xlsx all --cells        # one streaming run
python tools/bench_write.py openrsxl [--write-only]
```

Every measurement runs in a fresh process; memory is the peak resident set
size and its increase over the process baseline.

## CI and releases

`.github/workflows/CI.yml`:

* `test` - Linux, CPython 3.10 - 3.14 and 3.14t: the test-suite (lxml and
  ElementTree back-ends, Python fallbacks) and openpyxl's ported suite;
* `lint` - `pre-commit run --all-files`;
* `build` - wheels for Linux (manylinux / musllinux: x86_64, aarch64, i686,
  armv7, ppc64le, s390x), macOS (x86_64, arm64) and Windows (x64, x86,
  arm64), plus the sdist; `test-wheels` installs them on clean machines (no
  Rust) and runs the tests;
* `build-pyemscripten` / `test-pyemscripten` - Pyodide wheel (allowed to
  fail);
* on a tag: `release` publishes everything to PyPI (repository secret
  `PYPI_API_TOKEN`) and `gh-release` creates the GitHub release.

To release: bump `version` in `Cargo.toml` (the Python package takes it from
there), add a `CHANGELOG.md` entry, commit, then `git tag vX.Y.Z && git push
origin main vX.Y.Z`.

## How exactness is guaranteed

* The Python layer *is* openpyxl's code (renamed imports), so every class,
  default, validation rule and serialisation is identical.
* The Rust engine replaces hot paths only and reproduces Python semantics
  explicitly (`int()`/`float()` parsing, `"%.16g"`, round-half-even, expat
  normalisation, ElementTree/lxml escaping, 16 KB iterparse blocks, ...).
* Whenever the engine is not certain - malformed XML, DOCTYPEs, non UTF-8
  encodings, unusual values, *any* error while reading cells - it raises
  `NativeFallback` internally and the original Python code is used, so
  errors, warnings and results are openpyxl's.
* Differential tests compare against openpyxl itself (a test dependency
  only).

## Porting a new openpyxl release

1. `tools/port_upstream_tests.py <openpyxl checkout>` regenerates
   `tests/upstream`.
2. Diff the new openpyxl sources against `python/openrsxl` (renamed imports)
   and merge; the openrsxl specific changes are marked with `openrsxl:`
   comments and live mainly in `worksheet/_reader.py`, `_read_only.py`,
   `_writer.py`, `worksheet.py` and `reader/strings.py`.
3. Run all suites in both XML modes, then a corpus run.
