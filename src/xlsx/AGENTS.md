# XLSX Layer - High-Level Excel API

## Package Identity
- **Purpose**: High-level XLSX read/write operations
- **Tech**: TypeScript async iterators, streaming architecture
- **Exports**: `readXlsx()`, `writeXlsx()`, `Workbook`, `Sheet` classes

## Setup & Run
```bash
# Test this layer specifically
bun test src/xlsx/

# Run XLSX benchmarks
bun run src/tests/benchmarks/write-sheet-xml.bench.ts
```

## Patterns & Conventions

### Workbook Operations
- ✅ **DO**: Use async generators for large datasets
  ```typescript
  // src/xlsx/writer.test.ts example
  await writeXlsx('file.xlsx', {
    sheets: [{
      name: 'Data',
      rows: (async function* () {
        yield row([cell('Name'), cell('Age')]);
        yield row([cell('Alice'), cell(30)]);
      })()
    }]
  });
  ```

- ❌ **DON'T**: Load entire datasets into memory
  ```typescript
  // Anti-pattern: blocks memory
  const allRows = await loadAllDataFromDatabase();
  await writeXlsx('file.xlsx', { sheets: [{ name: 'Data', rows: allRows }] });
  ```

## Touch Points / Key Files
- **Main API**: `writer.ts`, `reader.ts` - Entry points
- **Data structures**: `workbook.ts` - Workbook/Sheet classes
- **Types**: `types.ts` - All interfaces and options
- **Structure**: `structure.ts` - XLSX file organization
- **Properties**: `properties.ts` - Document metadata
- **Strings**: `shared-strings.ts` - String deduplication
- **Caching**: `shared-strings-caching/` - Memory-aware caching strategies for shared strings

## JIT Index Hints
- Find writer function: `rg -n "export.*writeXlsx"`
- Find reader function: `rg -n "export.*readXlsx"`
- Find workbook class: `rg -n "export.*class.*Workbook"`
- Find sheet class: `rg -n "export.*class.*Sheet"`
- Find test examples: `rg -n "writeXlsx.*testFile" src/xlsx/*.test.ts`

## Common Gotchas
- Workbook sheets are 0-indexed for access but 1-indexed in Excel UI
- Shared strings reduce file size but increase complexity - use if file size is prioritised over speed
- Sheet names must be unique and valid Excel identifiers
- Empty sheets still generate valid XML structure

## Pre-PR Checks
```bash
bun test src/xlsx/ && bun run lint src/xlsx/
```
