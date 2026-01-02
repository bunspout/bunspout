# bunspout

A fast, streaming Excel read/write library for XLSX files. Built with Bun runtime for high-performance data processing.

## Features

- ":rocket:" **Streaming architecture** - Process large Excel files without loading everything into memory
- ":bar_chart:" **Full XLSX support** - Read and write Excel workbooks with multiple sheets
- ":wrench:" **TypeScript-first** - Comprehensive type safety with strict mode
- ":zap:" **Bun-optimized** - Built for Bun runtime with adapters for other environments
- ":chart_with_upwards_trend:" **Performance-focused** - Benchmarks included for measuring performance
- ":jigsaw:**Modular design** - Clean separation between XLSX, XML, ZIP, and sheet layers

Traditional XLSX writers often load the entire dataset into memory before writing. Bunspout uses AsyncIterable streams to process rows one by one, which allows:
	•	Low memory footprint
	•	Efficient handling of millions of rows
	•	Better performance in Bun/Node

## Installation

```bash
bun install
```

## Usage

### Writing Excel Files

```typescript
import { writeXlsx, cell, row } from 'bunspout';

// Create a simple workbook with one sheet
await writeXlsx('output.xlsx', {
  sheets: [
    {
      name: 'Data',
      rows: (async function* () {
        yield row([cell('Name'), cell('Age'), cell('City')]);
        yield row([cell('Alice'), cell(30), cell('New York')]);
        yield row([cell('Bob'), cell(25), cell('San Francisco')]);
      })(),
    },
  ],
});
```

### Reading Excel Files

```typescript
import { readXlsx } from 'bunspout';

// Node.js 20.6.0+ and Bun: automatic cleanup with await using
await using workbook = await readXlsx('input.xlsx');

// Access sheets by name or index
const sheet = workbook.sheet('Data'); // or workbook.sheet(0)

// Iterate through rows
for await (const row of sheet.rows()) {
  console.log(row.cells.map(cell => cell.value));
}
// Automatically cleaned up when exiting scope
```

**For Node.js < 20.6.0:** Explicitly call `cleanup()` when done:

```typescript
const workbook = await readXlsx('input.xlsx');
// ... use workbook ...
await workbook.cleanup(); // Release temporary files and zip file handles
```

### Working with Large Datasets

For large datasets, use streaming patterns:

```typescript
import { writeXlsx, cell, row, mapRows, filterRows } from 'bunspout';

// Process data from a large source
const rawData = getLargeDataset(); // Some async iterable

await writeXlsx('large-report.xlsx', {
  sheets: [
    {
      name: 'Filtered Data',
      rows: filterRows(
        mapRows(rawData, transformRow),
        row => row.cells.length > 0
      ),
    },
  ],
});
```


### Database Integration: Exporting from Prisma

A common scenario is exporting database records to Excel. Here's how to export Prisma query results:

```typescript
// ...
import { writeXlsx, cell, row } from 'bunspout';

// Query users from database
const users = await prisma.user.findMany();

// Transform database records to Excel rows
await writeXlsx('users-report.xlsx', {
  sheets: [
    {
      name: 'Users',
      rows: (async function* () {
        // Header row
        yield row([
          cell('ID'),
          cell('Name'),
          cell('Email'),
          cell('Created At')
        ]);

        // Data rows
        for (const user of users) {
          yield row([
            cell(user.id),
            cell(user.name),
            cell(user.email),
            cell(user.createdAt)
          ]);
        }
      })(),
    },
  ],
});
```

For large datasets, use streaming with Prisma cursors:

```typescript
import { writeXlsx, cell, row, mapRows } from 'bunspout';

const userStream = async function* () {
  let cursor: string | undefined;

  while (true) {
    const batch = await prisma.user.findMany({
      take: 1000,
      skip: cursor ? 1 : 0,
      cursor: cursor ? { id: cursor } : undefined,
      orderBy: { id: 'asc' }
    });

    if (batch.length === 0) break;

    yield* batch;
    cursor = batch[batch.length - 1].id;
  }
};

await writeXlsx('large-users-export.xlsx', {
  sheets: [
    {
      name: 'Users',
      rows: mapRows(userStream(), user => row([
        cell(user.id),
        cell(user.name),
        cell(user.email),
        cell(user.createdAt)
      ])),
    },
  ],
});
```


### Cell Types

Bunspout supports all Excel cell types:

```typescript
import { cell } from 'bunspout';

// Different cell types
cell('text');          // String
cell(42);              // Number
cell(new Date());      // Date
cell(true);            // Boolean
cell(null);            // Empty cell
```

## Limitations / Known Issues

Bunspout focuses on fast, streaming Excel operations. The package is currently in the alpha stage.
Here are current limitations:

### ":no_entry_sign:" **Limited Formula Evaluation**
- When reading, formula's computed values are read, if available
- No calculation engine - formulas won't execute in the generated Excel files
- Formulas are stored as text (strings starting with `=`)
- Use case: Data export with formula templates that users can edit

### ":art:**Limited Styling Support**
- Conditional formatting not supported
- Basic text formatting only (dates, numbers, booleans)
- Use case: Clean data export without visual styling

### ":bar_chart:**No Charts or Graphics**
- Pure data tables only
- No embedded images, charts, or shapes
- Use case: Data export for further processing in Excel/other tools

### ":1234:**Data Type Constraints**
- Large numbers may lose precision in Excel (Excel's limit: 15 significant digits) - bunspout passes numbers through without validation; Excel handles precision according to its own rules
- Very long text may be truncated in some Excel versions - bunspout does not limit text length; any truncation is Excel's behavior
- Date handling follows Excel's date serial number system

### ":chart_with_upwards_trend:**Performance Considerations**
- Shared strings mode is slower but produces smaller file sizes
- Column width auto-detection requires buffering (slightly slower)
- Memory usage scales with concurrent operations

### ":recycle:**Excel Compatibility**
- Generated files are valid XLSX but may not include all Excel features
- Some advanced Excel features (macros, data validation, etc.) not supported
- Focus on core spreadsheet functionality for maximum compatibility

**Need these features?** Consider post-processing the generated Excel files with libraries like ExcelJS or Apache POI, or contribute to bunspout's development!

## API Reference

### Core Functions

- `writeXlsx(filePath, workbookDefinition, options?)` - Write XLSX file
- `readXlsx(filePath)` - Read XLSX file and return Workbook

### Data Structures

- `cell(value)` - Create a cell with automatic type detection
- `row(cells, options?)` - Create a row from cells
- `Workbook` - Represents an Excel workbook
- `Sheet` - Represents a worksheet

### Utilities

- `mapRows(rows, mapper)` - Transform rows
- `filterRows(rows, predicate)` - Filter rows
- `collect(rows)` - Collect async iterable to array

## Development

```bash
# Run tests
bun test

# Run benchmarks
bun bench

# Lint code
bun run lint

# Auto-fix linting issues
bun run lint:fix
```

## Performance

Bunspout is designed for high performance with large Excel files:

- Streaming XML generation avoids memory bloat
- Async iterators process data on-demand
- Shared strings option reduces file size for repeated text
- Benchmarks included in `src/tests/benchmarks/`

## Contributing

See the hierarchical AGENTS.md files for development guidance:
- Root `AGENTS.md` - Universal conventions
- `src/*/AGENTS.md` - Layer-specific patterns and examples