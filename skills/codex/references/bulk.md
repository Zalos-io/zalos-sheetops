# Typed bulk reads

For a complete, values-only XLSX table, use the optional `zalos-sheetops[bulk]`
Python backend after confirming the sheet and header row with a bounded XML
preview. Keep small ranges, formulas, and exact stored values on XML.

```python
from pathlib import Path
import polars as pl
from zalos_sheetops.io.read_bulk import read_bulk_table

batch = read_bulk_table(Path("source.xlsx"), "Data", header_row=1)
frame = pl.from_arrow(batch)
```

The fastexcel backend returns an Arrow RecordBatch with normalized headers and
typed values, not `read-table`'s strings. It retains interior blank rows and
text IDs, and requires explicit sheet/header coordinates. It rejects formulas,
Excel errors, mixed column types and unsafe large numeric values. For a justified
mixed text column, pass `column_types={"normalized header": "string"}`; this adds
an XML precision check and may be slower. Never coerce away an error to get a
result. Use XML to investigate rejected inputs. Formula evaluation and output
verification remain separate requirements. Bulk reads materialize the sheet;
use XML for bounded reads or when memory is constrained.
Text IDs already remain strings without overrides. Read the table once and
reuse the frame for inspection, transformation and control totals.
