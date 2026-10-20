# Workbook operations

Start with:

```bash
sheetops schema
sheetops inspect --json '{"path":"workbook.xlsx"}'
sheetops read-table --json '{"path":"workbook.xlsx","sheet":"Data","headerStrategy":"row","headerRow":1,"maxRows":50}'
sheetops cells read --json '{"path":"workbook.xlsx","sheet":"Summary","range":"A1:D20","include":["values","formulas"],"engine":"xml","valueMode":"typed"}'
sheetops cells search --json '{"path":"workbook.xlsx","query":"#4BACC6","fields":["fillColors"],"sheet":"Summary","range":"A1:Z200","limit":100,"offset":0}'
sheetops styles assert --json '{"path":"workbook.xlsx","sheet":"Merchant","checks":[{"cell":"E2","fill":"#D9EAF7"},{"cell":"H2","commentContains":"Method of Payment Code"},{"cell":"G3","fill":"#D9EAF7","formula":true}]}'
sheetops validate --json '{"path":"workbook.xlsx","requiredSheets":["Data"],"requiredColumns":{"Data":["Amount"]},"uniqueKeys":{}}'
sheetops formulas scan --json '{"path":"output.xlsx","sheet":"Summary","range":"A1:D50","include":["errors"],"failFast":true,"engine":"auto"}'
sheetops verify-edits --json '{"originalPath":"original.xlsx","path":"output.xlsx","allowedRanges":{"Data":["B2:B20"]},"evaluatedPath":"output.evaluated.xlsx"}'
```

Other available commands:

```bash
sheetops normalize preview --json '{"path":"input.xlsx","sheet":"Data","headerStrategy":"detect","columns":[{"source":"Amount","kind":"amount","name":"amount"}]}'
sheetops doctor libreoffice
sheetops recalc --json '{"path":"output.xlsx","outputDir":"outputs/recalc"}'
sheetops diff --json '{"left":{"path":"actual.csv"},"right":{"path":"expected.csv"},"keyColumns":["reference"],"compareColumns":["amount"],"amountTolerance":"0.00"}'
sheetops reconcile --json '{"left":{"path":"bank.csv"},"right":{"path":"processor.csv"},"leftKeyColumns":["reference"],"rightKeyColumns":["reference"],"leftAmountColumn":"amount","rightAmountColumn":"amount","amountTolerance":"0.01"}'
```

Run `sheetops doctor libreoffice` before `sheetops recalc`. If the doctor
reports that LibreOffice is missing or broken, run `scripts/setup-libreoffice`
and follow the printed platform-specific install command.


Create and edit in separate paths. Use XlsxWriter for new reports and a compatible editing library for existing workbooks. Preserve unrelated cells and formatting. Verify edits against the original with explicit allowed ranges. Evaluate formula outputs in an actual calculation engine and check independent expected results; cached values and successful serialization do not prove calculation or business correctness. Block formula delivery until evaluation succeeds.
