# Formula Trace verification

Formula Trace supports direct dependency tracing and single-sheet formula maps. Desktop Excel for Mac acceptance testing is pending.

## Evidence and release checks

- The adapter tests exercise single-cell, block and disjoint selections, drilling and returning, errors, paging, and scroll requests. The map tests exercise classification, rule cleanup and preservation of other rules in a simulated workbook.
- The browser test loads the compiled HTML and JavaScript over HTTPS. With a simulated Office host it sends successive arrow keys through a cross-sheet trace, checks the selected address and focused row, drills down and returns to the same sibling, traces both dependents, closes and returns to the original cell.
- The Mac launcher test checks the sideload manifest and real HTTPS response in a temporary home directory. On macOS it uses Apple’s certificate verifier. The test substitutes for interactive Keychain trust and the Excel launch. The workflow runs the browser test in WebKit on both Mac processor architectures. [Mac workflow](.github/workflows/test.yml)
- Desktop checks still require an interactive Microsoft 365 desktop Excel session on Mac. Office.js direct dependency APIs are scoped to the active workbook. Cross-workbook navigation and separate explorer windows need demonstrated designs before feature equivalence can be promised. [Microsoft’s dependency API limits](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-ranges-precedents-dependents)

## Desktop keyboard and map checks

Use a fresh copy of the same sample workbook in Excel. Record the add-in and Excel builds and whether a shortcut conflict prompt appeared. For each keystroke capture the active workbook, sheet, selected address, focused window or pane, displayed formula, highlight colors and whether the next keystroke works without a mouse.

| Case | Observable result to check |
| --- | --- |
| Precedents of Summary!B2 | Control+Q exposes Cash Flow!B4 (150) and Debt!B4 (100). Arrow through both, follow Debt!B4 into its inputs, and return to the same sibling. Check the active sheet and selection after every arrow and record any ordering difference. |
| Dependents of Summary!B2 | Control+Shift+Q exposes Cash Flow!B7 and Debt!B7. Move forward and backward. Check that the source formula stays visible and the grid highlights the expected B7 on each sheet. |
| Multiple source areas | On Debt select B2, B3 and an unrelated blank cell. Trace dependents. Both B4 and B5 use the inputs. Inspect grouping, include any shared dependent once, and verify Home restores the source areas. |
| Close and return | Run separate trials for Enter and Escape on a cross-sheet preview. Record the resulting selection and focus and the key needed to return from the grid. Formula Trace currently uses Control+Option+B on the Mac grid and Control+Backspace inside the pane. |
| Map and restore | On Map example compare C4 and D4 (copied formulas) with E4 (hardcoded). Record the categories and colors. Save with the map on, quit, reopen and remove it. E4 should retain its original pink fill and bold type. Repeat on grouped sheets and with pre-existing conditional formats. |
