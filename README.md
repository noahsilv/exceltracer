# Formula Trace for Excel on Mac

Explore a formula with the keyboard. Control+Q lists its precedents and Control+Shift+Q lists its dependents. Up and Down activate each referenced sheet and select the relevant cell or range. Right follows a reference one level deeper. Left returns to the previous list at the step you left. Control+Shift+M toggles a formula map on the active sheet.

**Arixcel parity is incomplete.** The [feature audit and desktop acceptance cases](PARITY.md) list the implemented behaviors, missing commands and open Mac checks.

## Run on a Mac

1. Use Microsoft 365 desktop Excel for Mac with add-ins enabled and Node.js 22 or newer. The same portable files run on Intel and Apple silicon with the corresponding Node installation.
2. Save your work and quit Excel before the first launch so Excel picks up the new manifest. Clone this repository or extract its ZIP. Open **Start Formula Trace.command**. From Terminal you can run `bash "Start Formula Trace.command"` in the project folder. The compiled pane and its launcher are included.
3. Approve the Keychain prompt for the **Formula Trace localhost** certificate. The launcher serves the add-in on HTTPS port 3000 and copies its manifest to Excel’s per-user sideload catalog. Excel loads Office.js from Microsoft’s CDN.
4. The launcher opens **Debt-Model-Mac.xlsx**. In Excel, select **Home → Add-ins → Formula Trace**. If the catalog has not refreshed, quit and reopen Excel while the launcher is running.
5. Select **Summary!B2**, press **Control+Q**, then press **Down** twice. The references are **Cash Flow!B4** and **Debt!B4**. From **Summary!B2**, **Control+Shift+Q** gives the two reverse links on those sheets. Use **Control**, since **Command+Q** quits Excel.

Keep the launcher window open during use. Control+C stops the server and removes this session’s certificate from the login keychain. After an abrupt exit the next launch checks for its previous certificate. To remove the sideload registration, remove `030c881e-c1db-4256-b653-21a708cec84d.manifest.xml` from `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/` with Excel closed.

[Microsoft’s Mac sideloading instructions](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)

## Run on Windows

1. Use your existing Node.js 22 or newer and a current Microsoft 365 desktop Excel with developer add-ins enabled. Excel uses the Office.js runtime from Microsoft's CDN.
2. Extract the whole portable ZIP to a new folder. Double-click **Start Formula Trace.cmd**. It starts the included compiled add-in using Node and serves the pane on https://localhost:3000. The portable launcher skips npm entirely.
3. Windows may ask you to trust the **Formula Trace localhost** certificate for your current account. The launcher registers the included manifest with Excel and opens **Debt-Model.xlsx** at **Summary!B2**. Select **Home → Explore Formula** if the pane is closed.
4. Keep the launcher open while tracing. Press **Ctrl+C** in the launcher to stop the server and remove the session certificate. If the window closes abruptly, the next run removes the previous session certificate. The private key exists only in the server process.
5. To remove this development add-in from your Excel account, stop the launcher and run this command in Command Prompt.

```bat
reg delete "HKCU\Software\Microsoft\Office\16.0\Wef\Developer" /v 030c881e-c1db-4256-b653-21a708cec84d /f
```

If port 3000 is busy, close the earlier launcher window before retrying. To check that the pane is served, open https://localhost:3000/help.html while the launcher is running.

## Desktop acceptance run

Close the earlier launcher and its sample workbook before starting this build. Use a workbook with Formula Trace loaded. If Excel asks which action owns a shortcut, choose the Formula Trace command to try it. The Home ribbon also includes **Explore Formula** and **Toggle Formula Map**. Worksheet shortcut assignments can be edited in the pane under **Keyboard shortcuts** on supported Excel versions.

| Action | Expected result |
| --- | --- |
| On the **Map example** sheet press **Ctrl+Shift+M**. | The revenue formulas in C4 and D4 take the copied-across color. E4 takes the hardcoded-number color. |
| Press **Ctrl+Shift+M** again. | The original pink fill and bold font on E4 are visible. On a disposable copy, save and reopen with the map on and repeat the removal check. |
| Select **Summary!B2**, press **Ctrl+Q**, then arrow down through the tree. | DSCR is 1.50x. The references to Cash Flow!B4 and Debt!B4 show 150 and 100. Previewing a reference selects it on the worksheet while the summary formula stays in the pane. |
| With Debt!B4 selected, press **Right**, **Down**, then **Left**. | Right opens the debt-service inputs, Down previews interest expense, and Left returns to the DSCR list at Debt!B4. |
| Select **Summary!B2** in the grid and press **Ctrl+Shift+Q**, then **Down** twice. | The two dependent formulas are on Cash Flow!B7 and Debt!B7. Each step activates the corresponding sheet and selects B7. Excel supplies the dependent order. |
| Select **Map example!B6**, press **Ctrl+Q**, and enable **Include formula components**. | IF has condition, true, and false branches. Left and right collapse and expand branches. Home selects the root. Esc returns to the explored cell and closes. Enter closes at the highlighted reference. |

The on-sheet highlight is Excel’s native range selection. Each preview activates the referenced worksheet and selects the full address. On supported desktop builds the explorer also asks Excel to scroll its first cell into view, then requests focus for the current pane row. Reference navigation does not write fills or add formatting rules. The explored formula stays above the selected reference’s formula. Hidden or protected cells that Excel cannot select are identified in the status line.

The map uses a snapshot of the active sheet. Toggle off and on to refresh after edits. The rules are saved with the workbook when you save a mapped sheet. Colors and enabled categories are editable in **Map colors** and take effect on the next application. Formula categories compare relative R1C1 formulas to the formula immediately above and to the left. Unique formulas can be marked as cross-sheet, external, or literal-only. The map preserves the cell’s base fill and removes its own conditional-format rules on the second toggle.

The explorer uses one task pane. It traces a cell, a contiguous block or a disjoint worksheet selection in the active workbook and pages through all direct reference areas returned by Excel, 100 explorer rows at a time. The first 100 area values are prefetched. Later values load as you browse. The default view lists cell and range references. The optional formula tree parses common operators, parentheses, and function calls. Excel supplies values for the root and traced cells. Full intermediate-expression evaluation, in-tree editing, tracing into other workbooks, mapping grouped sheets, and version comparison are outside this build. Ctrl+Backspace is handled inside the explorer; the default return shortcut from the grid is Ctrl+Alt+B.

## Verification and source

For development, run `npm ci`. Run `npm run build:portable` to rebuild the compiled pane and bundled localhost certificate helper. Run `npm test` to compile TypeScript and execute the Office adapter and navigation tests against a simulated Excel host. Run `npx playwright install webkit` and `npm run test:browser` for the compiled pane’s keyboard test with a simulated Office host. Run `npm run validate` to submit the development manifest to Microsoft's validator.

The included workbook has synthetic formulas and cached results. To regenerate it, install Python's XlsxWriter and run `python scripts/make_fixture.py` after `npm ci`. Development sideload metadata in the workbook points to this manifest.

The bundled certificate helper uses node-forge under its bundled license. The Windows sample uses workbook sideload metadata from Microsoft’s Office Add-in Scripts templates. Its license is in the third-party directory.

The portable Windows launcher has been opened in desktop Excel. Automated tests cover reference traversal, formula maps and Mac launcher setup. The Mac launcher test simulates Keychain consent and Excel launch and verifies the actual local HTTPS response and catalog files. Interactive Excel for Mac behavior, including focus after a sheet switch, still requires a desktop acceptance run.
