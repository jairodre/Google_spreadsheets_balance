# Purchase Tools for Google Sheets

Create purchase orders in Google Docs and monthly summaries in Google Sheets. The repository contains Apps Script source and the definitions that build the blank purchase workbook and order document. Each installation creates native Google files in its own Drive.

## What you configure

**A manual installation requires no edits to `PurchaseTemplate.gs`, no pasted Drive IDs, and no API keys.** Install the script from **Extensions → Apps Script** in the working spreadsheet and authorize it with a Google account that can edit the workbook and create files in the purchase folder.

| Setting | How it is supplied |
| --- | --- |
| Working spreadsheet ID | Read automatically from the spreadsheet where you run Purchase Tools. |
| Order template, Received Log, and output folder IDs | Resolved automatically in Drive. The initializer finds or creates the named assets in the main purchase folder and places the working Sheet in `General_Purchase_Orders`. Keep the names and nesting shown below. |
| Other project spreadsheet IDs | **Sync projects in this folder** reads the IDs of the workbooks you select and saves them in their private **Projects** tabs. |
| Grant number | Enter it in **Initialize workbook**. It is stored in `Control!B5`, and every vendor tab displays it in `A1`. |
| Project name and project code | Enter both in the same dialog. They are stored in `Control!B3` and `Control!B4`. Use a unique code of 2–10 letters or digits for each project. |

The portable version builds the blank workbook and a Stage 1-style order document from code. The order, summary, and project-sync workflows carry over to the portable version.

Workbook setup and monthly report generation use the **United States spreadsheet locale** to match the Stage 1 formula syntax and formatting. This changes how Sheets displays and interprets dates and numbers. Your Google account interface language and workbook time zone are separate settings.

## Install

1. **Create a new blank Google Sheet.** In Google Drive, create a main purchase folder with a name of your choice. Open that folder and choose **New → Google Sheets → Blank spreadsheet**. Give the spreadsheet your project's name. This becomes your working purchase workbook.
2. **Install the script in that new spreadsheet.** With the blank Sheet open, choose **Extensions → Apps Script**. Replace the starter code in `Code.gs` with the contents of `PurchaseTemplate.gs` and save. The default Apps Script manifest is sufficient.
3. Reload the spreadsheet and look for **Purchase Tools** next to **Help**.

   **If Purchase Tools is missing after reloading:**

   1. **Zoom out in the browser first.** Press **⌘−** on Mac or **Ctrl−** on Windows and Linux once or twice, or maximize the browser window. The zoom control inside Sheets only changes the grid.
   2. In that workbook, open **Extensions → Apps Script**. Confirm the full script is saved, then reload or reopen the spreadsheet with edit access.
   3. **As a last resort**, select `onOpen` in the Apps Script function dropdown and click **Run**. Return to the spreadsheet to check for the menu. If it’s still missing, check **Executions** for the error.

   Once the menu appears, run **Purchase Tools → Initialize Drive templates and folders** and authorize the script when prompted. It creates the missing subfolders and templates, including Received Log, then moves the working Sheet into `General_Purchase_Orders`. For additional workbooks, either start in the main folder and run the initializer again or create the workbook directly inside the existing `General_Purchase_Orders` folder.
4. Run **Purchase Tools → Initialize workbook**. Enter the grant number, project name, and unique 2–10 character project code. Uncheck any vendor tabs you want to skip, then click **Continue**. Keep the progress window open until it shows the result. All optional vendors start checked. **Others**, **Shipment**, **Control**, and **Projects** are always created. The workbook remembers your choices, and you can add more vendors later. Setup reuses the default empty starter tab when possible, links the shared Drive assets, and installs the Received Log edit trigger when available. Use an account that can edit the shared Log.
5. Verify the grant in `Control!B5`, the project name in `Control!B3`, and the project code in `Control!B4`. Add a sample purchase on a vendor tab and generate an order and a monthly summary before using the workbook for production.

```text
Your purchase folder/
├── General_Purchase_Orders/     working Google Sheets
├── Monthly_summaries/
│   └── Archive_Summary_changes/
├── Orders_forms/                generated order documents
└── Templates/
    ├── Template_for_Purchase_Orders   native Google Sheet
    ├── Template_form                  native Google Doc
    └── Received_Log                   native Google Sheet
```

`Template_for_Purchase_Orders` contains the blank layout and formulas. Install the GitHub script in each working workbook to enable Purchase Tools. You can adjust the styling of `Template_form` in Drive. Keep the `{{ITEM_START}}` and `{{ITEM_END}}` markers in separate body paragraphs and keep the repeating item content between them as paragraphs. Existing files and folders with these names are reused when you run the initializer again.

The **Control** tab links to the order output, monthly summaries, summary archive, order document template, blank workbook template, and Received Log. The installer resolves these links from your Drive folders and refreshes them during setup.

To apply the Stage 1 order spacing to an existing installation, update the script and rerun **Initialize Drive templates and folders**. It updates `Template_form` when its body text matches either older generated layout. If you customized only the styling of an older layout, save a copy of the template before rerunning the initializer. The update applies to newly generated orders.

The generated order template uses the Stage 1 layout, with company and grant at the top, followed by category and item count, then each item’s link, unit cost, quantity, subtotal, and the order total. Templates with customized text are preserved. Drive initialization also refreshes the blank workbook template with the grant field in `Control!B5`. Each working workbook gets the latest commands when you replace its script and reload.

## Manage vendor tabs

The vendor commands appear in **Purchase Tools** below **Sync projects in this folder**, in the order **Add**, **Remove**, **Rename**.

- **Add vendor tab** when a company needs its own purchase tab. Run it from any tab and enter the company name and a unique code of 2–6 letters or digits. The code is saved in uppercase and used in the Item number. The new tab starts blank with the standard layout, grant link, and formulas. Orders, monthly summaries, workbook checks, and Received Log recognize it automatically. For occasional purchases, use **Others** and enter the company on the purchase row.
- To add a built-in vendor you skipped during initialization, enter its original name and code, such as **Amazon (AMZ)** or **DigiKey (DIG)**. The workbook then expects that tab during repair.
- **Remove vendor tab** when you no longer need Amazon or a custom vendor tab and are ready to delete its purchase rows. Open the target tab, verify its name in the warning, and type `DELETE_TAB` exactly. The command updates the Item number formulas that reference the tab. Workbook repair treats a menu removal as intentional. Saved orders, reports, archives, and log entries remain available, while future or regenerated summaries omit the deleted rows. Keep the other built-in tabs.
- **Rename vendor tab** to correct a company name or repurpose an empty regular vendor tab. Open the target tab and review the prefilled name and code. Keep the code to retain the existing Item number prefix. Changing the code updates the visible Item numbers, while fixed Purchase Record IDs stay with their rows. Future orders, summaries, and log entries use the new name, including for purchases already on the tab. Saved files and existing log entries retain the original name. On regeneration, a Department total under the former name stays in the archive if no matching vendor remains. Amazon can be renamed. Others and Shipment keep their names because they have special roles.

## Use

- For an order, open the vendor tab, fill the purchase fields through **Final price** and the **Category**, then use **Generate order from current vendor tab**. Enter item suffixes, a range such as `0001-0003,0007`, or `ALL`. Priority items appear first with a soft red highlight in the generated document.
- Orders with multiple categories display **Mixed**. A grant number is required for both orders and monthly summaries.
- For a summary, enter each purchase's **Reported on** date and choose its month and year in **Generate monthly summary**. Each reporting folder has one workbook per year, such as `Monthly_Summary_2026`, with a tab for each generated month, such as `2026-11`. The report header lists the included projects and codes.
- Regenerating an existing month archives only its previous tab as a standalone workbook, starting with `Monthly_Summary_2026-11_recreated_v0` and incrementing for that month. This also happens when the data is identical. Review the project changes shown before confirming a replacement. Generation stops if an enabled project cannot be read.
- Enter the department’s reconciliation amounts under **Department total**. Regeneration carries forward the amount, including zero, for the same grant and vendor, including entries from the former **Connie** column. **Our total** and **Difference** recalculate. If a grant or vendor disappears, the dialog identifies entries kept only in the archive.
- The completion dialog shows the count and known total of paid purchases without a **Reported on** date across the selected projects. Those purchases stay out of the report’s monthly totals.
- Keep independent reporting groups in separate output folders. Earlier standalone monthly files stay in place. Generate each month again to populate the annual workbook. If a run is interrupted and leaves a `_build_` or `_previous_` tab, review that month and its archive before clearing the temporary tab and retrying.
- **Item number** includes the fixed Purchase Record ID in brackets when available, in a 365 px single-line column.
- Double-click **Ordered on** or **Reported on** to choose a date. Both columns use date validation. **Reported on** can stay blank until the purchase is assigned to a reporting month.
- To combine project workbooks, give each one a distinct project code and run **Sync projects in this folder**. Select the workbooks to link. Sync removes entries from each selected workbook’s **Projects** tab if the spreadsheet is absent from the current folder listing. This includes workbooks moved elsewhere, deleted, or no longer visible to the account running Sync. Workbooks still listed in the folder are kept even if their contents cannot be read. Select just one workbook to clean its list.
- Move complete purchase rows. The **Ordered on** date and **Purchase Record ID** stay with the row, while the Item number suffix follows the current item order. Clearing all purchase data in a row resets its Purchase Record ID. The repair command preserves existing rows, including rows beyond 1,000.
- **Received Log** records individual Paid, Arrived, and Notes edits. Every new entry captures Ordered on, Item, Link, Quantity, Unit price, and Final price as one clipped line in **Paid details**. Checking Arrived stamps a fixed date and unchecking it clears that date.

## Upgrade an existing portable workbook

Each workbook has its own attached Apps Script. Update the script separately in every workbook that uses Purchase Tools. You only need to do this once per workbook, even if it has several vendor tabs.

1. In each existing working workbook, open **Extensions → Apps Script**, replace the old script with the current `PurchaseTemplate.gs`, save, and reload the spreadsheet.
2. Run **Purchase Tools → Check / repair workbook**. If `Control!B5` is blank, repair copies the legacy grant from `Amazon!A1` before replacing vendor grant cells with links to Control. A grant already entered in `Control!B5` takes precedence.
3. Verify the grant in `Control!B5` and confirm that vendor tabs display it in `A1`. If the original Amazon tab was already removed or its grant is missing, enter the grant directly in `Control!B5`. If Amazon was renamed manually before the upgrade, restore its original name before repair, then rename it through Purchase Tools. Project name and code remain in `Control!B3:B4`.
4. Run **Initialize Drive templates and folders** if shared assets are missing or you want to refresh the blank workbook template. Use **Initialize workbook** when starting a blank working workbook.

If you previously installed the batch updater’s manifest, remove the `script.projects` and `script.external_request` OAuth scopes from that manifest if they were added for the updater. Preserve settings or scopes used by other code in the workbook.

**Check / repair workbook** checks vendor tabs in batches and writes repairs only where needed. It restores grant links, missing Purchase Record IDs, Item number formulas on purchase rows, and the script's Paid and Arrived color rule. It samples validation on the first and last active rows and repairs the affected column if those checks fail. A one-off validation change between the samples may go undetected.

The check repairs a broken Final price formula when the output column has no conflicting values. When it finds values that could be typed totals or results of a displaced array formula, it lists the cells and expected totals for review and leaves that column intact. It offers to recreate missing expected tabs as blank layouts after confirmation. Vendors skipped during initialization stay skipped. Older workbooks without a saved selection keep their existing defaults. A blank replacement tab cannot recover deleted purchases.

Missing or inaccessible Drive assets are reported separately. The check can finish repairing the workbook without creating any Drive files. An existing `Received_Log` is reused and must have the current headers in its `Log` tab.

**Check / repair workbook** shows the current step and the number of steps completed, including individual vendor tabs. The bar measures completed steps, since different steps can take different amounts of time. Initialization, Item number refresh, and project sync show a moving activity bar. Completion details, warnings, and errors appear in the same window. Missing tabs still require confirmation before recreation. Sheets may briefly show its own running-script banner.

If Sheets closes the **Initialize workbook** dialog, setup may still be running. Open **Purchase Tools → Initialize workbook** again to check progress or view the recent result. Sheets also shows a completion or error notification.

## Publish or deploy

### Files in this repository

| File | What it does | Where it goes |
| --- | --- | --- |
| `PurchaseTemplate.gs` | Runs Purchase Tools, including setup, order generation, monthly summaries, and Received Log updates. | Commit it to GitHub. For a manual install, paste its contents into `Code.gs` under **Extensions → Apps Script** in your workbook. |
| `README.md` | Explains installation, daily use, and troubleshooting. | Commit it to GitHub so it appears on the repository’s main page. |
| `appsscript.json` | Optional Apps Script manifest with the V8 runtime and exception logging. | Keep it in GitHub for reference. Google’s default manifest is sufficient for a manual installation. |
| `test/portable.test.cjs` | Runs local checks with sample data and simulated Google services. | Keep the `test/` folder in GitHub for development. Only Node.js runs it. |
| `.gitignore` | Keeps common local settings and credential files out of new Git commits. | Commit it to GitHub. |

### Run the local tests

With Node.js installed, open a terminal in the repository folder and run:

```sh
node --test
```

The tests check workbook setup and repair, annual summaries and monthly archives, and whether regeneration preserves a **Department total** of zero. They use sample data and run without a Google login or access to live Drive files. Use a test workbook in Google Sheets to verify authorization, file creation, and the appearance of generated reports.

### Publish to GitHub

Before committing, review the files you’re about to upload. Keep production workbooks, private Drive links and IDs, grant numbers, generated reports, OAuth files, API keys, and local deployment settings out of Git. A `.gitignore` entry only affects untracked files. Previously committed content remains in history.

### Optional developer tool

`clasp` lets developers push local code changes directly to an Apps Script project. Keep `.clasp.json` and your local authorization file `.clasprc.json` out of Git. [Google’s clasp guide](https://developers.google.com/apps-script/guides/clasp)

## TODO

- Improve centralized script updates so a maintainer can update one shared `.gs` source and roll out the new version to all participating workbooks in one operation.
