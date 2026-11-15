# BreedX VBA

VBA code for `coding Breedx.xlsm`.

| File | VBA component | What it is |
|---|---|---|
| `vba/modSales.bas` | Module `modSales` | Sales entry form: data access, validation, `ShowSalesForm`, `CreateLaunchButton` |
| `vba/clsDayButton.cls` | Class `clsDayButton` | Click handler for each calendar day in the form |
| `vba/frmSalesEntry.txt` | UserForm `frmSalesEntry` | Code to paste into an empty UserForm named `frmSalesEntry` |
| `vba/modPepperLines.bas` | Module `modPepperLines` | Reads the Pepper lines from a proforma PDF with Excel's PDF connector; `TestReadPepperLines` |
| `vba/Module2.bas` | Module `Module2` | `CreateSeedOrder` (builds the seed order workbook); uses `modPepperLines` |

## Installation

Import the `.bas` and `.cls` files with **File > Import File...** in the VBA editor
(Alt+F11). Don't copy and paste them: the `Attribute ...` lines at the top are only
valid when imported. Remove an existing copy of a module first (right-click >
Remove, answer No to export), or the import creates `modSales1` and so on.

For the form, **Insert > UserForm**, set **(Name)** to `frmSalesEntry`, and paste
the whole content of `vba/frmSalesEntry.txt` into its code window.

Select the target project (`coding Breedx.xlsm`) before importing, then run
**Debug > Compile VBAProject**.

## Sales entry form

- Product and Company drop-downs come from the `Table` sheet (columns A and B).
- Delivery Date defaults to today and is picked from a pop-up calendar.
- Quantity of Seeds is a whole number from 0 to 1,000,000.
- Sales Representative is a drop-down from the `Table` sheet, column C.
- Submit appends the record to the last row of `Master` with the update time.
  Master columns: A Product, B Company, C Delivery Date, D Quantity,
  E Update Time, F Sales Representative.
  Product + Company + Delivery Date is the primary key; a duplicate asks before
  replacing the old row.

The data workbook is set at the top of `modSales`:

```vba
Public Const DATA_FOLDER As String = "C:\Users\YoavEilenberg\Breedx LTD Dropbox\IT\Sales"
Public Const DATA_FILE As String = "BreedX Peppers sales examples.xlsm"
Public Const SAVE_AFTER_SUBMIT As Boolean = True
```

If that workbook is already open it is used; otherwise it is opened from
`DATA_FOLDER`. It is saved after each submit, and the form refuses to write to a
read-only copy.

## Seed order from a proforma PDF

`CreateSeedOrder` reads the invoice's Pepper lines with `modPepperLines`
(Power Query `Pdf.Tables`, Microsoft 365 / Excel 2021 or later). Word's PDF
conversion of the Invoice4U PDFs exposes only the footer text, so Word is used
only as a last-resort fallback.
