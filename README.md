# yayoi-excel
Excel VBA macro to make Yayoi-Kaikei importable format CSV

Yayoi.xslm : Excel (Data holder) + VBA included
Yayoi.vba : VBA source (exported from .xslm)

[Usage]

Put suitable Shiwake data onto Excel sheet and call output2yayoi Macro.

Then "yayoi_import.txt" file will be generated.

On Yayoi Kaikei, open Shiwake-Nikki, click "File"-"Import" then select "yayoi_import.txt".

[Example]

sample/Paypal2Yayoi.xlsm

  Import Paypal CSV to Yayoi

  Caution: Data conversion is not automated.
