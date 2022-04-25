# NOTES

Working with workbooks.

Workbook is a collection of Documents/Sheets.

Names: [Book]Sheet

Selection: xltype = xltypeRef | xltypeSref

```
Workbook wb;
wb.New(); // create a new workbook
Sheet s = wb.Insert("Sheet"); // insert and select sheet
Selection sel = s.Select("A1"); // select cell/range in sheet
sel = 1.23;
sel.Move(Down);
sel = "foo";
sel.Move(Right);
sel.Type("A comment.", 0.1s); // type 1 char every 1/10 seconds
```