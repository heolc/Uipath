# Uipath

### How to Find the End Row of Specific Column (F) in Excel using UiPath
    ( dt1.Rows.IndexOf( dt1.AsEnumerable.Where(Function(r) r("Fri").ToString.Trim <> "").ToArray.Last ) + 2 ).ToString
