#This is a LibreOffice Calc macro for automated find and replace countries with multiple entries, so the data sum up
#Made from Prahlad Yeri tutorial on medium : https://medium.com/@prahladyeri/ten-useful-libreoffice-macro-recipes-72732ad210fa

#names = name of the country with multiple entries, in JH CSSE database
#cnames = new name that will permit the sum of the multiple entries in ods sheet (where are the formulas)

Sub find_replace_provCountry
  Dim names() As String
  Dim cnames() As String
  Dim n As Long
  Dim document As Object
  Dim sheets as Object
  Dim sheet as Object
  Dim replace As Object

  names() = Array("Australia", "Canada", "China", "Denmark", "France", "Netherlands", "United Kingdom", "US")
  cnames() = Array("C-Australia", "C-Canada", "C-China", "C-Denmark", "C-France", "C-Netherlands", "C-United Kingdom", "C-US")
  document = ThisComponent rem .CurrentController.Frame
  rem sheet = doc.CurrentSelection.Spreadsheet
  sheets = document.getSheets()
  sheet = sheets.getByIndex(0)
  replace = sheet.createReplaceDescriptor rem document.createReplaceDescriptor in case of Writer
  rem replace.SearchRegularExpression = True
  For n = lbound(names()) To ubound(names())
    replace.SearchString = names(n)
    replace.ReplaceString = cnames(n)
    sheet.replaceAll(replace)
  Next n
  MsgBox("Done")
End Sub
