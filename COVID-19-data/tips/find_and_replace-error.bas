#This is a LibreOffice Calc macro for fixing errors from the provCountries macro (US replacement causing error in names)
#Made from Prahlad Yeri tutorial on medium : https://medium.com/@prahladyeri/ten-useful-libreoffice-macro-recipes-72732ad210fa

#errors = name with errors
#fixed = new name with no error

Sub find_replace_errors
  Dim errors() As String
  Dim fixed() As String
  Dim n As Long
  Dim document As Object
  Dim sheets as Object
  Dim sheet as Object
  Dim replace As Object

  errors() = Array("AC-UStria", "C-AC-UStralia", "RC-USsian Federation", "CyprC-US", "Brunei DarC-USsalam", "BelarC-US", "MauritiC-US")
  fixed() = Array("Austria","C-Australia", "Russian Federation", "Cyprus", "Brunei Darussalam", "Belarus", "Mauritius")
  document = ThisComponent rem .CurrentController.Frame
  rem sheet = doc.CurrentSelection.Spreadsheet
  sheets = document.getSheets()
  sheet = sheets.getByIndex(0)
  replace = sheet.createReplaceDescriptor rem document.createReplaceDescriptor in case of Writer
  rem replace.SearchRegularExpression = True
  For n = lbound(errors()) To ubound(errors())
    replace.SearchString = errors(n)
    replace.ReplaceString = fixed(n)
    sheet.replaceAll(replace)
  Next n
  MsgBox("Done")
End Sub
