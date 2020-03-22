#This is a LibreOffice Calc macro for automated find and replace countries so the country get recognized by Dtawrapper with the english name
#Made from Prahlad Yeri tutorial on medium : https://medium.com/@prahladyeri/ten-useful-libreoffice-macro-recipes-72732ad210fa

#hopkins = name found in Johns Hopkins CSSE database (find)
#datawrapper = name for datawrapper maps (replace)

Sub find_replace_countries
  Dim hopkins() As String
  Dim datawrapper() As String
  Dim n As Long
  Dim document As Object
  Dim sheets as Object
  Dim sheet as Object
  Dim replace As Object

  hopkins() = Array("Egypt", "Brunei", "Cabo Verde", "Czechia", "Korea, North", "Congo (Kinshasa)", "Congo (Brazzaville)", "Iran", "Kyrgyzstan", "Venezuela" ,"Korea, South", "Russia")
  datawrapper() = Array("Arab Republic of Egypt", "Brunei Darussalam", "Cape Verde", "Czech Republic", "D. P. R. of Korea", "Democratic Republic of Congo", "Congo", "Islamic Republic of Iran", "Kyrgyz Republic", "R. B. de Venezuela", "Republic of Korea", "Russian Federation")
  document = ThisComponent rem .CurrentController.Frame
  rem sheet = doc.CurrentSelection.Spreadsheet
  sheets = document.getSheets()
  sheet = sheets.getByIndex(0)
  replace = sheet.createReplaceDescriptor rem document.createReplaceDescriptor in case of Writer
  rem replace.SearchRegularExpression = True
  For n = lbound(hopkins()) To ubound(hopkins())
    replace.SearchString = hopkins(n)
    replace.ReplaceString = datawrapper(n)
    sheet.replaceAll(replace)
  Next n
  MsgBox("Done")
End Sub
