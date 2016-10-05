<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/conn_webprodmx.asp" -->
<%
Dim rsCategories
Dim rsCategories_numRows

Set rsCategories = Server.CreateObject("ADODB.Recordset")
rsCategories.ActiveConnection = MM_conn_webprodmx_STRING
rsCategories.Source = "SELECT * FROM dbo.tbl_Categories ORDER BY Category ASC"
rsCategories.CursorType = 0
rsCategories.CursorLocation = 2
rsCategories.LockType = 1
rsCategories.Open()

rsCategories_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsCategories_numRows = rsCategories_numRows + Repeat1__numRows
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<table border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td>CategoryID</td>
    <td>Category</td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT rsCategories.EOF)) %>
  <tr>
    <td><%=(rsCategories.Fields.Item("CategoryID").Value)%></td>
    <td><%=(rsCategories.Fields.Item("Category").Value)%></td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsCategories.MoveNext()
Wend
%>
</table>
</body>
</html>
<%
rsCategories.Close()
Set rsCategories = Nothing
%>
