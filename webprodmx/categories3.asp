<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
' *** Logout the current user.
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("MM_Username")
  Session.Contents.Remove("MM_UserAuthorization")
  MM_logoutRedirectPage = "login.asp"
  ' redirect with URL parameters (remove the "MM_Logoutnow" query param).
  if (MM_logoutRedirectPage = "") Then MM_logoutRedirectPage = CStr(Request.ServerVariables("URL"))
  If (InStr(1, UC_redirectPage, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
    MM_newQS = "?"
    For Each Item In Request.QueryString
      If (Item <> "MM_Logoutnow") Then
        If (Len(MM_newQS) > 1) Then MM_newQS = MM_newQS & "&"
        MM_newQS = MM_newQS & Item & "=" & Server.URLencode(Request.QueryString(Item))
      End If
    Next
    if (Len(MM_newQS) > 1) Then MM_logoutRedirectPage = MM_logoutRedirectPage & MM_newQS
  End If
  Response.Redirect(MM_logoutRedirectPage)
End If
%>
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
<title>Categories</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<table width="400" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>Welcome <%= Session("MM_Username") %>&nbsp;<a href="<%= MM_Logout %>">Logout</a></td>
  </tr>
<% If TRIM(Session("MM_UserAuthorization")) = "Administrator" Then %>
  <tr>
    <td><a href="insert_category.asp">Add a category</a> </td>
  </tr>
<% End If %>
</table>
<table width="400"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="171">Category ID </td>
    <td width="215">Category</td>
  </tr>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT rsCategories.EOF)) 
%>
  <tr>
    <td><%=(rsCategories.Fields.Item("CategoryID").Value)%></td>
    <td><a href="books.asp?CategoryID=<%=(rsCategories.Fields.Item("CategoryID").Value)%>"><%=(rsCategories.Fields.Item("Category").Value)%></a></td>
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
