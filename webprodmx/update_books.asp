<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/conn_webprodmx.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="Administrator"
MM_authFailedURL="login.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (false Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_conn_webprodmx_STRING
  MM_editTable = "dbo.tbl_Books"
  MM_editColumn = "BookID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "categories3.asp"
  MM_fieldsStr  = "BookTitle|value|BookAuthorFirstName|value|BookAuthorLastName|value|BookPrice|value|BookISBN|value|BookPageCount|value|BookImage|value|CategoryID|value"
  MM_columnsStr = "BookTitle|',none,''|BookAuthorFirstName|',none,''|BookAuthorLastName|',none,''|BookPrice|none,none,NULL|BookISBN|',none,''|BookPageCount|none,none,NULL|BookImage|',none,''|CategoryID|none,none,NULL"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim rsUpdateBook__varBookID
rsUpdateBook__varBookID = "0"
If (Request.QueryString("BookID") <> "") Then 
  rsUpdateBook__varBookID = Request.QueryString("BookID")
End If
%>
<%
Dim rsUpdateBook
Dim rsUpdateBook_numRows

Set rsUpdateBook = Server.CreateObject("ADODB.Recordset")
rsUpdateBook.ActiveConnection = MM_conn_webprodmx_STRING
rsUpdateBook.Source = "SELECT BookID, BookTitle, BookAuthorFirstName, BookAuthorLastName, BookPrice, BookISBN, BookPageCount, BookImage, CategoryID  FROM dbo.tbl_Books  WHERE BookID = " + Replace(rsUpdateBook__varBookID, "'", "''") + ""
rsUpdateBook.CursorType = 0
rsUpdateBook.CursorLocation = 2
rsUpdateBook.LockType = 1
rsUpdateBook.Open()

rsUpdateBook_numRows = 0
%>
<%
Dim rsCategories
Dim rsCategories_numRows

Set rsCategories = Server.CreateObject("ADODB.Recordset")
rsCategories.ActiveConnection = MM_conn_webprodmx_STRING
rsCategories.Source = "SELECT CategoryID, Category  FROM dbo.tbl_Categories  ORDER BY Category"
rsCategories.CursorType = 0
rsCategories.CursorLocation = 2
rsCategories.LockType = 1
rsCategories.Open()

rsCategories_numRows = 0
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<p>Update books</p>
<form method="post" action="<%=MM_editAction%>" name="form1">
  <table align="center">
    <tr valign="baseline">
      <td nowrap align="right">Book Title:</td>
      <td>
        <input type="text" name="BookTitle" value="<%=(rsUpdateBook.Fields.Item("BookTitle").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Author First Name:</td>
      <td>
        <input type="text" name="BookAuthorFirstName" value="<%=(rsUpdateBook.Fields.Item("BookAuthorFirstName").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Author Last Name:</td>
      <td>
        <input type="text" name="BookAuthorLastName" value="<%=(rsUpdateBook.Fields.Item("BookAuthorLastName").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Price:</td>
      <td>
        <input type="text" name="BookPrice" value="<%=(rsUpdateBook.Fields.Item("BookPrice").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">ISBN:</td>
      <td>
        <input type="text" name="BookISBN" value="<%=(rsUpdateBook.Fields.Item("BookISBN").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Page Count:</td>
      <td>
        <input type="text" name="BookPageCount" value="<%=(rsUpdateBook.Fields.Item("BookPageCount").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Image Name:</td>
      <td>
        <input type="text" name="BookImage" value="<%=(rsUpdateBook.Fields.Item("BookImage").Value)%>" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Category:</td>
      <td>
        <select name="CategoryID">
          <%
While (NOT rsCategories.EOF)
%>
          <option value="<%=(rsCategories.Fields.Item("CategoryID").Value)%>" <%If (Not isNull(rsUpdateBook.Fields.Item("CategoryID").Value)) Then If (CStr(rsCategories.Fields.Item("CategoryID").Value) = CStr(rsUpdateBook.Fields.Item("CategoryID").Value)) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(rsCategories.Fields.Item("Category").Value)%></option>
          <%
  rsCategories.MoveNext()
Wend
If (rsCategories.CursorType > 0) Then
  rsCategories.MoveFirst
Else
  rsCategories.Requery
End If
%>
        </select>
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">&nbsp;</td>
      <td>
        <input type="submit" value="Update record">
      </td>
    </tr>
  </table>
  <input type="hidden" name="MM_update" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= rsUpdateBook.Fields.Item("BookID").Value %>">
</form>
<p>&nbsp;</p>
</body>
</html>
<%
rsUpdateBook.Close()
Set rsUpdateBook = Nothing
%>
<%
rsCategories.Close()
Set rsCategories = Nothing
%>
