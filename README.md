<div align="center">

## ADO GetRows example


</div>

### Description

You can use ADO GetRows to output and ADO recordset to an array. This is often useful in n-tier applications when you are moving data between tiers--or if you want to persist your data in another way.

http://adozone.cnw.com/default.htm
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Found on the World Wide Web](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/found-on-the-world-wide-web.md)
**Level**          |Beginner
**User Rating**    |4.2 (21 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/found-on-the-world-wide-web-ado-getrows-example__4-21/archive/master.zip)





### Source Code

```
<HTML>
<HEAD>
<TITLE>Place Document Title Here</TITLE>
</HEAD>
<BODY BGColor=ffffff Text=000000>
<%
Set cn = Server.CreateObject("ADODB.Connection")
cn.Open Application("guestDSN")
sql = "SELECT * FROM authors"
Set RS = cn.Execute(sql)
ary = rs.GetRows(10)
rs.close
cn.close
%>
<P>
<TABLE BORDER=1>
<%
nRows = UBound( ary, 2 )
For row = 0 to nRows %>
<TR>
<% For col = 0 to UBound( ary, 1 ) %>
<TD><%= ary( col, row ) %> </TD>
<% Next %>
</TR>
<% Next %>
</TABLE>
</HTML>
```

