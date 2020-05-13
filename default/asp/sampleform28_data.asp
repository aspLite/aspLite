<%
'Vue
set rs=db.execute("select top 5 * from testdata")
aspl.json.dump(rs)
%>