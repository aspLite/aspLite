<%
on error resume next

dim template
template=aspL.loadText("html/demo_asp/hash.resx")

dim password
password=aspl.plugin("randomizer").randomText(30)		

dim md5
md5=aspL.plugin("md5").md5(password,32)

dim sha256
sha256=aspL.plugin("sha256").sha256(password)

template=replace(template,"[pw]",password,1,-1,1)
template=replace(template,"[md5]",md5,1,-1,1)
template=replace(template,"[sha256]",sha256,1,-1,1)

aspL.aspError("hashing")
		
json.dump(template)
%>