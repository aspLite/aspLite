<%
'randomizer plugin example
dim i, randomizer
set randomizer=aspL.plugin("randomizer")

body="ASP randomizer-plugin: "
for i=1 to 25
	'generate some random words with random lengths (5-10)
	body=body & randomizer.randomtext(randomizer.randomnumber(5,10)) & " "
next	

set randomizer=nothing

%>