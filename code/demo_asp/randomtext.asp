<%
'randomizer plugin example
dim i, randomizer
set randomizer=aspL.plugin("randomizer")

for i=1 to 100
	'generate some random words with random lengths (10-20)
	body=body & randomizer.randomtext(randomizer.randomnumber(10,20)) & " "
next	

set randomizer=nothing
%>