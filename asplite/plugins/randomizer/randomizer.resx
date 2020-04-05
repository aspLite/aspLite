<%
class cls_asplite_randomizer

	Private Sub Class_Initialize()
	
		randomize()			
	
	end sub

	public function randomText(nmbrChars)	

		dim i
		for i=1 to nmbrChars	
			randomText=randomText & CHR(Int((122-97+1)*Rnd+97))
		next
	   
	End Function

	public function randomNumber(startnr,stopnr)

		randomNumber=Int((stopnr-startnr+1)*Rnd+startnr)
		
	end function

end class
%>