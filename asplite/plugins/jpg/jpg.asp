<%
class cls_asplite_jpg

	public maxsize,path,effect,square
	
	'square=0/1 - crop square: 0:no / 1:yes
	'effect=1/2/3 - 1:bw / 2:grayscale / 3:sepia

	private sub class_initialize
	
		maxsize=1920 'max= 2560		
		
	end sub
	
	public function src
	
		src=aspL_path & "/plugins/jpg/jpg.aspx?img=" & server.urlencode(path) & "&maxsize="& maxsize & "&se=" & effect & "&square=" & square
	
	end function

end class 
%>