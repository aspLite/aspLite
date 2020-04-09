<%@ Language= "Javascript" codepage="65001" %> 
<%

// this script has access to the Request() collection.
 
  var message = 'This is my message: '
  for (var i = 0, endI = 10 ; i < endI ; i++){ 
      Response.Write( message + ' ' + i + '<br>' )
  }
 
%>