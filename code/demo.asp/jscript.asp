<script runat="server" language="jscript">

// this script has access to the Request() collection.
 
  var message = 'This is my message: '
  for (var i = 0, endI = 10 ; i < endI ; i++){ 
      Response.Write( message + ' ' + i + '<br>' )
  }
  
</script>