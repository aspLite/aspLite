# How to write plugins?

Plugins are stored in the /asplite/plugins folder.

1. Add a seperate folder for your plugin: /asplite/plugins/myplugin/
2. Add an asp file with SAME name: /asplite/plugins/myplugin/myplugin.asp
3. In myplugin.asp, create a class:

	class cls_asplite_myplugin
		...
	end class

4. To create an instance of your plugin, use: 

  dim myObj<br>
  set myObj=asp.plugin("myplugin")<br>
  response.write myObj.XX()<br>

  You could also just:

  response.write asp.plugin("myplugin").XX()

  (where XX() is a public function of your myplugin-class)