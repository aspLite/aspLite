# How to write plugins?

Plugins are stored in the /asp/plugins folder.

The filename of your plugin (eg "helloworld.asp") also names the plugin: "helloworld" (without .asp)

Finally, make sure to name your class cls_asp_helloworld. 

To create an instance of your plugin, use: 

dim myplugin<br>
set myplugin=asp.plugin("helloworld")<br>
response.write myplugin.hw()<br>

You could also just:

response.write asp.plugin("helloworld").hw()

(where hw() is a public function of the helloworld-class)