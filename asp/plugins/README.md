# How to write plugins?

Plugins are stored in the /asp/plugins folder.

The filename of your plugin (eg "helloworld.asp") also names the plugin: "helloworld" (without .asp)

Finally, make sure to name your class cls_asp_helloworld. 

To create an instance of your plugin, use: 

dim myplugin<br>
set myplugin=asp.plugin("helloworld")<br>
response.write myplugin.hw()<br>

from there on you can use all public properties, subs and functions of your plugin

You could also just:

response.write asp.plugin("helloworld").hw()
