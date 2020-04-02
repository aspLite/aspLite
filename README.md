# ASP-VBScript-Framework

This framework can help to develop better ASP/VBScript applications. 

Install as a website or virtual directory in IIS.

Preview: https://aspfw.quickersite.com

Highlights:
* split html/css/javascript and asp
* conditionally load asp files (rather than include them 'all')
* charset utf-8
* .net-alike codebehind approach
* light-weight and fast ASP classes
* write your own plugins

What this framework actually does is wrap useful ASP components in reusable functions and classes:

* file streaming (text + binary)
* ServerXMLHTTP & DOMDocument
* application caching
* generic ASP/VBScript functions

The following plugins are included already (but are not necessary to use the framework i.e., they can be ignored or deleted

* connector for Access databases (OLEDB)
* mailer for CDO.Message
* jpg resize/recolor (.net required)
* json read/write
* text/number randomizing
* rss feed reader


This ASP/VBScript framework does not include or favorize frontend CSS/JavaScript libraries. But just for the fun of it, 
I added a Bootstrap sample and some AJAX examples each using jQuery. 

## Roadmap

* make file (and complete folder?) uploads in ASP classic easier
* keep security high (protect against sqli, csrf, brute force attacks, cookie theft, etc)
* implement one or more WYSIWYG editors (as plugins)
* facilitate multilingual websites and applications