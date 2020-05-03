<%
dim rss : set rss=aspL.plugin("rss")
rss.maxitems=1

'formbuilder sample - built on bootstrap css
aspl.exec("code/demo_asp/formbuilder/formbuilder.asp")

dim form : set form=new cls_formbuilder
form.listenTo "e","sampleform12"
form.targetDiv="sampleform12"
form.id="cnn"
	
'result
dim feedback : set feedback=form.field
feedback.add "type","comment"
feedback.add "html",rss.read("http://rss.cnn.com/rss/cnn_topstories.rss")
feedback.add "tag","div"

form.build()

%>