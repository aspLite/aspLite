<%
dim rss : set rss=aspL.plugin("rss")
rss.maxitems=1

dim form : set form=new cls_formbuilder
form.targetDiv="sampleform12"
form.id="cnn"
	
'result
dim feedback : set feedback=form.field
feedback.add "type","element"
feedback.add "html",rss.read("http://rss.cnn.com/rss/cnn_topstories.rss")
feedback.add "tag","div"

form.build()

%>