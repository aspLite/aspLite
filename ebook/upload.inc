<%
dim form : set form=aspl.form
form.onSubmit=""
form.write(aspl.loadText("ebook/uploadhtml.inc"))
form.writeJS(aspl.loadText("ebook/uploadjs.inc"))
form.build
%>