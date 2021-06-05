#TYPE = s5
TYPE = slidy
#TYPE = slideous 
#TYPE = dzslides 

talk:
	pandoc -t $(TYPE) -s talk.md -o talk.html
