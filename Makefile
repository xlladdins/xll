#TYPE = s5
#TYPE = slidy
#TYPE = slideous 
#TYPE = dzslides 
TYPE = html5

%.html: %.md
	pandoc -t $(TYPE) -s talk.md -o talk.html

%.pdf: %.md
	pandoc talk.md -o talk.pdf
