DOCX = $(wildcard *.docx)
MD = $(DOCX:.docx=.md)

%.md: %.docx
	pandoc -gfm --ascii --extract-media=. $^ $@

all:
	echo $(MD)
