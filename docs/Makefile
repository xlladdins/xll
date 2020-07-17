DOCX = $(wildcard *.docx)
MD = $(DOCX:*.docx=.md)

%.md: %.docx
	pandoc -gfm --extract-media=. $< $>

all: $(MD)
