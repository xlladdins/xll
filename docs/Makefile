DOCX = $(wildcard *.docx)
MD = $(DOCX:.docx=.md)

%.md: %.docx
	pandoc -t gfm --ascii --extract-media=. $^ -o $@

all: $(MD)
