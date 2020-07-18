DOCX = $(wildcard *.docx)
MD = $(DOCX:.docx=.md)

%.md: %.docx
	pandoc -t gfm --ascii --extract-media=. $^ -o $@

all: md readme

md: $(MD)

readme: md
	(cd Excel4Macros; make)

clean:
	rm $(MD)
	rm -rf media
	(cd Excel4Macros; make clean)
