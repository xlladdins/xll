#!/usr/bin/awk -f

BEGIN {
    FS = "[ \t]+"
    okey = ""
    skip = 1
    related = 0
	print "# Index of Excel 4 Macro Functions" > "README.md"
	print >> "README.md"
}

/^Return to/ { next }
/^# / {
    related = 0
    if (length($2) > 1) {
        skip = 0
        file = $2
        gsub(/,/, "", file)
        key = substr(file, 1, 1)
        if (key != okey) {
            print "\n## "key"\n" >> "README.md"
            okey = key
        }
        sub(/^# /, "")
		print "["$0"]("file".md)" >> "README.md"
    }
}
/Related Function/ {
    related = 1
    print >> file".md"
    next
}
skip == 0 {
    if (related) {
        sub(/^[A-Z\.]+/, "[&](&.md)")
    }
    print >> file".md"
}
# TODO - fix up tables
