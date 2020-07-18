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
/^> \&nbsp;$/ { next }
/^# / {
    related = 0
    if (length($2) > 1) {
        if ($2 == "Tips") { # bogus pandoc parse
            print "**Tips**" >> file".md"
            next
        }
        else {
            print "\nReturn to [README](README.md)\n" >> file".md"
            skip = 0
            file = $2
            gsub(/,/, "", file)
            key = substr(file, 1, 1)
            if (key != okey) {
                print "\n## "key"\n" >> "README.md"
                okey = key
                bull = ""
            }
            sub(/^# /, "")
            print bull"["$0"]("file".md)" >> "README.md"
            bull = " &bull; "
        }
    }
}
/Related Function/ {
    related = 1
    print >> file".md"
    next
}
skip == 0 {
    sub("&nbsp;&nbsp;&nbsp;**", "**&nbsp;&nbsp;&nbsp;")
    if (related) {
        sub(/^[A-Z\.]+/, "[&](&.md)")
    }
    print >> file".md"
}
# TODO - fix up tables
