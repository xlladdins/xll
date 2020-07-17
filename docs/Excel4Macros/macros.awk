#!/usr/bin/awk -f

BEGIN {
    okey = ""
    skip = true
	print "# Index of Excel 4 Macro Functions" > "README.md"
	print >> "README.md"
}

/^Return to/ { next }
/^# / {
    if (length($2) > 1) {
        skip = false
print "file = "file
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
skip == false {
    print >> file".md"
}
