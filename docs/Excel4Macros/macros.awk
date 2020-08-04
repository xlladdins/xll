#!/usr/bin/awk -f

BEGIN {
    okey = ""
    skip = 1
    related = 0
}

/^<span id=/ { skip = 1 }
/^Return to/ { next }
/^> \&nbsp;$/ { next }
/^# Tips/ { $0 = "**Tips**" } # fix up bogus pandoc parse
/^# / {
    related = 0
    if (length($2) > 1 && $0 !~ /Syntax/) {
        print "\nReturn to [README](README.md#"key")\n" >> file".md"
        close(file".md")
        skip = 0
        file = $2
        # split on "," and store all fields instead
        # file[file] = $i
        gsub(/,/, "", file)
        key = substr(file, 1, 1)
        if (key != okey) {
            okey = key
            bull = ""
            keys[key] = "\n## "key"\n"
        }
        link = $0
        sub(/^# /, "", link)
        split(link, links, ".")
        key2 = links[1]
        if (key2 != okey2) {
            bull = ""
            okey2 = key2
            keys[key] = keys[key]"  \n"
        }
        keys[key] = keys[key]""bull"["link"]("file".md)"
        bull = " &nbsp; "
    }
}
/Related Function/ {
    related = 1
}
/&nbsp;&nbsp;&nbsp;\*\*(\*\*)?/ {
    sub(/&nbsp;&nbsp;&nbsp;\*\*(\*\*)?/, "**\\&nbsp;\\&nbsp;\\&nbsp;")
}
skip == 0 {
    if (related) {
        sub(/^[A-Z\.]+/, "[&](&.md)")
    }
    print >> file".md"
}

END {
    print "# Index of Excel 4 Macro Functions\n" > "README.md"

    n = asorti(keys, keyi)
    for (i = 1; i <= n; ++i) {
        print " ["keyi[i]"](#"keyi[i]")" >> "README.md"
    }
    for (i = 1; i <= n; ++i) {
        print keys[keyi[i]] >> "README.md"
    }
}
