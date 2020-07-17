#!/usr/bin/awk -f

BEGIN {
	print "# Index of Excel 4 Macro Functions" > "README.md"
	print >> "README.md"
}


/^Return to/ { skip }
/==/ {
	key = substr(file,1,1)
	if (okey != key) {
		print "\n## "key"\n" >> "README.md"
		okey = key
	}
	if (file) {
		#print "Previous ["file"]("file".md)  "
		print "close "file".md"
		close(file".md")
		print "["link"]("file".md)" >> "README.md"
	}
	file = prev
	link = name
	print file > file".md"
	print $0 >> file".md"
}
!/==/ {
	name = $0
	prev = $1
	gsub(/,/, "", prev)
	print >> file".md"
 }
