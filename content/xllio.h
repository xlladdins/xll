#include "oper.h"
#include "excel.h"

namespace xll {

    // xlfFwrite only writes 255 chars at a time.
    inline OPER fwrite(const OPER& fd, const OPER& text)
    {
        OPER start{ 1 };
        OPER len = Excel(xlfLen, text);

        if (len.xltype != xltypeNum) {
            return xll::ErrNA;
        }

        while (len.val.num > 255) {
            const OPER& mid = Excel(xlfMid, text, start, OPER(255));
            auto result = Excel(xlfFwrite, fd, mid);
            ensure(result == 255);
            start.val.num += 255;
            len.val.num -= 255;
        }

        return Excel(xlfFwrite, fd, Excel(xlfMid, text, start, len));
    }

    // Read one line at a time.
    inline OPER fread(const OPER& fd)
    {
        ensure(fd.xltype == xltypeNum);

        OPER text;
        while (OPER line = Excel(xlfFreadln, fd)) {
            text = text & line;
        }

        return text;
    }

    class xlfFile {
        OPER h; // file handle
    public:
        // 1 = read/write 2 = readonly 3 = create read/write
        xlfFile(const OPER& file, int access = 3)
            : h(Excel(xlfFopen, file, OPER(access)))
        { }
        xlfFile(const xlfFile&) = delete;
        xlfFile& operator=(const xlfFile&) = delete;
        ~xlfFile()
        {
            Excel(xlfFclose, h);
        }
        operator bool() const
        {
            return h.xltype == xltypeNum;
        }
        OPER read() const
        {
            return fread(h);
        }
        OPER write(const OPER& text) const
        {
            return fwrite(h, text);
        }
        OPER writeln(const OPER& text) const
        {
            ensure(Excel(xlfLen, text).val.num < 256);

            return Excel(xlfFwriteln, h, text);
        }
    };

    template<class X>
    inline XOPER<X> find_last_of(const XOPER<X>& find, const XOPER<X>& within)
    {
        XOPER<X> off{ 0. };

        while (auto next = XExcel<X>(xlfFind, find, within, XOPER<X>(off.val.num + 1))) {
            off = next;
        }

        return off;
    }

    template<class X>
    inline XOPER<X> basename(const XOPER<X>& path, bool strip = false)
    {
        XOPER<X> off = find_last_of(XOPER<X>(TEXT("\\")), path);
        XOPER<X> base = XExcel<X>(xlfRight, path, XOPER<X>(Excel(xlfLen, path).val.num - off.val.num));

        if (strip) {
            base = XExcel<X>(xlfLeft, base, XOPER<X>(find_last_of(XOPER<X>(L"."), base).val.num - 1));
        }

        return base;
    }

    template<class X>
    inline XOPER<X> dirname(const XOPER<X>& path)
    {
        XOPER<X> off = find_last_of(XOPER<X>(TEXT("\\")), path);

        return XExcel<X>(xlfLeft, path, off);
    }

}
