// tstring.h

#ifndef _INC_TSTRING
#define _INC_TSTRING

#include <string>
#include <tchar.h>

typedef std::basic_string<TCHAR, std::char_traits<TCHAR>, std::allocator<TCHAR> > tstring;

#endif // _INC_TSTRING