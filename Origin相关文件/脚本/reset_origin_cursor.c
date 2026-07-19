#include <Origin.h>

// Force Origin to restore the standard arrow cursor before graph export.
void reset_origin_cursor()
{
    waitCursor wc;
    SetCursorPos(0, 0);
}
