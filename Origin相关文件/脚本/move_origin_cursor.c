#include <Origin.h>

// Move the system pointer away from the graph page before export.
int move_origin_cursor()
{
    return SetCursorPos(0, 0);
}
