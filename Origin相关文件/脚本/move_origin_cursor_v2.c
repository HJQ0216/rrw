#include <Origin.h>
#include "move_cursor_api.h"

int move_origin_cursor_v2()
{
    // Move inside the Origin window, but below/right of the graph page, so
    // Origin receives a mouse-move event and clears its stale rotation cursor.
    return SetCursorPos(1350, 900);
}

int hide_origin_cursor()
{
    return ShowCursor(FALSE);
}

int show_origin_cursor()
{
    return ShowCursor(TRUE);
}
