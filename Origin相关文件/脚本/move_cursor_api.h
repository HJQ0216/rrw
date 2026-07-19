#pragma dll(user32, system)

BOOL __stdcall SetCursorPos(int x, int y);
int __stdcall ShowCursor(BOOL bShow);
