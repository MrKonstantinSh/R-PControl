#include <Windows.h>

#include "resource.h"

constexpr auto WINDOW_NAME = "RPC-Inviter";

ATOM RegisterWindowClass(HINSTANCE);
BOOL InitWindowInstance(HINSTANCE, int);
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

int APIENTRY WinMain(_In_ HINSTANCE hInstance, 
	_In_opt_ HINSTANCE hPrevInstance,
	_In_ LPTSTR cmdLine,
	_In_ int cmdShowMode) 
{
	MSG msg;

    RegisterWindowClass(hInstance);

    if (!InitWindowInstance(hInstance, cmdShowMode))
        return FALSE;

    while (GetMessage(&msg, NULL, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return msg.wParam;
}

ATOM RegisterWindowClass(HINSTANCE hInstance)
{
    WNDCLASSEX windowClassEx;

    windowClassEx.cbSize = sizeof(WNDCLASSEX);
    windowClassEx.style = CS_HREDRAW | CS_VREDRAW;
    windowClassEx.lpfnWndProc = WndProc;
    windowClassEx.cbClsExtra = 0;
    windowClassEx.cbWndExtra = 0;
    windowClassEx.hInstance = hInstance;
    windowClassEx.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_WINDOW_ICON));
    windowClassEx.hCursor = LoadCursor(hInstance, IDC_ARROW);
    windowClassEx.hbrBackground = CreateSolidBrush(RGB(255, 255, 255));
    windowClassEx.lpszMenuName = 0;
    windowClassEx.lpszClassName = WINDOW_NAME;
    windowClassEx.hIconSm = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_WINDOW_ICON));

    return RegisterClassEx(&windowClassEx);
}

BOOL InitWindowInstance(HINSTANCE hInstance, int cmdShowMode)
{
    HWND hWnd = CreateWindow(WINDOW_NAME,
        WINDOW_NAME,
        WS_OVERLAPPEDWINDOW, 
        CW_USEDEFAULT,
        CW_USEDEFAULT, 
        CW_USEDEFAULT, 
        CW_USEDEFAULT, 
        NULL, 
        NULL, 
        hInstance,
        NULL);

    if (!hWnd)
        return FALSE;

    ShowWindow(hWnd, cmdShowMode);

    return TRUE;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }

    return 0;
}