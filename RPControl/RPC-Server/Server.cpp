#include <Windows.h>

#include "resource.h"

constexpr auto WINDOW_NAME = "RPC-Inviter";
constexpr auto WINDOW_WIDTH = 400;
constexpr auto WINDOW_HEIGHT = 400;

ATOM RegisterWindowClass(HINSTANCE);
BOOL InitWindowInstance(HINSTANCE, int);
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

RECT GetCenterWindow(HWND parentWindow, int windowWidth, int windowHeight)
{
    RECT rect;

    GetClientRect(parentWindow, &rect);
    rect.left = (rect.right / 2) - (windowWidth / 2);
    rect.top = (rect.bottom / 2) - (windowHeight / 2);

    return rect;
}

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
    RECT windowPosition = GetCenterWindow(GetDesktopWindow(), WINDOW_WIDTH, WINDOW_HEIGHT);

    HWND hWnd = CreateWindow(WINDOW_NAME,
        WINDOW_NAME,
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
        windowPosition.left,
        windowPosition.top, 
        WINDOW_WIDTH,
        WINDOW_HEIGHT,
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