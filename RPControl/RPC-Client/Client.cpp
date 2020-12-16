#include <Windows.h>
#include <RdpEncomAPI.h>

#include "resource.h"

#define override

constexpr auto WINDOW_NAME = "RPC-Inviter";
constexpr auto WINDOW_WIDTH = 1280;
constexpr auto WINDOW_HEIGHT = 720;

constexpr auto MAX_NUM_ATTENDEES = 1;

HWND _mainWindow;
HWND _logTextBox;
HWND _connectBtn;
HWND _disconnectBtn;
HWND _accessLevelCombobox;

char* accessLevelList[3] = { (char*)"CTRL_NONE", (char*)"CTRL_VIEW", (char*)"CTRL_INTERACTIVE" };

ATOM RegisterWindowClass(HINSTANCE);
BOOL InitWindowInstance(HINSTANCE, int);
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

void PrintTextToLog(HWND log, LPCTSTR text);

RECT GetCenterWindow(HWND parentWindow, int windowWidth, int windowHeight)
{
    RECT rect;

    GetClientRect(parentWindow, &rect);
    rect.left = (rect.right / 2) - (windowWidth / 2);
    rect.top = (rect.bottom / 2) - (windowHeight / 2);

    return rect;
}

void RenderLabel(HWND hWnd, LPCTSTR text, int x, int y, int width, int height)
{
    CreateWindow("static", text,
        WS_VISIBLE | WS_CHILD | SS_LEFT,
        x, y, width, height,
        hWnd, NULL, NULL, NULL);
}

HWND RenderTextBox(HWND hWnd, LPCTSTR text, BOOL isReadOnly, int x, int y, int width, int height)
{
    HWND log = CreateWindow("edit", text,
        WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_AUTOVSCROLL | ES_MULTILINE,
        x, y, width, height,
        hWnd, NULL, NULL, NULL);

    if (isReadOnly)
        SendMessage(log, EM_SETREADONLY, TRUE, NULL);

    return log;
}

HWND RenderButton(HWND hWnd, LPCSTR text, int x, int y, int width, int height) {
    return CreateWindow("button", text,
        WS_CHILD | WS_VISIBLE | BS_TEXT,
        x, y, width, height,
        hWnd, NULL, NULL, NULL);
}

HWND RenderCombobox(HWND hWnd, int x, int y, int width, int height)
{
    HWND combobox = CreateWindow("combobox", NULL,
        WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | CBS_HASSTRINGS,
        x, y, width, height,
        hWnd, NULL, NULL, NULL);

    for (int i = 0; i < 3; i++)
    {
        SendMessage(combobox, CB_ADDSTRING, NULL, (LPARAM)accessLevelList[i]);
    }

   SendMessage(combobox, CB_SETCURSEL, 1, 0);
   SendMessage(combobox, EM_SETREADONLY, TRUE, NULL);

    return combobox;
}

void PrintTextToLog(HWND log, LPCTSTR text)
{
    int len = GetWindowTextLength(log);
    SendMessage(log, EM_SETSEL, (WPARAM)len, (LPARAM)len);
    SendMessage(log, EM_REPLACESEL, 0, (LPARAM)text);
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

    _mainWindow = hWnd;

    if (!hWnd)
        return FALSE;

    ShowWindow(hWnd, cmdShowMode);

    return TRUE;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_CREATE:
        _connectBtn = RenderButton(hWnd, "Connect", 10, 5, 177, 30);
        _disconnectBtn = RenderButton(hWnd, "Disconnect", 197, 5, 177, 30);

        RenderLabel(hWnd, "Access level:", 10, 45, 90, 18);
        _accessLevelCombobox = RenderCombobox(hWnd, 10, 67, 365, 75);

        RenderLabel(hWnd, "Log:", 10, 97, 30, 18);
        _logTextBox = RenderTextBox(hWnd, NULL, TRUE, 10, 119, 365, 550);

        EnableWindow(_disconnectBtn, FALSE);
        EnableWindow(_accessLevelCombobox, FALSE);
        break;
    case WM_COMMAND:
        if ((HWND)lParam == _connectBtn)
        {
            EnableWindow(_connectBtn, FALSE);
            EnableWindow(_disconnectBtn, TRUE);
            EnableWindow(_accessLevelCombobox, TRUE);
        }
        if ((HWND)lParam == _disconnectBtn)
        {
            EnableWindow(_connectBtn, TRUE);
            EnableWindow(_disconnectBtn, FALSE);
            EnableWindow(_accessLevelCombobox, FALSE);
            SendMessage(_accessLevelCombobox, CB_SETCURSEL, 1, 0);
        }
        break;
    case WM_CTLCOLORSTATIC:
        return (INT_PTR)CreateSolidBrush(RGB(255, 255, 255));
    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }

    return 0;
}