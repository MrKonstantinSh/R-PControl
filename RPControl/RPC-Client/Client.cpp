#include <Windows.h>
#include <fstream>
#include <comdef.h>
#include <RdpEncomAPI.h>

#include "resource.h"
#include "AX.h"

using namespace std;

#define override

constexpr auto WINDOW_NAME = "RPC-Client";
constexpr auto WINDOW_WIDTH = 1280;
constexpr auto WINDOW_HEIGHT = 720;

constexpr auto MAX_NUM_ATTENDEES = 1;

HWND h_mainWindow;
HWND h_logTextBox;
HWND h_connectBtn;
HWND h_disconnectBtn;
HWND h_accessLevelCombobox;
HWND h_streamWindowWrapper;
HWND h_streamWindow;

IRDPSRAPIViewer* rdp_viewer = NULL;
IRDPSRAPIInvitationManager* rdp_invitationManager = NULL;
IRDPSRAPIInvitation* rdp_invitation = NULL;
IRDPSRAPIAttendeeManager* rdp_attendeeManager = NULL;
IRDPSRAPIAttendee* rdp_attendee = NULL;

IConnectionPointContainer* conn_picpc = NULL;
IConnectionPoint* conn_picp = NULL;

int _connectionId = NULL;
char* _accessLevelList[3] = { (char*)"DO NOT SHOW", (char*)"VIEW SCREEN", (char*)"PC CONTROL" };

ATOM RegisterWindowClass(HINSTANCE);
BOOL InitWindowInstance(HINSTANCE, int);
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

void PrintTextToLog(HWND, LPCTSTR);

void OnConnectionFailed();
void OnConnectionEstablished();
int ConnectEvent(IUnknown*, REFIID, IUnknown*,
    IConnectionPointContainer**, IConnectionPoint**);
void DisconnectEvent(IConnectionPointContainer*, IConnectionPoint*, unsigned int);
BOOL TryConnect();
void TryDisconnect();

class RDPSessionEvents : public _IRDPSessionEvents
{
public:
    virtual HRESULT STDMETHODCALLTYPE override QueryInterface(
        REFIID iid,
        void** ppvObject)
    {
        *ppvObject = 0;

        if (iid == IID_IUnknown || iid == IID_IDispatch || iid == __uuidof(_IRDPSessionEvents))
            *ppvObject = this;

        if (*ppvObject)
        {
            ((IUnknown*)(*ppvObject))->AddRef();

            return S_OK;
        }

        return E_NOINTERFACE;
    }

    virtual ULONG STDMETHODCALLTYPE override AddRef(void)
    {
        return 0;
    }

    virtual ULONG STDMETHODCALLTYPE override Release(void)
    {
        return 0;
    }

    virtual HRESULT STDMETHODCALLTYPE override GetTypeInfoCount(
        __RPC__out UINT* pctinfo)
    {
        return E_NOTIMPL;
    }

    virtual HRESULT STDMETHODCALLTYPE override GetTypeInfo(
        UINT iTInfo,
        LCID lcid,
        __RPC__deref_out_opt ITypeInfo** ppTInfo)
    {
        return E_NOTIMPL;
    }

    virtual HRESULT STDMETHODCALLTYPE override GetIDsOfNames(
        __RPC__in REFIID riid,
        __RPC__in_ecount_full(cNames) LPOLESTR* rgszNames,
        UINT cNames,
        LCID lcid,
        __RPC__out_ecount_full(cNames) DISPID* rgDispId)
    {
        return E_NOTIMPL;
    }

    virtual HRESULT STDMETHODCALLTYPE override Invoke(
        DISPID dispIdMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS FAR* pDispParams,
        VARIANT FAR* pVarResult,
        EXCEPINFO FAR* pExcepInfo,
        unsigned int FAR* puArgErr)
    {
        switch (dispIdMember) 
        {
        case DISPID_RDPSRAPI_EVENT_ON_VIEWER_CONNECTFAILED:
            OnConnectionFailed();
            break;
        case DISPID_RDPSRAPI_EVENT_ON_VIEWER_CONNECTED:
            OnConnectionEstablished();
            break;
        }

        return S_OK;
    }
};

RDPSessionEvents rdpEvents;

void OnConnectionFailed()
{
    PrintTextToLog(h_logTextBox, "Connection failed!\r\n");
    TryDisconnect();
}

void OnConnectionEstablished() 
{
    PrintTextToLog(h_logTextBox, "Connection successful!\r\n\r\n");
}

int ConnectEvent(IUnknown* Container, REFIID riid, IUnknown* Advisor,
    IConnectionPointContainer** picpc, IConnectionPoint** picp)
{
    HRESULT hr = 0;
    unsigned long tid = 0;
    IConnectionPointContainer* icpc = 0;
    IConnectionPoint* icp = 0;
    *picpc = 0;
    *picp = 0;

    Container->QueryInterface(IID_IConnectionPointContainer, (void**)&icpc);

    if (icpc)
    {
        *picpc = icpc;
        icpc->FindConnectionPoint(riid, &icp);
        if (icp)
        {
            *picp = icp;
            hr = icp->Advise(Advisor, &tid);
        }
    }

    return tid;
}

void DisconnectEvent(IConnectionPointContainer* icpc, IConnectionPoint* icp, unsigned int connectionId)
{
    icp->Unadvise(connectionId);
    icp->Release();
    icpc->Release();
}

char* SelectInvitationFile() 
{
    OPENFILENAME openFileName;

    char* fileName = new char[100];

    ZeroMemory(&openFileName, sizeof(openFileName));

    openFileName.lStructSize = sizeof(openFileName);
    openFileName.hwndOwner = h_mainWindow;
    openFileName.lpstrFile = fileName;
    openFileName.lpstrFile[0] = 0;
    openFileName.nMaxFile = sizeof(openFileName);
    openFileName.lpstrFilter = "XML\0*.xml\0";
    openFileName.nFilterIndex = 1;
    openFileName.lpstrFileTitle = 0;
    openFileName.nMaxFileTitle = 0;
    openFileName.lpstrInitialDir = 0;
    openFileName.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

    if (GetOpenFileName(&openFileName))
        return fileName;

    return NULL;
}

BOOL IsInvitationFileExist(char* fileName)
{
    ifstream file;
    file.open(fileName);

    if (!file)
        return FALSE;

    return TRUE;
}

BOOL TryConnect() 
{
    char* invitationFile = SelectInvitationFile();

    if (invitationFile) 
    {
        if (IsInvitationFileExist(invitationFile))
        {
            PrintTextToLog(h_logTextBox, "The existence of the invitation file has been verified!\r\n");

            if (rdp_viewer == NULL) 
            {
                HRESULT resOfCreatingInstance = CoCreateInstance(
                    __uuidof(RDPViewer),
                    NULL,    
                    CLSCTX_INPROC_SERVER,
                    __uuidof(IRDPSRAPIViewer),
                    (void**)&rdp_viewer);

                if (resOfCreatingInstance == S_OK)
                {
                    PrintTextToLog(h_logTextBox, "Instance created!\r\n");

                    _connectionId = ConnectEvent((IUnknown*)rdp_viewer, __uuidof(_IRDPSessionEvents),
                        (IUnknown*)&rdpEvents, &conn_picpc, &conn_picp);

                    PrintTextToLog(h_logTextBox, "Reading invitation file!\r\n");
                    
                    ifstream fileStream(invitationFile);

                    if (fileStream.is_open())
                    {
                        char invitationString[2000];
                        ZeroMemory(invitationString, sizeof(invitationString));
                       
                        fileStream.getline(invitationString, 2000);
                        
                        fileStream.close();

                        if (rdp_viewer->Connect(_bstr_t(invitationString), SysAllocString(L"R-PControl"), SysAllocString(L"")) == S_OK)
                        {
                            PrintTextToLog(h_logTextBox, "Connection line active!\r\n");

                            return TRUE;
                        }
                        else {
                            PrintTextToLog(h_logTextBox, "Connection line error!\r\n");

                            return FALSE;
                        }
                    }
                    else 
                    {
                        PrintTextToLog(h_logTextBox, "Error reading invitation!\r\n");

                        return FALSE;
                    }
                }
                else 
                {
                    PrintTextToLog(h_logTextBox, "Error creating instance!\r\n");

                    return FALSE;
                }
            }
            else 
            {
                PrintTextToLog(h_logTextBox, "Error starting: Session already exists!\r\n");

                return FALSE;
            }
        }
        else 
        {
            PrintTextToLog(h_logTextBox, "Invalid invitation file!\r\n");

            return FALSE;
        }
    }
    else 
    {
        PrintTextToLog(h_logTextBox, "Error: An invitation file must be selected!\r\n");

        return FALSE;
    }
}

void TryDisconnect() 
{
    PrintTextToLog(h_logTextBox, "\r\nDisconnecting...\r\n");

    if (rdp_viewer) 
    {
        DisconnectEvent(conn_picpc, conn_picp, _connectionId);
        rdp_viewer->Disconnect();
        rdp_viewer->Release();
        rdp_viewer = NULL;

        PrintTextToLog(h_logTextBox, "Disconnected!\r\n");
    }
    else
        PrintTextToLog(h_logTextBox, "Error disconnecting: No active connection!\r\n");
}

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
        SendMessage(combobox, CB_ADDSTRING, NULL, (LPARAM)_accessLevelList[i]);
    }

   SendMessage(combobox, CB_SETCURSEL, 1, 0);
   SendMessage(combobox, EM_SETREADONLY, TRUE, NULL);

    return combobox;
}

HWND RenderStreamWindowWrapper(HWND hWnd, int x, int y, int width, int height)
{
    HWND view = CreateWindow("edit", NULL,
        WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN,
        x, y, width, height,
        hWnd, NULL, NULL, NULL);

    SendMessage(view, EM_SETREADONLY, TRUE, NULL);

    return view;
}

HWND RenderStreamWindow(HWND hWnd) {

    HWND streamWindow = CreateWindow("AX", "}32be5ed2-5c86-480f-a914-0ff8885a1b3f}", WS_CHILD | WS_VISIBLE,
        0, 0, 1, 1,
        hWnd, NULL, NULL, NULL);

    SendMessage(streamWindow, AX_RECREATE, 0, (LPARAM)rdp_viewer);
    SendMessage(streamWindow, AX_INPLACE, 1, 0);

    ShowWindow(streamWindow, SW_MAXIMIZE);

    return streamWindow;
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

    AXRegister();

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

    h_mainWindow = hWnd;

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
    {
        h_connectBtn = RenderButton(hWnd, "Connect", 10, 5, 177, 30);
        h_disconnectBtn = RenderButton(hWnd, "Disconnect", 197, 5, 177, 30);

        RenderLabel(hWnd, "Access level:", 10, 45, 90, 18);
        h_accessLevelCombobox = RenderCombobox(hWnd, 10, 67, 365, 75);

        RenderLabel(hWnd, "Log:", 10, 97, 30, 18);
        h_logTextBox = RenderTextBox(hWnd, NULL, TRUE, 10, 119, 365, 550);

        RECT windowInfo;
        GetClientRect(hWnd, &windowInfo);

        int width = windowInfo.right;
        int height = windowInfo.bottom;
        h_streamWindowWrapper = RenderStreamWindowWrapper(hWnd, 400, 5, width - 410, height - 15);

        EnableWindow(h_disconnectBtn, FALSE);
        EnableWindow(h_accessLevelCombobox, FALSE);

        break;
    }
    case WM_COMMAND:
        if ((HWND)lParam == h_connectBtn)
        {
            EnableWindow(h_connectBtn, FALSE);
            EnableWindow(h_disconnectBtn, TRUE);
            EnableWindow(h_accessLevelCombobox, TRUE);

            if (TryConnect())
            {
                h_streamWindow = RenderStreamWindow(h_streamWindowWrapper);
            }
        }
        if ((HWND)lParam == h_disconnectBtn)
        {
            TryDisconnect();

            EnableWindow(h_connectBtn, TRUE);
            EnableWindow(h_disconnectBtn, FALSE);
            EnableWindow(h_accessLevelCombobox, FALSE);
            SendMessage(h_accessLevelCombobox, CB_SETCURSEL, 1, 0);
        }
        if ((HWND)lParam == h_accessLevelCombobox) 
        {
            if (HIWORD(wParam) == CBN_SELCHANGE)
            {
                int index = SendMessage(h_accessLevelCombobox, CB_GETCURSEL, 0, 0);

                switch (index) 
                {
                case 0:
                    if (rdp_viewer->RequestControl(CTRL_LEVEL_NONE) == S_OK) 
                    {
                        PrintTextToLog(h_logTextBox, "Access level changed to \"DO NOT SHOW\"!\r\n");
                        ShowWindow(h_streamWindow, SW_HIDE);
                    }
                    else
                        PrintTextToLog(h_logTextBox, "Error requesting access level!\r\n");
                    break;
                case 1:
                    if (rdp_viewer->RequestControl(CTRL_LEVEL_VIEW) == S_OK)
                    {
                        PrintTextToLog(h_logTextBox, "Access level changed to \"VIEW SCREEN\"!\r\n");
                        ShowWindow(h_streamWindow, SW_MAXIMIZE);
                    }
                    else
                        PrintTextToLog(h_logTextBox, "Error requesting access level!\r\n");
                    break;
                case 2:
                    if (rdp_viewer->RequestControl(CTRL_LEVEL_INTERACTIVE) == S_OK)
                    {
                        PrintTextToLog(h_logTextBox, "Access level changed to \"PC CONTROL\"!\r\n");
                        ShowWindow(h_streamWindow, SW_MAXIMIZE);
                    }
                    else
                        PrintTextToLog(h_logTextBox, "Error requesting access level!\r\n");
                    break;
                }
            }
        }
        break;
    case WM_CTLCOLORSTATIC:
        return (INT_PTR)CreateSolidBrush(RGB(255, 255, 255));
    case WM_CLOSE:
        TryDisconnect();
   
        DestroyWindow(h_streamWindowWrapper);
        DestroyWindow(hWnd);
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }

    return 0;
}