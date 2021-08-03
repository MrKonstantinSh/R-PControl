#include <Windows.h>
#include <fstream>
#include <comdef.h>
#include <RdpEncomAPI.h>

#include "resource.h"

using namespace std;

#define override

constexpr auto WINDOW_NAME = "RPC-Inviter";
constexpr auto WINDOW_WIDTH = 400;
constexpr auto WINDOW_HEIGHT = 400;

constexpr auto MAX_NUM_ATTENDEES = 1;

HWND h_mainWindow;
HWND h_logTextBox;
HWND h_startSharingBtn;
HWND h_stopSharingBtn;

IRDPSRAPISharingSession* rdp_session = NULL;
IRDPSRAPIInvitationManager* rdp_invitationManager = NULL;
IRDPSRAPIInvitation* rdp_invitation = NULL;
IRDPSRAPIAttendeeManager* rdp_attendeeManager = NULL;
IRDPSRAPIAttendee* rdp_attendee = NULL;

IConnectionPointContainer* conn_picpc = NULL;
IConnectionPoint* conn_picp = NULL;

ATOM RegisterWindowClass(HINSTANCE);
BOOL InitWindowInstance(HINSTANCE, int);
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

void PrintTextToLog(HWND log, LPCTSTR text);

void OnAttendeeConnected(IDispatch* pAttendee);
void OnAttendeeDisconnected(IDispatch* pAttendee);
void OnControlLevelChangeRequest(IDispatch* pAttendee, CTRL_LEVEL RequestedLevel);

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

    // IDispatch
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
        case DISPID_RDPSRAPI_EVENT_ON_ATTENDEE_CONNECTED:
            OnAttendeeConnected(pDispParams->rgvarg[0].pdispVal);
            break;
        case DISPID_RDPSRAPI_EVENT_ON_ATTENDEE_DISCONNECTED:
            OnAttendeeDisconnected(pDispParams->rgvarg[0].pdispVal);
            break;
        case DISPID_RDPSRAPI_EVENT_ON_CTRLLEVEL_CHANGE_REQUEST:
            OnControlLevelChangeRequest(pDispParams->rgvarg[1].pdispVal, 
                (CTRL_LEVEL)pDispParams->rgvarg[0].intVal);
            break;
        }

        return S_OK;
    }
};

RDPSessionEvents rdpEvents;

void OnAttendeeConnected(IDispatch* pAttendee) 
{
    IRDPSRAPIAttendee* pRDPAtendee;

    pAttendee->QueryInterface(__uuidof(IRDPSRAPIAttendee), (void**)&pRDPAtendee);
    pRDPAtendee->put_ControlLevel(CTRL_LEVEL::CTRL_LEVEL_VIEW);
    pAttendee->Release();

    PrintTextToLog(h_logTextBox, "An attendee connected!\r\n");
}

void OnAttendeeDisconnected(IDispatch* pAttendee)
{
    IRDPSRAPIAttendeeDisconnectInfo* info;
    ATTENDEE_DISCONNECT_REASON reason;

    pAttendee->QueryInterface(__uuidof(IRDPSRAPIAttendeeDisconnectInfo), (void**)&info);

    if (info->get_Reason(&reason) == S_OK)
    {
        LPCTSTR textReason;

        switch (reason) 
        {
        case ATTENDEE_DISCONNECT_REASON_APP:
            textReason = "Viewer terminated session!\r\n";
            break;
        case ATTENDEE_DISCONNECT_REASON_ERR:
            textReason = "Internal Error!\r\n";
            break;
        case ATTENDEE_DISCONNECT_REASON_CLI:
            textReason = "Attendee requested termination!\r\n";
            break;
        default:
            textReason = "Unknown reason!\r\n";
            break;
        }

        PrintTextToLog(h_logTextBox, textReason);

        pAttendee->Release();
        conn_picp = 0;
        conn_picpc = 0;
    }
}

void OnControlLevelChangeRequest(IDispatch * pAttendee, CTRL_LEVEL RequestedLevel)
{
    IRDPSRAPIAttendee* pRDPAtendee;
    pAttendee->QueryInterface(__uuidof(IRDPSRAPIAttendee), (void**)&pRDPAtendee);

    if (pRDPAtendee->put_ControlLevel(RequestedLevel) == S_OK) 
    {
        switch (RequestedLevel) 
        {
        case CTRL_LEVEL_NONE:
            PrintTextToLog(h_logTextBox, "Access level changed to \"DO NOT SHOW\"!\r\n");
            break;
        case CTRL_LEVEL_VIEW:
            PrintTextToLog(h_logTextBox, "Access level changed to \"VIEW SCREEN\"!\r\n");
            break;
        case CTRL_LEVEL_INTERACTIVE:
            PrintTextToLog(h_logTextBox, "Access level changed to \"PC CONTROL\"!\r\n");
            break;
        }
    }

    pAttendee->Release();
}

int ConnectEvent(IUnknown* container, REFIID riid, IUnknown* advisor,
    IConnectionPointContainer** picpc, IConnectionPoint** picp)
{
    HRESULT hRes = 0;
    unsigned long tid = 0;
    IConnectionPointContainer* icpc = 0;
    IConnectionPoint* icp = 0;
    *picpc = 0;
    *picp = 0;

    container->QueryInterface(IID_IConnectionPointContainer, (void**)&icpc);

    if (icpc)
    {
        *picpc = icpc;

        icpc->FindConnectionPoint(riid, &icp);

        if (icp)
        {
            *picp = icp;

            hRes = icp->Advise(advisor, &tid);
        }
    }

    return tid;
}

char* CreateInvitationFile() {
    OPENFILENAME openFileName;

    char* fileName = new char[100];

    ZeroMemory(&openFileName, sizeof(openFileName));

    openFileName.lStructSize = sizeof(openFileName);
    openFileName.hwndOwner = h_mainWindow;
    openFileName.lpstrFile = fileName;
    openFileName.nMaxFile = sizeof(openFileName);
    openFileName.nFilterIndex = 1;
    openFileName.lpstrDefExt = "xml";
    openFileName.lpstrFilter = "XML\0*.xml\0";
    openFileName.lpstrFileTitle = 0;
    openFileName.nMaxFileTitle = 0;
    openFileName.lpstrInitialDir = 0;
    openFileName.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;
    openFileName.lpstrFile[0] = '\0';

    if (GetSaveFileName(&openFileName))
        return fileName;

    return NULL;
}

void StartServerInviter() {
    if (rdp_session == NULL) 
    {
        HRESULT resOfCreatingInstance = CoCreateInstance(
            __uuidof(RDPSession),               
            NULL,                               
            CLSCTX_INPROC_SERVER,               
            __uuidof(IRDPSRAPISharingSession),  
            (void**)&rdp_session);              

        if (resOfCreatingInstance == S_OK)
        {
            PrintTextToLog(h_logTextBox, "Instance created!\r\n");
            ConnectEvent((IUnknown*)rdp_session, __uuidof(_IRDPSessionEvents), (IUnknown*)&rdpEvents, &conn_picpc, &conn_picp);

            if (rdp_session->Open() == S_OK) 
            {
                PrintTextToLog(h_logTextBox, "Session opened!\r\n");

                if (rdp_session->get_Invitations(&rdp_invitationManager) == S_OK) 
                {
                    HRESULT resOfCreatingInvitation = rdp_invitationManager->CreateInvitation(
                        SysAllocString(L"R-PControl"),        // String to use for the authorization.
                        SysAllocString(L"R-PControl-Group"),  // The name of the group.
                        SysAllocString(L""),                  // Password to use for authentication.
                        MAX_NUM_ATTENDEES,                    // The maximum number of attendees.
                        &rdp_invitation);                     // An IRDPSRAPIInvitation interface pointer.

                    if (resOfCreatingInvitation == S_OK)
                    {
                        PrintTextToLog(h_logTextBox, "Invitation obtained!\r\n");

                        char* invitationFile = CreateInvitationFile();

                        if (invitationFile) 
                        {
                            ofstream fileStream;
                            fileStream.open(invitationFile);

                            if (fileStream.is_open()) 
                            {
                                BSTR invitationString;

                                if (rdp_invitation->get_ConnectionString(&invitationString) == S_OK) 
                                {
                                    fileStream << _bstr_t(invitationString, false);
                                    SysFreeString(invitationString);

                                    PrintTextToLog(h_logTextBox, "Invitation written to file!\r\n");
                                }

                                fileStream.close();
                            }

                            if (rdp_session->get_Attendees(&rdp_attendeeManager) == S_OK) 
                            {
                                PrintTextToLog(h_logTextBox, "\r\nWAITING FOR ATTENDEES...\r\n");
                            }
                        }
                        else
                            PrintTextToLog(h_logTextBox, "Invalid path to save the file!\r\n");
                    }
                    else
                        PrintTextToLog(h_logTextBox, "Error obtaining invitation!\r\n");
                }
                else
                    PrintTextToLog(h_logTextBox, "Get invitations error!\r\n");
            }
            else
                PrintTextToLog(h_logTextBox, "Error opening session!\r\n");
        }
        else
            PrintTextToLog(h_logTextBox, "Error creating instance!\r\n");
    }
    else
        PrintTextToLog(h_logTextBox, "Error starting: Session already exists!\r\n");
}

void CloseSession() 
{
    PrintTextToLog(h_logTextBox, "\r\nSTOPPING...\r\n");

    if (rdp_session)
    {
        rdp_session->Close();
        rdp_session->Release();
        rdp_session = NULL;

        PrintTextToLog(h_logTextBox, "Session stopped!\r\n");
    }
    else
        PrintTextToLog(h_logTextBox, "Error stopping: No active session!\r\n");
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

HWND RenderButton(HWND hWnd, LPCSTR text, int x, int y, int width, int height) 
{
    return CreateWindow("button", text,
        WS_CHILD | WS_VISIBLE | BS_TEXT,
        x, y, width, height,
        hWnd, NULL, NULL, NULL);
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
        RenderLabel(hWnd, "Log:", 10, 5, 30, 18);
        h_logTextBox = RenderTextBox(hWnd, NULL, TRUE, 10, 28, 365, 280);
        h_startSharingBtn = RenderButton(hWnd, "Start sharing", 10, 320, 177, 30);
        h_stopSharingBtn = RenderButton(hWnd, "Stop sharing", 197, 320, 177, 30);
        EnableWindow(h_stopSharingBtn, FALSE);
        break;
    case WM_COMMAND:
        if ((HWND)lParam == h_startSharingBtn) 
        {
            EnableWindow(h_startSharingBtn, FALSE);
            EnableWindow(h_stopSharingBtn, TRUE);

            StartServerInviter();
        }
        if ((HWND)lParam == h_stopSharingBtn)
        {
            EnableWindow(h_startSharingBtn, TRUE);
            EnableWindow(h_stopSharingBtn, FALSE);

            CloseSession();
        }
        break;
    case WM_CTLCOLORSTATIC:
        return (INT_PTR)CreateSolidBrush(RGB(255, 255, 255));
    case WM_CLOSE:
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