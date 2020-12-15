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

HWND _mainWindow;
HWND _logTextBox;
HWND _startSharingBtn;
HWND _stopSharingBtn;

IRDPSRAPISharingSession* session = NULL;
IRDPSRAPIInvitationManager* invitationManager = NULL;
IRDPSRAPIInvitation* invitation = NULL;
IRDPSRAPIAttendeeManager* attendeeManager = NULL;
IRDPSRAPIAttendee* attendee = NULL;

IConnectionPointContainer* picpc = NULL;
IConnectionPoint* picp = NULL;

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
    /*
      Queries a COM object for a pointer to one of its interface;
      identifying the interface by a reference to its interface identifier (IID).
      If the COM object implements the interface,
      then it returns a pointer to that interface after calling IUnknown::AddRef on it.
    */
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

    // Provides access to properties and methods exposed by an object.
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

    PrintTextToLog(_logTextBox, "An attendee connected!\r\n");
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

        PrintTextToLog(_logTextBox, textReason);

        pAttendee->Release();
        picp = 0;
        picpc = 0;
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
            PrintTextToLog(_logTextBox, "Access level changed to \"DO NOT SHOW\"!\r\n");
            break;
        case CTRL_LEVEL_VIEW:
            PrintTextToLog(_logTextBox, "Access level changed to \"VIEW SCREEN\"!\r\n");
            break;
        case CTRL_LEVEL_INTERACTIVE:
            PrintTextToLog(_logTextBox, "Access level changed to \"PC CONTROL\"!\r\n");
            break;
        }
    }

    pAttendee->Release();
}

char* CreateInvitationFile() {
    OPENFILENAME openFileName;

    char* fileName = new char[100];

    ZeroMemory(&openFileName, sizeof(openFileName));

    openFileName.lStructSize = sizeof(openFileName);
    openFileName.hwndOwner = _mainWindow;
    openFileName.lpstrFile = fileName;
    openFileName.lpstrFile[0] = 0;
    openFileName.nMaxFile = sizeof(openFileName);
    openFileName.nFilterIndex = 1;
    openFileName.lpstrDefExt = "xml";
    openFileName.lpstrFilter = "XML\0*.xml\0";
    openFileName.lpstrFileTitle = 0;
    openFileName.nMaxFileTitle = 0;
    openFileName.lpstrInitialDir = 0;
    openFileName.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

    if (GetSaveFileName(&openFileName))
        return fileName;

    return NULL;
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

    /*
        Queries a COM object for a pointer to one of its interface;
        identifying the interface by a reference to its interface identifier (IID).
        If the COM object implements the interface, 
        then it returns a pointer to that interface after calling IUnknown::AddRef on it.
    */ 
    container->QueryInterface(IID_IConnectionPointContainer, (void**)&icpc);

    if (icpc)
    {
        *picpc = icpc;

        /*
            Returns a pointer to the IConnectionPoint interface of a connection point for a specified IID,
            if that IID describes a supported outgoing interface.
        */
        icpc->FindConnectionPoint(riid, &icp);

        if (icp)
        {
            *picp = icp;

            // Establishes a connection between a connection point object and the client's sink.
            hRes = icp->Advise(advisor, &tid);
        }
    }

    return tid;
}

void CloseSession() {
    PrintTextToLog(_logTextBox, "\r\nSTOPPING...\r\n");
    
    if (session) 
    {
        session->Close();
        session->Release();
        session = NULL;

        PrintTextToLog(_logTextBox, "Session stopped!\r\n");
    }
    else
        PrintTextToLog(_logTextBox, "Error stopping: No active session!\r\n");
}

void StartServerInviter() {
    if (session == NULL) 
    {
        // Creates and default-initializes a single object of the class associated with a specified CLSID (RDPSession).
        HRESULT resOfCreatingInstance = CoCreateInstance(
            __uuidof(RDPSession),               // CLSID.
            NULL,                               // Indicates that the object is not being created as part of an aggregate.
            CLSCTX_INPROC_SERVER,               // Context in which the code that manages the newly created object will run.
            __uuidof(IRDPSRAPISharingSession),  // A reference to the id of the interface to be used to communicate with the object.
            (void**)&session);                  // Address of pointer variable that receives the interface pointer requested in riid.

        if (resOfCreatingInstance == S_OK)
        {
            PrintTextToLog(_logTextBox, "Instance created!\r\n");
            ConnectEvent((IUnknown*)session, __uuidof(_IRDPSessionEvents), (IUnknown*)&rdpEvents, &picpc, &picp);

            if (session->Open() == S_OK) 
            {
                PrintTextToLog(_logTextBox, "Session opened!\r\n");

                if (session->get_Invitations(&invitationManager) == S_OK) 
                {
                    HRESULT resOfCreatingInvitation = invitationManager->CreateInvitation(
                        SysAllocString(L"R-PControl"),        // String to use for the authorization.
                        SysAllocString(L"R-PControl-Group"),  // The name of the group.
                        SysAllocString(L""),                  // Password to use for authentication.
                        MAX_NUM_ATTENDEES,                    // The maximum number of attendees.
                        &invitation);                         // An IRDPSRAPIInvitation interface pointer.

                    if (resOfCreatingInvitation == S_OK)
                    {
                        PrintTextToLog(_logTextBox, "Invitation obtained!\r\n");

                        char* invitationFile = CreateInvitationFile();

                        if (invitationFile) 
                        {
                            ofstream fileStream;
                            fileStream.open(invitationFile);

                            if (fileStream.is_open()) 
                            {
                                BSTR invitationString;

                                if (invitation->get_ConnectionString(&invitationString) == S_OK) 
                                {
                                    fileStream << _bstr_t(invitationString, false);
                                    SysFreeString(invitationString);

                                    PrintTextToLog(_logTextBox, "Invitation written to file!\r\n");
                                }

                                fileStream.close();
                            }

                            if (session->get_Attendees(&attendeeManager) == S_OK) 
                            {
                                PrintTextToLog(_logTextBox, "\r\nWAITING FOR ATTENDEES...\r\n");
                            }
                        }
                        else
                            PrintTextToLog(_logTextBox, "Invalid path to save the file.!\r\n");
                    }
                    else
                        PrintTextToLog(_logTextBox, "Error obtaining invitation!\r\n");
                }
                else
                    PrintTextToLog(_logTextBox, "Get invitations error!\r\n");
            }
            else
                PrintTextToLog(_logTextBox, "Error opening session!\r\n");
        }
        else
            PrintTextToLog(_logTextBox, "Error creating instance!\r\n");
    }
    else
        PrintTextToLog(_logTextBox, "Error starting: Session already exists!\r\n");
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
        RenderLabel(hWnd, "Log:", 10, 5, 30, 18);
        _logTextBox = RenderTextBox(hWnd, NULL, TRUE, 10, 28, 365, 280);
        _startSharingBtn = RenderButton(hWnd, "Start sharing", 10, 320, 177, 30);
        _stopSharingBtn = RenderButton(hWnd, "Stop sharing", 197, 320, 177, 30);
        EnableWindow(_stopSharingBtn, FALSE);
        break;
    case WM_COMMAND:
        if ((HWND)lParam == _startSharingBtn) 
        {
            EnableWindow(_startSharingBtn, FALSE);
            EnableWindow(_stopSharingBtn, TRUE);

            StartServerInviter();
        }
        if ((HWND)lParam == _stopSharingBtn)
        {
            EnableWindow(_startSharingBtn, TRUE);
            EnableWindow(_stopSharingBtn, FALSE);

            CloseSession();
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