#define _WIN32_DCOM
#include <windows.h>
#include <iostream>
#include <comdef.h>
#include <taskschd.h>
#include <string>

#pragma comment(lib, "taskschd.lib")
#pragma comment(lib, "comsupp.lib")

using namespace std;

#ifdef _DEBUG
    WCHAR wPath[64] = L"C:\\Users\\Аркадий\\Source\\Repos\\BSIT_3\\Debug\\handler.exe";
#else
    WCHAR wPath[64] = L"C:\\Users\\Аркадий\\Source\\Repos\\BSIT_3\\Release\\handler.exe";
#endif

int tasksList()
{
    HRESULT hr = CoInitializeEx(
        NULL, 
        COINIT_MULTITHREADED
    );
    if (FAILED(hr))
    {
        cout << "CoInitializeEx error " << hr << endl;
        return 1;
    }

    hr = CoInitializeSecurity(
        NULL, 
        -1, 
        NULL, 
        NULL, 
        RPC_C_AUTHN_LEVEL_PKT_PRIVACY, 
        RPC_C_IMP_LEVEL_IMPERSONATE, 
        NULL, 
        0, 
        NULL
    );
    if (FAILED(hr))
    {
        cout << "CoInitializeSecurity error " << hr << endl;
        CoUninitialize();
        return 1;
    }

    ITaskService *pService = NULL;
    hr = CoCreateInstance(
        CLSID_TaskScheduler,
        NULL,
        CLSCTX_INPROC_SERVER,
        IID_ITaskService,
        (void **)&pService
    );
    if (FAILED(hr))
    {
        cout << "CoCreateInstance error " << hr << endl;
        CoUninitialize();
        return 1;
    }

    hr = pService->Connect(
        _variant_t(), 
        _variant_t(),
        _variant_t(), 
        _variant_t()
    );
    if (FAILED(hr))
    {
        cout << "ITaskService::Connect error " << hr << endl;
        pService->Release();
        CoUninitialize();
        return 1;
    }

    ITaskFolder *pRootFolder = NULL;
    hr = pService->GetFolder(
        _bstr_t(L"\\"), 
        &pRootFolder
    );
    pService->Release();
    if (FAILED(hr))
    {
        cout << "ITaskFolder::GetFolder error " << hr << endl;
        CoUninitialize();
        return 1;
    }

    IRegisteredTaskCollection *pTaskCollection = NULL;
    hr = pRootFolder->GetTasks(
        NULL, 
        &pTaskCollection
    );
    pRootFolder->Release();
    if (FAILED(hr))
    {
        cout << "ITaskService::GetTasks error " << hr << endl;
        CoUninitialize();
        return 1;
    }


    LONG lTasksNum = 0;
    hr = pTaskCollection->get_Count(&lTasksNum);

    if (lTasksNum == 0)
    {
        cout << "No active tasks" << endl;
        pTaskCollection->Release();
        CoUninitialize();
        return 0;
    }

    cout << "Tasks list:" << endl << endl;

    TASK_STATE taskState;

    for (LONG i = 0; i < lTasksNum; i++)
    {
        IRegisteredTask *pRegisteredTask = NULL;
        hr = pTaskCollection->get_Item(_variant_t(i + 1), &pRegisteredTask);

        if (SUCCEEDED(hr))
        {
            BSTR taskName = NULL;
            hr = pRegisteredTask->get_Name(&taskName);
            if (SUCCEEDED(hr))
            {
                wcout << "Name:  " << taskName << endl;
                SysFreeString(taskName);

                hr = pRegisteredTask->get_State(&taskState);
                if (SUCCEEDED(hr))
                {
                    switch (taskState)
                    {
                    case TASK_STATE_UNKNOWN:
                    {
                        cout << "State: Unknown (TASK_STATE_UNKNOWN)" << endl;
                        break;
                    }
                    case TASK_STATE_DISABLED:
                    {
                        cout << "State: Disabled (TASK_STATE_DISABLED)" << endl;
                        break;
                    }
                    case TASK_STATE_QUEUED:
                    {
                        cout << "State: Queued (TASK_STATE_QUEUED)" << endl;
                        break;
                    }
                    case TASK_STATE_READY:
                    {
                        cout << "State: Ready (TASK_STATE_READY)" << endl;
                        break;
                    }
                    case TASK_STATE_RUNNING:
                    {
                        cout << "State: Running (TASK_STATE_RUNNING)" << endl;
                        break;
                    }
                    default:
                    {
                        break;
                    }
                    }
                }
                else
                {
                    cout << "IRegisteredTask::get_State error " << hr << endl;
                }
            }
            else
            {
                cout << "IRegisteredTask::get_Name error " << hr << endl;
            }
            pRegisteredTask->Release();
        }
        else
        {
            cout << "IRegisteredTask::get_Item error " << hr << endl;
        }
        cout << endl;
    }

    pTaskCollection->Release();
    CoUninitialize();
    return 0;
}

int createTask(
    LPCWSTR wTaskName, 
    LPCWSTR wTaskPath, 
    LPCWSTR wQuery, 
    LPCWSTR wArgs, 
    BOOL    bPutValues, 
    LPCWSTR wValueName, 
    LPCWSTR wValueValue
)
{
    HRESULT hr = CoInitializeEx(
        NULL, 
        COINIT_MULTITHREADED
    );
    if (FAILED(hr))
    {
        cout << "CoInitializeEx error " << hr << endl;
        return 1;
    }

    hr = CoInitializeSecurity(
        NULL, 
        -1, 
        NULL, 
        NULL, 
        RPC_C_AUTHN_LEVEL_PKT_PRIVACY, 
        RPC_C_IMP_LEVEL_IMPERSONATE, 
        NULL, 
        0, 
        NULL
    );
    if (FAILED(hr))
    {
        cout << "CoInitializeSecurity error " << hr << endl;
        CoUninitialize();
        return 1;
    }

    ITaskService *pService = NULL;
    hr = CoCreateInstance(
        CLSID_TaskScheduler,
        NULL,
        CLSCTX_INPROC_SERVER,
        IID_ITaskService,
        (void **)&pService);
    if (FAILED(hr))
    {
        cout << "CoCreateInstance error " << hr << endl;
        CoUninitialize();
        return 1;
    }

    hr = pService->Connect(
        _variant_t(), 
        _variant_t(),
        _variant_t(), 
        _variant_t()
    );
    if (FAILED(hr))
    {
        cout << "ITaskService::Connect error " << hr << endl;
        pService->Release();
        CoUninitialize();
        return 1;
    }

    ITaskFolder *pRootFolder = NULL;
    hr = pService->GetFolder(
        _bstr_t(L"\\"), 
        &pRootFolder
    );
    if (FAILED(hr))
    {
        cout << "ITaskFolder::GetFolder error " << hr << endl;
        pService->Release();
        CoUninitialize();
        return 1;
    }

    pRootFolder->DeleteTask(
        _bstr_t(wTaskName), 
        0
    );
    ITaskDefinition *pTask = NULL;
    hr = pService->NewTask(
        0, 
        &pTask
    );
    pService->Release();
    if (FAILED(hr))
    {
        cout << "ITaskService::NewTask error " << hr << endl;
        pRootFolder->Release();
        CoUninitialize();
        return 1;
    }

    IRegistrationInfo *pRegInfo = NULL;
    hr = pTask->get_RegistrationInfo(&pRegInfo);
    if (FAILED(hr))
    {
        cout << "ITaskDefinition::get_RegistrationInfo error " << hr << endl;
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    ITaskSettings *pSettings = NULL;
    hr = pTask->get_Settings(&pSettings);
    if (FAILED(hr))
    {
        cout << "ITaskDefinition::get_Settings error " << hr << endl;
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    hr = pSettings->put_StartWhenAvailable(VARIANT_TRUE);
    pSettings->Release();
    if (FAILED(hr))
    {
        cout << "ITaskSettings::put_StartWhenAvailable error " << hr << endl;
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    ITriggerCollection *pTriggerCollection = NULL;
    hr = pTask->get_Triggers(&pTriggerCollection);
    if (FAILED(hr))
    {
        cout << "ITaskDefinition::get_Triggers error " << hr << endl;
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    ITrigger *pTrigger = NULL;
    hr = pTriggerCollection->Create(
        TASK_TRIGGER_EVENT,
        &pTrigger
    );
    pTriggerCollection->Release();
    if (FAILED(hr))
    {
        cout << "ITriggerCollection::Create error " << hr << endl;
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    IEventTrigger *pEventTrigger = NULL;
    hr = pTrigger->QueryInterface(
        IID_IEventTrigger, 
        (void **)&pEventTrigger
    );
    pTrigger->Release();
    if (FAILED(hr))
    {
        cout << "QueryInterface error " << hr << endl;
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    hr = pEventTrigger->put_Id(_bstr_t(L"Trigger1"));
    if (FAILED(hr))
    {
        cout << "ITrigger::put_Id error " << hr << endl;
    }

    hr = pEventTrigger->put_Subscription((BSTR)wQuery);
    if (FAILED(hr))
    {
        cout << "IEventTrigger::put_Subscription error " << hr << endl;
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    if (bPutValues)
    {
        ITaskNamedValueCollection *pValueQueries = NULL;
        pEventTrigger->get_ValueQueries(&pValueQueries);
        pValueQueries->Create(
            (BSTR)wValueName, 
            (BSTR)wValueValue, 
            NULL
        );
        hr = pEventTrigger->put_ValueQueries(pValueQueries);
        if (FAILED(hr))
        {
            cout << "IEventTrigger::put_ValueQueries error " << hr << endl;
            pRootFolder->Release();
            pTask->Release();
            CoUninitialize();
            return 1;
        }
    }

    pEventTrigger->Release();
    IActionCollection *pActionCollection = NULL;
    hr = pTask->get_Actions(&pActionCollection);
    if (FAILED(hr))
    {
        cout << "ITaskDefinition::get_Actions error " << hr << endl;
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    IAction *pAction = NULL;
    hr = pActionCollection->Create(
        TASK_ACTION_EXEC, 
        &pAction
    );
    pActionCollection->Release();
    if (FAILED(hr))
    {
        cout << "IActionCollection::Create error " << hr << endl;
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    IExecAction2 *pExecAction = NULL;
    hr = pAction->QueryInterface(
        IID_IExecAction2, 
        (void **)&pExecAction
    );
    pAction->Release();
    if (FAILED(hr))
    {
        cout << "QueryInterface error " << hr << endl;
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    hr = pExecAction->put_Path((BSTR)wTaskPath);
    if (FAILED(hr))
    {
        cout << "IExecAction::put_Path error " << hr << endl;
        pRootFolder->Release();
        pExecAction->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    hr = pExecAction->put_Arguments((BSTR)wArgs);
    if (FAILED(hr))
    {
        cout << "IExecAction::put_Arguments error " << hr << endl;
        pRootFolder->Release();
        pExecAction->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }
    pExecAction->Release();

    IRegisteredTask *pRegisteredTask = NULL;
    hr = pRootFolder->RegisterTaskDefinition(
        _bstr_t(wTaskName),
        pTask,
        TASK_CREATE_OR_UPDATE,
        _variant_t(_bstr_t(L"")),
        _variant_t(_bstr_t(L"")),
        TASK_LOGON_INTERACTIVE_TOKEN,
        _variant_t(L""),
        &pRegisteredTask
    );
    if (FAILED(hr))
    {
        cout << "ITaskFolder::RegisterTaskDefinition error " << hr << endl;
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        cout << "Task was not created" << endl;
        return 1;
    }

    pRootFolder->Release();
    pTask->Release();
    pRegisteredTask->Release();
    CoUninitialize();
    cout << "Task was successfully created" << endl;

    return 0;
}

int deleteTask(LPCWSTR wTaskName)
{
    HRESULT hr = CoInitializeEx(
        NULL, 
        COINIT_MULTITHREADED
    );
    if (FAILED(hr))
    {
        cout << "CoInitializeEx error " << hr << endl;
        return 1;
    }

    hr = CoInitializeSecurity(
        NULL,
        -1,
        NULL,
        NULL,
        RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        0,
        NULL);
    if (FAILED(hr))
    {
        cout << "CoInitializeSecurity error " << hr << endl;
        CoUninitialize();
        return 1;
    }

    ITaskService *pService = NULL;
    hr = CoCreateInstance(
        CLSID_TaskScheduler,
        NULL,
        CLSCTX_INPROC_SERVER,
        IID_ITaskService,
        (void **)&pService
    );
    if (FAILED(hr))
    {
        cout << "CoCreateInstance error " << hr << endl;
        CoUninitialize();
        return 1;
    }

    hr = pService->Connect(
        _variant_t(), 
        _variant_t(),
        _variant_t(), 
        _variant_t()
    );
    if (FAILED(hr))
    {
        cout << "ITaskService::Connect error " << hr << endl;
        pService->Release();
        CoUninitialize();
        return 1;
    }

    ITaskFolder *pRootFolder = NULL;
    hr = pService->GetFolder(
        _bstr_t(L"\\"), 
        &pRootFolder
    );
    if (FAILED(hr))
    {
        cout << "ITaskFolder::GetFolder error " << hr << endl;
        pService->Release();
        CoUninitialize();
        return 1;
    }

    hr = pRootFolder->DeleteTask(
        _bstr_t(wTaskName),
        0
    );
    if (FAILED(hr))
    {
        cout << "ITaskFolder::DeleteTask error " << hr << endl;
        pService->Release();
        CoUninitialize();
        cout << "Task was not deleted" << endl;
        return 1;
    }

    pRootFolder->Release();
    CoUninitialize();
    cout << "Task was successfully deleted" << endl;

    return 0;
}

int main()
{
    string sArg;
    while (1)
    {
        cout << "> ";
        cin >> sArg;
        if (sArg == "?")
        {
            cout << "Usage:"                          << endl;
            cout << "sh"                              << "\t\t\t\t-\t" << "show tasks"                                          << endl;
            cout << "notify {security|ping} {on|off}" << "\t-\t"       << "notifications:"                                      << endl;
            cout << "\t\t\t\t"                        << "\t\t"        << "security - changes in Windows Defender and firewall" << endl;
            cout << "\t\t\t\t"                        << "\t\t"        << "ping     - firewall blocks ping request"             << endl;
            cout << "\t\t\t\t"                        << "\t\t\t"      << "on  - enable"                                        << endl;
            cout << "\t\t\t\t"                        << "\t\t\t"      << "off - disable"                                       << endl;
            cout << "exit"                            << "\t\t\t\t-\t" << "exit program"                                        << endl;
        }
        else if (sArg == "sh")
        {
            tasksList();
        }
        else if (sArg == "notify")
        {
            cin >> sArg;
            if (sArg == "security")
            {
                cin >> sArg;
                if (sArg == "on")
                {
                    createTask(
                        L"FirewallNotification",
                        wPath,
                        L"<QueryList>\<Query Id=\"0\" Path=\"Microsoft-Windows-Windows Firewall With Advanced Security/Firewall\">\<Select Path=\"Microsoft-Windows-Windows Firewall With Advanced Security/Firewall\">*</Select>\</Query>\</QueryList>",
                        L"Firewall settings were changed",
                        FALSE,
                        NULL,
                        NULL
                    );
                    createTask(
                        L"WindowsDefenderNotification",
                        wPath,
                        L"<QueryList>\<Query Id=\"0\" Path=\"Microsoft-Windows-Windows Defender/Operational\">\<Select Path=\"Microsoft-Windows-Windows Defender/Operational\">*[System[Provider[@Name='Microsoft-Windows-Windows Defender'] and (EventID = 5007)]]</Select>\<Select Path=\"Microsoft-Windows-Windows Defender/WHC\">*[System[Provider[@Name='Microsoft-Windows-Windows Defender'] and (EventID = 5007)]]</Select>\</Query>\</QueryList>",
                        L"Windows Defender settings were changed",
                        FALSE,
                        NULL,
                        NULL
                    );
                }
                else if (sArg == "off")
                {
                    deleteTask(L"FirewallNotification");
                    deleteTask(L"WindowsDefenderNotification");
                }
                else
                {
                    cout << "notify {security|ping} {on|off}" <<  "\t-\t" << "notifications:" << endl;
                    cout << "\t\t\t\t"                        <<  "\t\t"  << "on  - enable"   << endl;
                    cout << "\t\t\t\t"                        <<  "\t\t"  << "off - disable"  << endl;
                }
            }
            else if (sArg == "ping")
            {
                cin >> sArg;
                if (sArg == "on")
                {
                    createTask(
                        L"PingBlockNotification",
                        wPath,
                        L"<QueryList>\<Query Id=\"0\" Path=\"Security\">\<Select Path=\"Security\">*[System[(EventID = 5152)]]</Select>\</Query>\</QueryList>",
                        L"ping $(srcIp)",
                        TRUE,
                        L"srcIp",
                        L"Event/EventData/Data[@Name='SourceAddress']"
                    );
                }
                else if (sArg == "off")
                {
                    deleteTask(L"PingBlockNotification");
                }
                else
                {
                    cout << "notify {security|ping} {on|off}" << "\t-\t" << "notifications:" << endl;
                    cout << "\t\t\t\t"                        << "\t\t"  << "on  - enable"   << endl;
                    cout << "\t\t\t\t"                        << "\t\t"  << "off - disable"  << endl;
                }
            }
            else
            {
                cout << "notify {security|ping} {on|off}" << "\t-\t" << "notifications:"                                      << endl;
                cout << "\t\t\t\t"                        << "\t\t"  << "security - changes in Windows Defender and firewall" << endl;
                cout << "\t\t\t\t"                        << "\t\t"  << "ping     - firewall blocks ping request"             << endl;
            }
        }
        else if (sArg == "exit")
        {
            return 0;
        }
        else
        {
            cout << "Unrecognized input. Try ? for help" << endl;
        }
    }
}
