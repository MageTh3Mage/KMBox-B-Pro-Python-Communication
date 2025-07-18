#define PY_SSIZE_T_CLEAN
#include <Python.h>
#include <windows.h>
#include <setupapi.h>
#include <devguid.h>
#include <regstr.h>
#include <string>
#include <cstdio>
#include <memory>

#pragma comment(lib, "setupapi.lib")
#pragma comment(lib, "advapi32.lib")

#define DEBUG_PRINT(self, fmt, ...) \
    do { if ((self)->debug) fprintf(stderr, fmt "\n", __VA_ARGS__); } while (0)

struct HandleWrapper {
    HANDLE handle = INVALID_HANDLE_VALUE;
    ~HandleWrapper() { if (handle != INVALID_HANDLE_VALUE) CloseHandle(handle); }
    explicit operator bool() const { return handle != INVALID_HANDLE_VALUE; }
};

typedef struct {
    PyObject_HEAD
    HandleWrapper hSerial;
    bool is_connected = false;
    int debug = 0;
} KmboxObject;

static char* find_ch340_port(int debug) {
    HDEVINFO hDevInfo = SetupDiGetClassDevs(&GUID_DEVCLASS_PORTS, nullptr, nullptr, DIGCF_PRESENT);
    if (hDevInfo == INVALID_HANDLE_VALUE) {
        if (debug) fprintf(stderr, "[ERROR] SetupDiGetClassDevs failed\n");
        return nullptr;
    }

    SP_DEVINFO_DATA DeviceInfoData = { sizeof(SP_DEVINFO_DATA) };
    for (DWORD i = 0; SetupDiEnumDeviceInfo(hDevInfo, i, &DeviceInfoData); i++) {
        DWORD DataT;
        char friendlyName[256] = {0};
        char portName[256] = {0};
        DWORD size = 0;

        if (SetupDiGetDeviceRegistryPropertyA(hDevInfo, &DeviceInfoData, SPDRP_FRIENDLYNAME,
            &DataT, (PBYTE)friendlyName, sizeof(friendlyName), &size)) {
            if (debug) fprintf(stderr, "[DEBUG] Found device: %s\n", friendlyName);
            if (strstr(friendlyName, "CH340") || strstr(friendlyName, "USB-SERIAL")) {
                HKEY hKey = SetupDiOpenDevRegKey(hDevInfo, &DeviceInfoData, DICS_FLAG_GLOBAL, 0, DIREG_DEV, KEY_READ);
                if (hKey != INVALID_HANDLE_VALUE) {
                    DWORD len = sizeof(portName);
                    if (RegQueryValueExA(hKey, "PortName", nullptr, nullptr, (LPBYTE)portName, &len) == ERROR_SUCCESS) {
                        if (portName[0] != '\0') {
                            char full_path[64];
                            snprintf(full_path, sizeof(full_path), "\\\\.\\%s", portName);

                            HANDLE hTest = CreateFileA(full_path, GENERIC_READ | GENERIC_WRITE, 0, nullptr,
                                                      OPEN_EXISTING, 0, nullptr);
                            if (hTest != INVALID_HANDLE_VALUE) {
                                CloseHandle(hTest);
                                RegCloseKey(hKey);
                                SetupDiDestroyDeviceInfoList(hDevInfo);
                                return _strdup(portName);
                            }
                            else if (debug) {
                                fprintf(stderr, "[WARN] Failed to open test port: %s\n", full_path);
                            }
                        }
                    }
                    RegCloseKey(hKey);
                }
            }
        }
    }

    SetupDiDestroyDeviceInfoList(hDevInfo);
    return nullptr;
}

static int Kmbox_init(KmboxObject* self, PyObject* args, PyObject* kwds) {
    const char* port = nullptr;
    int baudrate = 115200;
    int debug = 0;

    static char* kwlist[] = { "port", "baudrate", "debug", nullptr };
    if (!PyArg_ParseTupleAndKeywords(args, kwds, "|sip", kwlist, &port, &baudrate, &debug))
        return -1;

    self->debug = debug;
    self->is_connected = false;

    char* found_port = nullptr;
    if (!port || port[0] == '\0') {
        if (self->debug) fprintf(stderr, "[INFO] Searching for CH340 device...\n");
        found_port = find_ch340_port(self->debug);
        if (!found_port) {
            PyErr_SetString(PyExc_IOError, "No compatible CH340 device found");
            return -1;
        }
        port = found_port;
    }

    char full_port_path[64];
    snprintf(full_port_path, sizeof(full_port_path), "\\\\.\\%s", port);

    if (self->debug) fprintf(stderr, "[DEBUG] Opening port: %s\n", full_port_path);

    HANDLE hSerial = CreateFileA(full_port_path, GENERIC_READ | GENERIC_WRITE, 0, nullptr,
                                OPEN_EXISTING, 0, nullptr);

    if (found_port) free(found_port);

    if (hSerial == INVALID_HANDLE_VALUE) {
        DWORD err = GetLastError();
        if (self->debug)
            fprintf(stderr, "[ERROR] CreateFile failed: %s (error %lu)\n", full_port_path, err);
        PyErr_Format(PyExc_IOError, "Failed to open port: %s (error %lu)", port, err);
        return -1;
    }

    self->hSerial.handle = hSerial;
    DCB dcbSerialParams = {0};
    dcbSerialParams.DCBlength = sizeof(DCB);

    if (!GetCommState(self->hSerial.handle, &dcbSerialParams)) {
        PyErr_SetString(PyExc_IOError, "Failed to get serial parameters");
        return -1;
    }

    dcbSerialParams.BaudRate = baudrate;
    dcbSerialParams.ByteSize = 8;
    dcbSerialParams.StopBits = ONESTOPBIT;
    dcbSerialParams.Parity = NOPARITY;

    if (!SetCommState(self->hSerial.handle, &dcbSerialParams)) {
        PyErr_SetString(PyExc_IOError, "Failed to set serial parameters");
        return -1;
    }

    COMMTIMEOUTS timeouts = {0};
    timeouts.ReadIntervalTimeout = 50;
    timeouts.ReadTotalTimeoutConstant = 50;
    timeouts.ReadTotalTimeoutMultiplier = 10;
    timeouts.WriteTotalTimeoutConstant = 50;
    timeouts.WriteTotalTimeoutMultiplier = 10;
    SetCommTimeouts(self->hSerial.handle, &timeouts);

    self->is_connected = true;
    if (self->debug)
        fprintf(stderr, "[INFO] Successfully connected to %s\n", port);

    return 0;
}

static void send_command(KmboxObject* self, const char* cmd) {
    if (!self->is_connected) return;

    DWORD bytes_written = 0;
    const size_t len = strlen(cmd);
    WriteFile(self->hSerial.handle, cmd, (DWORD)len, &bytes_written, nullptr);
    if (self->debug) FlushFileBuffers(self->hSerial.handle);
}

static PyObject* Kmbox_move(KmboxObject* self, PyObject* args) {
    int x, y;
    if (!PyArg_ParseTuple(args, "ii", &x, &y))
        return nullptr;

    char cmd[64];
    int len = snprintf(cmd, sizeof(cmd), "km.move(%d,%d)\n", x, y);
    if (len > 0) send_command(self, cmd);

    Py_RETURN_NONE;
}

static PyObject* Kmbox_left_click(KmboxObject* self, PyObject*) {
    send_command(self, "km.click(0)\n");
    Py_RETURN_NONE;
}

static PyObject* Kmbox_right_click(KmboxObject* self, PyObject*) {
    send_command(self, "km.click(1)\n");
    Py_RETURN_NONE;
}

static PyObject* Kmbox_middle_click(KmboxObject* self, PyObject*) {
    send_command(self, "km.click(2)\n");
    Py_RETURN_NONE;
}

static PyObject* Kmbox_is_connected(KmboxObject* self, PyObject*) {
    return PyBool_FromLong(self->is_connected);
}

static PyObject* Kmbox_close(KmboxObject* self, PyObject*) {
    if (self->is_connected) {
        CloseHandle(self->hSerial.handle);
        self->hSerial.handle = INVALID_HANDLE_VALUE;
        self->is_connected = false;
        if (self->debug) fprintf(stderr, "[INFO] Connection closed\n");
    }
    Py_RETURN_NONE;
}

static void Kmbox_dealloc(KmboxObject* self) {
    if (self->is_connected) {
        CloseHandle(self->hSerial.handle);
        self->hSerial.handle = INVALID_HANDLE_VALUE;
    }
    Py_TYPE(self)->tp_free((PyObject*)self);
}

static PyMethodDef Kmbox_methods[] = {
    {"move", (PyCFunction)Kmbox_move, METH_VARARGS, "Move the mouse"},
    {"left_click", (PyCFunction)Kmbox_left_click, METH_NOARGS, "Left click"},
    {"right_click", (PyCFunction)Kmbox_right_click, METH_NOARGS, "Right click"},
    {"middle_click", (PyCFunction)Kmbox_middle_click, METH_NOARGS, "Middle click"},
    {"is_connected", (PyCFunction)Kmbox_is_connected, METH_NOARGS, "Check if connected"},
    {"close", (PyCFunction)Kmbox_close, METH_NOARGS, "Close the connection"},
    {nullptr, nullptr, 0, nullptr}
};

static PyTypeObject KmboxType = {
    PyVarObject_HEAD_INIT(nullptr, 0)
    "kmbox.Kmbox",                          /* tp_name */
    sizeof(KmboxObject),                    /* tp_basicsize */
    0,                                      /* tp_itemsize */
    (destructor)Kmbox_dealloc,              /* tp_dealloc */
    0,                                      /* tp_vectorcall_offset */
    nullptr,                                /* tp_getattr */
    nullptr,                                /* tp_setattr */
    nullptr,                                /* tp_as_async */
    nullptr,                                /* tp_repr */
    nullptr,                                /* tp_as_number */
    nullptr,                                /* tp_as_sequence */
    nullptr,                                /* tp_as_mapping */
    nullptr,                                /* tp_hash  */
    nullptr,                                /* tp_call */
    nullptr,                                /* tp_str */
    nullptr,                                /* tp_getattro */
    nullptr,                                /* tp_setattro */
    nullptr,                                /* tp_as_buffer */
    Py_TPFLAGS_DEFAULT,                     /* tp_flags */
    "KMBox serial communication object",    /* tp_doc */
    nullptr,                                /* tp_traverse */
    nullptr,                                /* tp_clear */
    nullptr,                                /* tp_richcompare */
    0,                                      /* tp_weaklistoffset */
    nullptr,                                /* tp_iter */
    nullptr,                                /* tp_iternext */
    Kmbox_methods,                          /* tp_methods */
    nullptr,                                /* tp_members */
    nullptr,                                /* tp_getset */
    nullptr,                                /* tp_base */
    nullptr,                                /* tp_dict */
    nullptr,                                /* tp_descr_get */
    nullptr,                                /* tp_descr_set */
    0,                                      /* tp_dictoffset */
    (initproc)Kmbox_init,                   /* tp_init */
    nullptr,                                /* tp_alloc */
    PyType_GenericNew,                      /* tp_new */
};

static PyModuleDef kmboxmodule = {
    PyModuleDef_HEAD_INIT,
    "kmbox",
    "Fast KMBox serial communication module.",
    -1,
    nullptr, nullptr, nullptr, nullptr, nullptr
};

PyMODINIT_FUNC PyInit_CH340(void) {
    PyObject* m;
    if (PyType_Ready(&KmboxType) < 0)
        return nullptr;

    m = PyModule_Create(&kmboxmodule);
    if (!m)
        return nullptr;

    Py_INCREF(&KmboxType);
    PyModule_AddObject(m, "Kmbox", (PyObject*)&KmboxType);
    return m;
}
