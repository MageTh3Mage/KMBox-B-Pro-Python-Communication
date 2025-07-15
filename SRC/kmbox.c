#define PY_SSIZE_T_CLEAN
#include <Python.h>
#include <windows.h>
#include <setupapi.h>
#include <devguid.h>
#include <regstr.h>
#include <stdio.h>
#include <string.h>
#include <errno.h>


#pragma comment(lib, "setupapi.lib")
#pragma comment(lib, "advapi32.lib")

typedef struct {
    PyObject_HEAD
    HANDLE hSerial;
    int is_connected;
    int debug;
} KmboxObject;

static char* find_ch340_port(int debug) {
    HDEVINFO hDevInfo;
    SP_DEVINFO_DATA DeviceInfoData;
    DWORD i;
    char* found_port = NULL;

    hDevInfo = SetupDiGetClassDevs(&GUID_DEVCLASS_PORTS, 0, 0, DIGCF_PRESENT);
    if (hDevInfo == INVALID_HANDLE_VALUE) {
        if (debug) printf("[ERROR] SetupDiGetClassDevs failed\n");
        return NULL;
    }

    DeviceInfoData.cbSize = sizeof(SP_DEVINFO_DATA);
    for (i = 0; SetupDiEnumDeviceInfo(hDevInfo, i, &DeviceInfoData); i++) {
        DWORD DataT;
        DWORD size;
        char friendlyName[256] = {0};
        char portName[256] = {0};

        if (SetupDiGetDeviceRegistryPropertyA(hDevInfo, &DeviceInfoData, SPDRP_FRIENDLYNAME,
                                              &DataT, (PBYTE)friendlyName, sizeof(friendlyName), &size)) {
            if (debug) printf("[DEBUG] Found device: %s\n", friendlyName);
            if (strstr(friendlyName, "CH340") || strstr(friendlyName, "USB-SERIAL")) {
                HKEY hKey = SetupDiOpenDevRegKey(hDevInfo, &DeviceInfoData, DICS_FLAG_GLOBAL, 0, DIREG_DEV, KEY_READ);
                if (hKey != INVALID_HANDLE_VALUE) {
                    DWORD len = sizeof(portName);
                    if (RegQueryValueExA(hKey, "PortName", NULL, NULL, (LPBYTE)portName, &len) == ERROR_SUCCESS) {
                        if (portName[0] != '\0') {
                            // Construct full path for COM10+
                            char full_path[64];
                            snprintf(full_path, sizeof(full_path), "\\\\.\\%s", portName);

                            HANDLE hTest = CreateFileA(full_path, GENERIC_READ | GENERIC_WRITE, 0, NULL,
                                                       OPEN_EXISTING, 0, NULL);
                            if (hTest != INVALID_HANDLE_VALUE) {
                                CloseHandle(hTest);
                                found_port = _strdup(portName);
                                if (debug) printf("[INFO] Found compatible device on %s\n", found_port);
                                RegCloseKey(hKey);
                                break;
                            } else {
                                if (debug) printf("[WARN] Failed to open test port: %s\n", full_path);
                            }
                        }
                    }
                    RegCloseKey(hKey);
                }
            }
        }
    }

    SetupDiDestroyDeviceInfoList(hDevInfo);
    return found_port;
}

static int Kmbox_init(KmboxObject *self, PyObject *args, PyObject *kwds) {
    const char *port = NULL;
    int baudrate = 115200;
    int debug = 0;
    char *found_port = NULL;
    char full_port_path[64] = {0};

    static char *kwlist[] = {"port", "baudrate", "debug", NULL};
    if (!PyArg_ParseTupleAndKeywords(args, kwds, "|sip", kwlist, &port, &baudrate, &debug))
        return -1;

    self->debug = debug;
    self->is_connected = 0;

    // Auto-discover CH340 device
    if (port == NULL || strcmp(port, "") == 0) {
        if (self->debug) printf("[INFO] Searching for CH340 device...\n");
        found_port = find_ch340_port(self->debug);
        if (!found_port) {
            if (self->debug) printf("[ERROR] No compatible CH340 device found\n");
            PyErr_SetString(PyExc_IOError, "No compatible CH340 device found");
            return -1;
        }
        port = found_port;
    }

    // Ensure full device path format for COM10+ and still valid for COM1–9
    snprintf(full_port_path, sizeof(full_port_path), "\\\\.\\%s", port);

    if (self->debug) printf("[DEBUG] Attempting to open: %s\n", full_port_path);

    self->hSerial = CreateFileA(
        full_port_path,
        GENERIC_READ | GENERIC_WRITE,
        0,
        NULL,
        OPEN_EXISTING,
        0,
        NULL
    );

    if (found_port) free(found_port);

    if (self->hSerial == INVALID_HANDLE_VALUE) {
        DWORD err = GetLastError();
        if (self->debug) {
            printf("[ERROR] CreateFileA failed: %s (error code: %lu)\n", full_port_path, err);
        }
        PyErr_Format(PyExc_IOError, "Failed to open port: %s (error %lu)", port, err);
        return -1;
    }

    DCB dcbSerialParams = {0};
    dcbSerialParams.DCBlength = sizeof(dcbSerialParams);

    if (!GetCommState(self->hSerial, &dcbSerialParams)) {
        CloseHandle(self->hSerial);
        PyErr_SetString(PyExc_IOError, "Failed to get serial parameters");
        return -1;
    }

    dcbSerialParams.BaudRate = baudrate;
    dcbSerialParams.ByteSize = 8;
    dcbSerialParams.StopBits = ONESTOPBIT;
    dcbSerialParams.Parity = NOPARITY;

    if (!SetCommState(self->hSerial, &dcbSerialParams)) {
        CloseHandle(self->hSerial);
        PyErr_SetString(PyExc_IOError, "Failed to set serial parameters");
        return -1;
    }

    COMMTIMEOUTS timeouts = {0};
    timeouts.ReadIntervalTimeout = 50;
    timeouts.ReadTotalTimeoutConstant = 50;
    timeouts.ReadTotalTimeoutMultiplier = 10;
    timeouts.WriteTotalTimeoutConstant = 50;
    timeouts.WriteTotalTimeoutMultiplier = 10;

    SetCommTimeouts(self->hSerial, &timeouts);

    self->is_connected = 1;
    if (self->debug) printf("[INFO] Successfully connected to %s\n", port);
    return 0;
}

static void send_command(KmboxObject *self, const char *cmd) {
    DWORD bytes_written;
    char buffer[256];
    snprintf(buffer, sizeof(buffer), "%s\n", cmd);
    WriteFile(self->hSerial, buffer, strlen(buffer), &bytes_written, NULL);
    FlushFileBuffers(self->hSerial);
}

static PyObject *Kmbox_move(KmboxObject *self, PyObject *args) {
    int x, y;
    if (!PyArg_ParseTuple(args, "ii", &x, &y))
        return NULL;

    char cmd[64];
    snprintf(cmd, sizeof(cmd), "km.move(%d,%d)", x, y);
    send_command(self, cmd);
    Py_RETURN_NONE;
}

static PyObject *Kmbox_left_click(KmboxObject *self, PyObject *Py_UNUSED(ignored)) {
    send_command(self, "km.click(0)");
    Py_RETURN_NONE;
}

static PyObject *Kmbox_right_click(KmboxObject *self, PyObject *Py_UNUSED(ignored)) {
    send_command(self, "km.click(1)");
    Py_RETURN_NONE;
}

static PyObject *Kmbox_middle_click(KmboxObject *self, PyObject *Py_UNUSED(ignored)) {
    send_command(self, "km.click(2)");
    Py_RETURN_NONE;
}

static PyObject *Kmbox_is_connected(KmboxObject *self, PyObject *Py_UNUSED(ignored)) {
    return PyBool_FromLong(self->is_connected);
}

static PyObject *Kmbox_close(KmboxObject *self, PyObject *Py_UNUSED(ignored)) {
    if (self->is_connected) {
        CloseHandle(self->hSerial);
        self->is_connected = 0;
        if (self->debug)
            printf("[INFO] Connection closed\n");
    }
    Py_RETURN_NONE;
}

static void Kmbox_dealloc(KmboxObject *self) {
    if (self->is_connected) {
        CloseHandle(self->hSerial);
    }
    Py_TYPE(self)->tp_free((PyObject *)self);
}

static PyMethodDef Kmbox_methods[] = {
    {"move", (PyCFunction)Kmbox_move, METH_VARARGS, "Move the mouse"},
    {"left_click", (PyCFunction)Kmbox_left_click, METH_NOARGS, "Left click"},
    {"right_click", (PyCFunction)Kmbox_right_click, METH_NOARGS, "Right click"},
    {"middle_click", (PyCFunction)Kmbox_middle_click, METH_NOARGS, "Middle click"},
    {"is_connected", (PyCFunction)Kmbox_is_connected, METH_NOARGS, "Check if connected"},
    {"close", (PyCFunction)Kmbox_close, METH_NOARGS, "Close the connection"},
    {NULL}
};

static PyTypeObject KmboxType = {
    PyVarObject_HEAD_INIT(NULL, 0)
    .tp_name = "kmbox.Kmbox",
    .tp_basicsize = sizeof(KmboxObject),
    .tp_itemsize = 0,
    .tp_flags = Py_TPFLAGS_DEFAULT,
    .tp_new = PyType_GenericNew,
    .tp_init = (initproc)Kmbox_init,
    .tp_dealloc = (destructor)Kmbox_dealloc,
    .tp_methods = Kmbox_methods,
};

static PyModuleDef kmboxmodule = {
    PyModuleDef_HEAD_INIT,
    .m_name = "kmbox",
    .m_doc = "Fast KMBox serial communication module.",
    .m_size = -1,
};

PyMODINIT_FUNC PyInit_kmbox(void) {
    PyObject *m;
    if (PyType_Ready(&KmboxType) < 0)
        return NULL;

    m = PyModule_Create(&kmboxmodule);
    if (!m)
        return NULL;

    Py_INCREF(&KmboxType);
    PyModule_AddObject(m, "Kmbox", (PyObject *)&KmboxType);
    return m;
}
