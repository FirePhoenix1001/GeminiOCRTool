import ctypes
from ctypes import wintypes
import sys

# Windows Constants
GWL_WNDPROC = -4
WM_DROPFILES = 0x0233

shell32 = ctypes.windll.shell32
user32 = ctypes.windll.user32

# Set function signatures for CallWindowProcW and Get/SetWindowLongPtrW to avoid type size conflicts on 64-bit
user32.CallWindowProcW.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint, ctypes.c_void_p, ctypes.c_void_p]
user32.CallWindowProcW.restype = ctypes.c_void_p

if sys.maxsize > 2**32:
    GetWindowLong = user32.GetWindowLongPtrW
    GetWindowLong.argtypes = [ctypes.c_void_p, ctypes.c_int]
    GetWindowLong.restype = ctypes.c_void_p
    
    SetWindowLong = user32.SetWindowLongPtrW
    SetWindowLong.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_void_p]
    SetWindowLong.restype = ctypes.c_void_p
else:
    GetWindowLong = user32.GetWindowLongW
    GetWindowLong.argtypes = [ctypes.c_void_p, ctypes.c_int]
    GetWindowLong.restype = ctypes.c_long
    
    SetWindowLong = user32.SetWindowLongW
    SetWindowLong.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_long]
    SetWindowLong.restype = ctypes.c_long

WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint, ctypes.c_void_p, ctypes.c_void_p)

# Global trackers to keep Python objects from being garbage collected
_callbacks = {}
_old_wndprocs = {}
_wndproc_refs = {}

def hook_dropfiles(hwnd, callback):
    """
    Hooks the WM_DROPFILES message of the window with the given HWND.
    When a file is dropped, callback(files) is called with a list of absolute file paths (strings).
    """
    hwnd = int(hwnd)
    if hwnd in _callbacks:
        _callbacks[hwnd] = callback
        return True

    # Enable drag acceptance for this window
    shell32.DragAcceptFiles(ctypes.c_void_p(hwnd), True)

    # Retrieve the old window procedure
    old_wndproc = GetWindowLong(ctypes.c_void_p(hwnd), GWL_WNDPROC)
    if not old_wndproc:
        return False

    _old_wndprocs[hwnd] = old_wndproc
    _callbacks[hwnd] = callback

    def new_wndproc(hwnd_val, msg, wp, lp):
        if msg == WM_DROPFILES:
            h_drop = wintypes.HANDLE(wp)
            # Query count of dropped files (-1 returns the count)
            count = shell32.DragQueryFileW(h_drop, -1, None, 0)
            files = []
            for i in range(count):
                # Query required buffer length (passing NULL returns length)
                length = shell32.DragQueryFileW(h_drop, i, None, 0)
                buf = ctypes.create_unicode_buffer(length + 1)
                # Retrieve the filename
                shell32.DragQueryFileW(h_drop, i, buf, len(buf))
                files.append(buf.value)
            
            # Finish drag operation and free memory
            shell32.DragFinish(h_drop)
            
            # Call the callback with the list of file paths
            if hwnd in _callbacks:
                try:
                    _callbacks[hwnd](files)
                except Exception as e:
                    print(f"Error in drag-drop callback: {e}")
            return 0

        # Pass other messages to the original window procedure
        return user32.CallWindowProcW(
            ctypes.c_void_p(_old_wndprocs[hwnd]), 
            ctypes.c_void_p(hwnd_val), 
            msg, 
            ctypes.c_void_p(wp), 
            ctypes.c_void_p(lp)
        )

    # Convert new_wndproc to ctypes WNDPROC type and store reference
    wndproc_cb = WNDPROC(new_wndproc)
    _wndproc_refs[hwnd] = wndproc_cb

    # Hook the window procedure
    SetWindowLong(ctypes.c_void_p(hwnd), GWL_WNDPROC, wndproc_cb)
    return True
