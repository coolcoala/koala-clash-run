//go:build windows
// +build windows

package main

import (
	"fmt"
	"runtime"
	"unsafe"

	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"golang.org/x/sys/windows"
)

var procIsProcessInJob = windows.NewLazySystemDLL("kernel32.dll").NewProc("IsProcessInJob")

// inJob reports whether this process runs in any job object.
// Task Scheduler puts every task it starts into the job of the tasks that are still running.
// If one of them belongs to an MSIX package (e.g. Office's opushutil.exe), the job is the
// package's job, and Windows terminates it when the package gets disabled during an update.
// Counting the processes in the job would see only the immediate job and only at this moment,
// and Koala gains nothing from any job, so it leaves whichever one it was started in.
func inJob() bool {
	var result int32
	ok, _, _ := procIsProcessInJob.Call(uintptr(windows.CurrentProcess()), 0,
		uintptr(unsafe.Pointer(&result)))
	return ok != 0 && result != 0
}

// startViaWMI starts cmdline through Win32_Process.Create: the process is created by the WMI
// service, so it is not placed into our job object.
func startViaWMI(cmdline string) error {
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()
	if err := ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED); err != nil {
		return err
	}
	defer ole.CoUninitialize()

	unknown, err := oleutil.CreateObject("WbemScripting.SWbemLocator")
	if err != nil {
		return err
	}
	defer unknown.Release()
	locator, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return err
	}
	defer locator.Release()

	service, err := oleutil.CallMethod(locator, "ConnectServer", ".", `root\cimv2`)
	if err != nil {
		return err
	}
	defer service.Clear()
	class, err := oleutil.CallMethod(service.ToIDispatch(), "Get", "Win32_Process")
	if err != nil {
		return err
	}
	defer class.Clear()
	result, err := oleutil.CallMethod(class.ToIDispatch(), "Create", cmdline)
	if err != nil {
		return err
	}
	defer result.Clear()
	if rc := result.Value(); rc != int32(0) {
		return fmt.Errorf("Win32_Process.Create returned %v", rc)
	}
	return nil
}
