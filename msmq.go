package msmq

import "github.com/go-ole/go-ole"

func init() {
	_ = ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED)
}
