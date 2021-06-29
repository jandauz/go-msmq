// +build windows

package msmq

import (
	"errors"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type QueueInfo struct {
	dispatch *ole.IDispatch
}

func NewQueueInfo() (*QueueInfo, error) {
	unknown, err := oleutil.CreateObject("MSMQ.MSMQQueueInfo")
	if err != nil && err.Error() == "Invalid class string" {
		return nil, ErrMSMQNotInstalled
	}

	dispatch, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return nil, err
	}

	return &QueueInfo{
		dispatch: dispatch,
	}, nil
}

var ErrMSMQNotInstalled = errors.New("msmq: message queuing has not been installed on this computer")

func (qi *QueueInfo) Open(accessMode AccessMode, shareMode ShareMode) (*Queue, error) {
	queue, err := qi.dispatch.CallMethod("Open", int(accessMode), int(shareMode))
	if err != nil {
		return nil, err
	}

	return &Queue{
		dispatch: queue.ToIDispatch(),
	}, nil
}

// AccessMode defines access modes for accessing messages within a queue.
type AccessMode int

const (
	// Receive grants permissions to read, peek, and delete messages from a local queue.
	Receive = 1

	// Send grants permissions to insert new messages into a queue.
	Send = 2

	// Peek grants permissions to peek but not delete messages from a local queue.
	Peek = 32

	// admin specifies that a remote queue is to be opened.
	admin = 128

	// PeekAndAdmin grants Peek permissions to a remote queue.
	PeekAndAdmin = Peek | admin

	// ReceiveAndAdmin grants Receive permissions to a remote queue.
	ReceiveAndAdmin = Receive | admin
)

// ShareMode defines the exclusivity level when accessing a queue. Default
// value is DenyNone.
type ShareMode int

const (
	// DenyNone indicates that accessing a queue is available to all members
	// of the EVERYONE group.
	DenyNone = 0

	// DenyReceive limits access to other processes.
	DenyReceive = 1
)

func (qi *QueueInfo) FormatName() (string, error) {
	res, err := qi.dispatch.GetProperty("FormatName")
	if err != nil {
		return "", err
	}

	return res.Value().(string), nil
}

func (qi *QueueInfo) SetFormatName(name string) error {
	_, err := qi.dispatch.PutProperty("FormatName", name)
	if err != nil {
		return err
	}

	return nil
}
