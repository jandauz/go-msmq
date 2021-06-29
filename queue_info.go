// +build windows

package msmq

import (
	"errors"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// QueueInfo provides queue management for a single queue. It provides methods
// for creating a queue (either a transactional or non-transactional queue),
// opening a queue, changing or retrieving properties of a queue, and deleting
// a queue.
type QueueInfo struct {
	dispatch *ole.IDispatch
}

// NewQueueInfo returns a pointer to a QueueInfo. The FormatName or PathName
// must be set before interacting with a queue.
// This can be done through options:
//   queueInfo, err := msmq.NewQueueInfo(msmq.WithFormatName(name))
// Alternatively, it can be done through the QueueInfo.SetFormatName() function:
//   err := queueInfo.SetFormatName(name)
func NewQueueInfo(opts ...QueueInfoOption) (*QueueInfo, error) {
	options := &queueInfoOptions{}
	for _, o := range opts {
		o.set(options)
	}

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

// QueueInfoOption represents an option to configure QueueInfo.
type QueueInfoOption struct {
	set func(opts *queueInfoOptions)
}

// queueInfoOptions contains all the options to configure QueueInfo.
type queueInfoOptions struct {
	formatName string
	pathName   string
}

// WithFormatName returns a QueueInfoOption that configures QueueInfo with the
// specified format name.
func WithFormatName(name string) QueueInfoOption {
	return QueueInfoOption{
		set: func(opts *queueInfoOptions) {
			opts.formatName = name
		},
	}
}

// WithPathName returns a QueueInfoOption that configures QueueInfo with the
// specified path name.
func WithPathName(name string) QueueInfoOption {
	return QueueInfoOption{
		set: func(opts *queueInfoOptions) {
			opts.pathName = name
		},
	}
}

// ErrMSMQNotInstalled is returned when trying to interact with MSMQ but it is
// not installed.
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

// AccessMode defines access modes for accessing messages within a queue. The
// access mode cannot be changed while a queue is open.
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

// FormatName returns the format name.
func (qi *QueueInfo) FormatName() (string, error) {
	res, err := qi.dispatch.GetProperty("FormatName")
	if err != nil {
		return "", err
	}

	return res.Value().(string), nil
}

// SetFormatName sets the format name. Format names are used to reference public
// or private queues without accessing directory service.
func (qi *QueueInfo) SetFormatName(name string) error {
	_, err := qi.dispatch.PutProperty("FormatName", name)
	if err != nil {
		return err
	}

	return nil
}

// PathName returns the path name.
func (qi *QueueInfo) PathName() (string, error) {
	res, err := qi.dispatch.GetProperty("PathName")
	if err != nil {
		return "", err
	}

	return res.Value().(string), nil
}

// SetPathName sets the path name which specifies the name of the computer where
// the messages in the queue will be stored, an optional PRIVATE$ keyword that
// indicates whether the queue is a private queue, and the name of the queue.
//
// Path name syntax can be any of:
//   ComputerName\QueueName
//   ComputerName\PRIVATE$\QueueName
//   .\QueueName
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms706083(v=vs.85)
func (qi *QueueInfo) SetPathName(name string) error {
	_, err := qi.dispatch.PutProperty("PathName", name)
	if err != nil {
		return err
	}

	return nil
}
