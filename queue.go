// +build windows

package msmq

import (
	"errors"
	"fmt"

	"github.com/go-ole/go-ole"
)

// Queue represents an instance of a queue that is represented by
// QueueInfo. It provides the methods needed read and delete the
// messages in the queue and the properties needed to manage the open
// queue.
type Queue struct {
	dispatch *ole.IDispatch
}

// Close closes this queue.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms705220(v=vs.85)
func (q *Queue) Close() error {
	_, err := q.dispatch.CallMethod("Close")
	if err != nil {
		return fmt.Errorf("msmq: Close() failed to close queue: %w", err)
	}

	return nil
}

// Peek returns the first message in the queue, or waits for a message to arrive
// if the queue is empty. It does not remove the message from the queue.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms704311(v=vs.85)
func (q *Queue) Peek(opts ...PeekOption) (Message, error) {
	msg, err := q.peek("Peek", opts)
	if err != nil {
		return Message{}, err
	}

	return Message{
		dispatch: msg.ToIDispatch(),
	}, nil
}

// PeekOption represents an option to peek messages in a queue.
type PeekOption struct {
	set func(opts *peekOptions)
}

// peekOptions contains all the options to peek messages in a queue.
type peekOptions struct {
	wantDestinationQueue bool
	wantBody             bool
	timeout              int
	wantConnectorType    bool
}

// PeekWithWantDestinationQueue returns a PeekOption that configures peeking
// message with the specified want value.
//
// The default is false. If set to true, the Message.DestinationQueueInfo
// property is updated when the message is read from the queue. Setting this
// option to true may slow down the operation.
func PeekWithWantDestinationQueue(want bool) PeekOption {
	return PeekOption{
		set: func(opts *peekOptions) {
			opts.wantDestinationQueue = want
		},
	}
}

// PeekWithWantBody returns a PeekOption that configures peeking messages with
// the specified want value.
//
// The default is true. It specifies that the body of the message should be
// retrieved. If the message body is not needed, set this option to false to
// optimize the speed of the application.
func PeekWithWantBody(want bool) PeekOption {
	return PeekOption{
		set: func(opts *peekOptions) {
			opts.wantBody = want
		},
	}
}

// PeekWithTimeout returns a PeekOption that configures peeking messages with
// the specified timeout value.
//
// The default is infinite (max value of int). It specifies the time in
// milliseconds that MSMQ will wait for a message to arrive.
func PeekWithTimeout(timeout int) PeekOption {
	return PeekOption{
		set: func(opts *peekOptions) {
			opts.timeout = timeout
		},
	}
}

// PeekWithWantConnectorType returns a PeekOption that configures peeking
// messages with the specified want value.
//
// The default is false. It specifies that MSMQ does not retrieve the
// Message.ConnectorTypeGuid property when it peeks at a message in the
// queue
func PeekWithWantConnectorType(want bool) PeekOption {
	return PeekOption{
		set: func(opts *peekOptions) {
			opts.wantConnectorType = want
		},
	}
}

// PeekByLookupID returns the message referenced by id but does not remove the
// message from the queue.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms699797(v=vs.85)
func (q *Queue) PeekByLookupID(id uint64, opts ...PeekByLookupIDOption) (Message, error) {
	msg, err := q.peek("PeekByLookupID", id, opts)
	if err != nil {
		return Message{}, fmt.Errorf("go-msmq: PeekByLookupID(%d) failed to peek message by lookup id: %w", id, err)
	}

	return Message{
		dispatch: msg.ToIDispatch(),
	}, nil
}

// PeekByLookupIDOption represents an option to peek messages in a queue.
type PeekByLookupIDOption struct {
	set func(opts *peekByLookupIDOptions)
}

// peekOptions contains all the options to peek messages in a queue.
type peekByLookupIDOptions struct {
	wantDestinationQueue bool
	wantBody             bool
	wantConnectorType    bool
}

// PeekByLookupIDWithWantDestinationQueue returns a PeekOption that configures peeking
// message with the specified want value.
//
// The default is false. If set to true, the Message.DestinationQueueInfo
// property is updated when the message is read from the queue. Setting this
// option to true may slow down the operation.
func PeekByLookupIDWithWantDestinationQueue(want bool) PeekByLookupIDOption {
	return PeekByLookupIDOption{
		set: func(opts *peekByLookupIDOptions) {
			opts.wantDestinationQueue = want
		},
	}
}

// PeekByLookupIDWithWantBody returns a PeekOption that configures peeking messages with
// the specified want value.
//
// The default is true. It specifies that the body of the message should be
// retrieved. If the message body is not needed, set this option to false to
// optimize the speed of the application.
func PeekByLookupIDWithWantBody(want bool) PeekByLookupIDOption {
	return PeekByLookupIDOption{
		set: func(opts *peekByLookupIDOptions) {
			opts.wantBody = want
		},
	}
}

// PeekByLookupIDWithWantConnectorType returns a PeekOption that configures peeking
// messages with the specified want value.
//
// The default is false. It specifies that MSMQ does not retrieve the
// Message.ConnectorTypeGuid property when it peeks at a message in the
// queue
func PeekByLookupIDWithWantConnectorType(want bool) PeekByLookupIDOption {
	return PeekByLookupIDOption{
		set: func(opts *peekByLookupIDOptions) {
			opts.wantConnectorType = want
		},
	}
}

// PeekCurrent returns the message at the current cursor position and moves the
// cursor to the next message, or waits for a message to arrive, but does not
// remove the message from the queue. If the cursor does not point to a specific
// message location, PeekCurrent moves the cursor to the front of the queue.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms706182(v=vs.85)
func (q *Queue) PeekCurrent(opts ...PeekOption) (Message, error) {
	msg, err := q.peek("PeekCurrent", opts)
	if err != nil {
		return Message{}, fmt.Errorf("go-msmq: PeekCurrent() failed to peek current message: %w", err)
	}

	return Message{
		dispatch: msg.ToIDispatch(),
	}, nil
}

// PeekFirstByLookupID returns the first message in the queue without removing
// the message from the queue.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms711410(v=vs.85)
func (q *Queue) PeekFirstByLookupID(opts ...PeekByLookupIDOption) (Message, error) {
	msg, err := q.peek("PeekFirstByLookupID", opts)
	if err != nil {
		return Message{}, fmt.Errorf("go-msmq: PeekFirstByLookupID() failed to peek first message by lookup id: %w", err)
	}

	return Message{
		dispatch: msg.ToIDispatch(),
	}, nil
}

// PeekLastByLookupID returns the last message in the queue without removing
// the message from the queue.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms705194(v=vs.85)
func (q *Queue) PeekLastByLookupID(opts ...PeekByLookupIDOption) (Message, error) {
	msg, err := q.peek("PeekLastByLookupID", opts)
	if err != nil {
		return Message{}, fmt.Errorf("go-msmq: PeekLastByLookupID() failed to peek last message by lookup id: %w", err)
	}

	return Message{
		dispatch: msg.ToIDispatch(),
	}, nil
}

// PeekNext returns the message after the current cursor position or waits for a
// message to arrive, but does not remove the message from the queue.
//
// PeekNext moves the cursor first and then looks at the message at the new
// location. PeekNext must be called before PeekCurrent.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms703247(v=vs.85)
func (q *Queue) PeekNext(opts ...PeekOption) (Message, error) {
	msg, err := q.peek("PeekNext", opts)
	if err != nil {
		return Message{}, fmt.Errorf("go-msmq: failed to peek next message: %w", err)
	}

	return Message{
		dispatch: msg.ToIDispatch(),
	}, nil
}

// PeekNextByLookupID returns the message that follows the message referenced
// by id but does not remove the message from the queue.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms706024(v=vs.85)
func (q *Queue) PeekNextByLookupID(id uint64, opts ...PeekByLookupIDOption) (Message, error) {
	msg, err := q.peek("PeekNextByLookupID", id, opts)
	if err != nil {
		return Message{}, fmt.Errorf("go-msmq: PeekNextByLookupID(%d) failed to peek next message by lookup id: %w", id, err)
	}

	return Message{
		dispatch: msg.ToIDispatch(),
	}, nil
}

// PeekPreviousByLookupID returns the message that follows the message referenced
// by id but does not remove the message from the queue.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms706024(v=vs.85)
func (q *Queue) PeekPreviousByLookupID(id uint64, opts ...PeekByLookupIDOption) (Message, error) {
	msg, err := q.peek("PeekPreviousByLookupID", id, opts)
	if err != nil {
		return Message{}, fmt.Errorf("go-msmq: PeekPreviousByLookupID(%d) failed to peek previous message by lookup id: %w", id, err)
	}

	return Message{
		dispatch: msg.ToIDispatch(),
	}, nil
}

func (q *Queue) peek(action string, params ...interface{}) (*ole.VARIANT, error) {
	open, err := q.IsOpen()
	if err != nil {
		return nil, err
	}

	if !open {
		return nil, errors.New("Exception occurred. (The queue is not open or might not exist. )")
	}

	switch action {
	case "Peek", "PeekCurrent", "PeekNext":
		options := &peekOptions{
			wantDestinationQueue: false,
			wantBody:             true,
			timeout:              1<<31 - 1,
			wantConnectorType:    false,
		}

		for _, o := range params[0].([]PeekOption) {
			o.set(options)
		}

		return q.dispatch.CallMethod(action, options.wantDestinationQueue, options.wantBody, options.timeout, options.wantConnectorType)

	case "PeekByLookupID", "PeekNextByLookupID", "PeekPreviousByLookupID":
		id := params[0].(uint64)
		options := &peekByLookupIDOptions{
			wantDestinationQueue: false,
			wantBody:             true,
			wantConnectorType:    false,
		}

		for _, o := range params[1].([]PeekByLookupIDOption) {
			o.set(options)
		}

		return q.dispatch.CallMethod(action, id, options.wantDestinationQueue, options.wantBody, options.wantConnectorType)

	case "PeekFirstByLookupID", "PeekLastByLookupID":
		options := &peekByLookupIDOptions{
			wantDestinationQueue: false,
			wantBody:             true,
			wantConnectorType:    false,
		}

		for _, o := range params[0].([]PeekByLookupIDOption) {
			o.set(options)
		}

		return q.dispatch.CallMethod(action, options.wantDestinationQueue, options.wantBody, options.wantConnectorType)

	default:
		return nil, nil
	}
}

// Purge deletes all the messages in the queue. The queue must be opened with
// Receive AccessMode in order to purge messages.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms703966(v=vs.85)
func (q *Queue) Purge() error {
	open, err := q.IsOpen()
	if err != nil {
		return fmt.Errorf("go-msmq: failed to purge messages: %w", err)
	}

	if !open {
		return fmt.Errorf("go-msmq: failed to purge messages: %w", errors.New("Exception occurred. (The queue is not open or might not exist. )"))
	}

	_, err = q.dispatch.CallMethod("Purge")
	if err != nil {
		return fmt.Errorf("go-msmq: Purge() failed to delete all messages: %w", err)
	}

	return nil
}

// Receive retrieves the first message in the queue, removing the message from
// the queue when the message is read. It does not use the cursor created when
// the queue is opened, and should not be called when navigating the queue
// using the cursor.
//
// If no message is found, Receive will block until a message arrives in the
// queue or the timeout specified has expired.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms706017(v=vs.85)
func (q *Queue) Receive(opts ...ReceiveOption) (Message, error) {
	msg, err := q.receive("Receive", opts)
	if err != nil {
		return Message{}, err
	}

	return Message{
		dispatch: msg.ToIDispatch(),
	}, nil
}

// ReceiveOption represents an option to receive messages from a queue.
type ReceiveOption struct {
	set (func(o *receiveOptions))
}

// receiveOptions contains all the options to receive messages from a queue.
type receiveOptions struct {
	level                TransactionLevel
	wantDestinationQueue bool
	wantBody             bool
	timeout              int
	wantConnectorType    bool
}

// ReceiveWithTransaction returns a ReceiveOption that configures receiving
// messages from a queue with the specified level value.
//
// The default is MTS.
func ReceiveWithTransaction(level TransactionLevel) ReceiveOption {
	return ReceiveOption{
		set: func(o *receiveOptions) {
			o.level = level
		},
	}
}

// ReceiveWithWantDestinationQueue returns a ReceiveOption that configures receiving
// messages from a queue with the specified want value.
//
// The default is false. If set to true, the Message.DestinationQueueInfo
// property is updated when the message is read from the queue. Setting this
// option to true may slow down the operation.
func ReceiveWithWantDestinationQueue(want bool) ReceiveOption {
	return ReceiveOption{
		set: func(opts *receiveOptions) {
			opts.wantDestinationQueue = want
		},
	}
}

// ReceiveWithWantBody returns a ReceiveOption that configures receiving
// messages from a queue with the specified want value.
//
// The default is true. It specifies that the body of the message should be
// retrieved. If the message body is not needed, set this option to false to
// optimize the speed of the application.
func ReceiveWithWantBody(want bool) ReceiveOption {
	return ReceiveOption{
		set: func(opts *receiveOptions) {
			opts.wantBody = want
		},
	}
}

// ReceiveWithTimeout returns a ReceiveOption that configures receiving messages
// with the specified timeout value.
//
// The default is infinite (max value of int). It specifies the time in
// milliseconds that MSMQ will wait for a message to arrive.
func ReceiveWithTimeout(timeout int) ReceiveOption {
	return ReceiveOption{
		set: func(opts *receiveOptions) {
			opts.timeout = timeout
		},
	}
}

// ReceiveWithWantConnectorType returns a ReceiveOption that configures receiving
// messages with the specified want value.
//
// The default is false. It specifies that MSMQ does not retrieve the
// Message.ConnectorTypeGuid property when it receives a message in the
// queue.
func ReceiveWithWantConnectorType(want bool) ReceiveOption {
	return ReceiveOption{
		set: func(opts *receiveOptions) {
			opts.wantConnectorType = want
		},
	}
}

// ReceiveByLookupID returns the message referenced by id and removes the message
// from the queue.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms701233(v=vs.85)
func (q *Queue) ReceiveByLookupID(id uint64, opts ...ReceiveByLookupIDOption) (Message, error) {
	msg, err := q.receive("ReceiveByLookupID", id, opts)
	if err != nil {
		return Message{}, fmt.Errorf("go-msmq: ReceiveByLookupID(%d) failed to receive messages by lookup id: %w", id, err)
	}

	return Message{
		dispatch: msg.ToIDispatch(),
	}, nil
}

// ReceiveByLookupIDOption represents an option to receive messages by lookup
// ID in a queue.
type ReceiveByLookupIDOption struct {
	set func(o *receiveByLookupIDOptions)
}

// receiveByLookupIDOptions contains all the options to receive messages by
// lookup ID in a queue.
type receiveByLookupIDOptions struct {
	level                TransactionLevel
	wantDestinationQueue bool
	wantBody             bool
	wantConnectorType    bool
}

// ReceiveByLookupIDWithTransaction returns a ReceiveOption that configures
// receiving messages by lookup ID from a queue with the specified level value.
//
// The default is MTS.
func ReceiveByLookupIDWithTransaction(level TransactionLevel) ReceiveByLookupIDOption {
	return ReceiveByLookupIDOption{
		set: func(o *receiveByLookupIDOptions) {
			o.level = level
		},
	}
}

// ReceiveByLookupIDWithWantDestinationQueue returns a ReceiveByLookupIDOption
// that configures receiving a message by lookup ID with the specified want
// value.
//
// The default is false. If set to true, the Message.DestinationQueueInfo
// property is updated when the message is read from the queue. Setting this
// option to true may slow down the operation.
func ReceiveByLookupIDWithWantDestinationQueue(want bool) ReceiveByLookupIDOption {
	return ReceiveByLookupIDOption{
		set: func(o *receiveByLookupIDOptions) {
			o.wantDestinationQueue = want
		},
	}
}

// ReceiveByLookupIDWithWantBody returns a ReceiveByLookupIDOption that configures
// receiving messages by lookup ID with the specified want value.
//
// The default is true. It specifies that the body of the message should be
// retrieved. If the message body is not needed, set this option to false to
// optimize the speed of the application.
func ReceiveByLookupIDWithWantBody(want bool) ReceiveByLookupIDOption {
	return ReceiveByLookupIDOption{
		set: func(opts *receiveByLookupIDOptions) {
			opts.wantBody = want
		},
	}
}

// ReceiveByLookupIDWithWantConnectorType returns a ReceiveByLookupIDOption that
// configures receiving messages by lookup ID with the specified want value.
//
// The default is false. It specifies that MSMQ does not retrieve the
// Message.ConnectorTypeGuid property when it receives a message in the
// queue.
func ReceiveByLookupIDWithWantConnectorType(want bool) ReceiveByLookupIDOption {
	return ReceiveByLookupIDOption{
		set: func(opts *receiveByLookupIDOptions) {
			opts.wantConnectorType = want
		},
	}
}

// ReceiveCurrent returns the message at the current cursor position, removes the
// message from the queue, and moves the cursor to the next message, or waits for
// a message to arrive. If the cursor does not point to a specific message location,
// ReceiveCurrent moves the cursor to the front of the queue.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms706011(v=vs.85)
func (q *Queue) ReceiveCurrent(opts ...ReceiveOption) (Message, error) {
	msg, err := q.receive("ReceiveCurrent", opts)
	if err != nil {
		return Message{}, fmt.Errorf("go-msmq: ReceiveCurrent() failed to receive message at the current cursor location: %w", err)
	}

	return Message{
		dispatch: msg.ToIDispatch(),
	}, nil
}

// ReceiveFirstByLookupID returns the first message in the queue and removes
// the message from the queue.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms701869(v=vs.85)
func (q *Queue) ReceiveFirstByLookupID(opts ...ReceiveByLookupIDOption) (Message, error) {
	msg, err := q.receive("ReceiveFirstByLookupID", opts)
	if err != nil {
		return Message{}, fmt.Errorf("go-msmq: ReceiveFirstByLookupID() failed to receive first message by lookup id: %w", err)
	}

	return Message{
		dispatch: msg.ToIDispatch(),
	}, nil
}

// ReceiveLastByLookupID returns the last message in the queue and removes
// the message from the queue.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms706134(v=vs.85)
func (q *Queue) ReceiveLastByLookupID(opts ...ReceiveByLookupIDOption) (Message, error) {
	msg, err := q.receive("ReceiveLastByLookupID", opts)
	if err != nil {
		return Message{}, fmt.Errorf("go-msmq: ReceiveLastByLookupID() failed to receive last message by lookup id: %w", err)
	}

	return Message{
		dispatch: msg.ToIDispatch(),
	}, nil
}

func (q *Queue) receive(action string, params ...interface{}) (*ole.VARIANT, error) {
	open, err := q.IsOpen()
	if err != nil {
		return nil, err
	}

	if !open {
		return nil, errors.New("Exception occurred. (The queue is not open or might not exist. )")
	}

	switch action {
	case "Receive", "ReceiveCurrent":
		options := &receiveOptions{
			level:                MTS,
			wantDestinationQueue: false,
			wantBody:             true,
			timeout:              1<<31 - 1,
			wantConnectorType:    false,
		}

		for _, o := range params[0].([]ReceiveOption) {
			o.set(options)
		}

		return q.dispatch.CallMethod(action, int(options.level), options.wantDestinationQueue, options.wantBody, options.timeout, options.wantConnectorType)

	case "ReceiveByLookupID", "ReceiveNextByLookupID", "ReceivePreviousByLookupID":
		id := params[0].(uint64)
		options := &receiveByLookupIDOptions{
			level:                MTS,
			wantDestinationQueue: false,
			wantBody:             true,
			wantConnectorType:    false,
		}

		for _, o := range params[1].([]ReceiveByLookupIDOption) {
			o.set(options)
		}

		return q.dispatch.CallMethod(action, id, int(options.level), options.wantDestinationQueue, options.wantBody, options.wantConnectorType)

	case "ReceiveFirstByLookupID", "ReceiveLastByLookupID":
		options := &receiveByLookupIDOptions{
			level:                MTS,
			wantDestinationQueue: false,
			wantBody:             true,
			wantConnectorType:    false,
		}

		for _, o := range params[0].([]ReceiveByLookupIDOption) {
			o.set(options)
		}

		return q.dispatch.CallMethod(action, int(options.level), options.wantDestinationQueue, options.wantBody, options.wantConnectorType)

	default:
		return nil, nil
	}
}

func (q *Queue) IsOpen() (bool, error) {
	res, err := q.dispatch.GetProperty("IsOpen2")
	if err != nil {
		return false, fmt.Errorf("go-msmq: IsOpen() failed to get IsOpen2: %w", err)
	}

	return res.Value().(bool), err
}
