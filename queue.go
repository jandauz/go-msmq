// +build windows

package msmq

import (
	"errors"
	"fmt"

	"github.com/go-ole/go-ole"
)

// Queue represents an open instance of a queue that is represented by
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
	open, err := q.IsOpen()
	if err != nil {
		return Message{}, fmt.Errorf("go-msmq: failed to peek message: %w", err)
	}

	if !open {
		return Message{}, fmt.Errorf("go-msmq: failed to peek message: %w", errors.New("Exception occurred. (The queue is not open or might not exist. )"))
	}

	options := &peekOptions{
		wantDestinationQueue: false,
		wantBody:             false,
		timeout:              1<<31 - 1,
		wantConnectorType:    false,
	}

	for _, o := range opts {
		o.set(options)
	}

	msg, err := q.dispatch.CallMethod(
		"Peek",
		options.wantDestinationQueue,
		options.wantBody,
		options.timeout,
		options.wantConnectorType)
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
	open, err := q.IsOpen()
	if err != nil {
		return Message{}, fmt.Errorf("go-msmq: failed to peek message by lookup id: %d: %w", id, err)
	}

	if !open {
		return Message{}, fmt.Errorf("go-msmq: failed to peek message by lookup id: %d: %w", id, errors.New("Exception occurred. (The queue is not open or might not exist. )"))
	}

	options := &peekByLookupIDOptions{
		wantDestinationQueue: false,
		wantBody:             true,
		wantConnectorType:    false,
	}

	for _, o := range opts {
		o.set(options)
	}

	msg, err := q.dispatch.CallMethod(
		"PeekByLookupId",
		id,
		options.wantDestinationQueue,
		options.wantBody,
		options.wantConnectorType,
	)
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

// PeekByLookupWithWantDestinationQueue returns a PeekOption that configures peeking
// message with the specified want value.
//
// The default is false. If set to true, the Message.DestinationQueueInfo
// property is updated when the message is read from the queue. Setting this
// option to true may slow down the operation.
func PeekByLookupWithWantDestinationQueue(want bool) PeekByLookupIDOption {
	return PeekByLookupIDOption{
		set: func(opts *peekByLookupIDOptions) {
			opts.wantDestinationQueue = want
		},
	}
}

// PeekByLookupWithWantBody returns a PeekOption that configures peeking messages with
// the specified want value.
//
// The default is true. It specifies that the body of the message should be
// retrieved. If the message body is not needed, set this option to false to
// optimize the speed of the application.
func PeekByLookupWithWantBody(want bool) PeekByLookupIDOption {
	return PeekByLookupIDOption{
		set: func(opts *peekByLookupIDOptions) {
			opts.wantBody = want
		},
	}
}

// PeekByLookupWithWantConnectorType returns a PeekOption that configures peeking
// messages with the specified want value.
//
// The default is false. It specifies that MSMQ does not retrieve the
// Message.ConnectorTypeGuid property when it peeks at a message in the
// queue
func PeekByLookupWithWantConnectorType(want bool) PeekByLookupIDOption {
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
	open, err := q.IsOpen()
	if err != nil {
		return Message{}, fmt.Errorf("go-msmq: failed to peek current message: %w", err)
	}

	if !open {
		return Message{}, fmt.Errorf("go-msmq: failed to peek current message: %w", errors.New("Exception occurred. (The queue is not open or might not exist. )"))
	}

	options := &peekOptions{
		wantDestinationQueue: false,
		wantBody:             false,
		timeout:              1<<31 - 1,
		wantConnectorType:    false,
	}

	for _, o := range opts {
		o.set(options)
	}

	msg, err := q.dispatch.CallMethod(
		"PeekCurrent",
		options.wantDestinationQueue,
		options.wantBody,
		options.timeout,
		options.wantConnectorType)
	if err != nil {
		return Message{}, err
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
	open, err := q.IsOpen()
	if err != nil {
		return Message{}, fmt.Errorf("go-msmq: failed to peek first message by lookup id: %w", err)
	}

	if !open {
		return Message{}, fmt.Errorf("go-msmq: failed to peek first message by lookup id: %w", errors.New("Exception occurred. (The queue is not open or might not exist. )"))
	}

	options := &peekByLookupIDOptions{
		wantDestinationQueue: false,
		wantBody:             true,
		wantConnectorType:    false,
	}
	for _, o := range opts {
		o.set(options)
	}

	msg, err := q.dispatch.CallMethod("PeekFirstByLookupId")
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
	open, err := q.IsOpen()
	if err != nil {
		return Message{}, fmt.Errorf("go-msmq: failed to peek last message by lookup id: %w", err)
	}

	if !open {
		return Message{}, fmt.Errorf("go-msmq: failed to peek last message by lookup id: %w", errors.New("Exception occurred. (The queue is not open or might not exist. )"))
	}

	options := &peekByLookupIDOptions{
		wantDestinationQueue: false,
		wantBody:             true,
		wantConnectorType:    false,
	}
	for _, o := range opts {
		o.set(options)
	}

	msg, err := q.dispatch.CallMethod("PeekLastByLookupID")
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
	open, err := q.IsOpen()
	if err != nil {
		return Message{}, fmt.Errorf("go-msmq: failed to peek next message: %w", err)
	}

	if !open {
		return Message{}, fmt.Errorf("go-msmq: failed to peek next message: %w", errors.New("Exception occurred. (The queue is not open or might not exist. )"))
	}

	options := &peekOptions{
		wantDestinationQueue: false,
		wantBody:             false,
		timeout:              1<<31 - 1,
		wantConnectorType:    false,
	}

	for _, o := range opts {
		o.set(options)
	}

	msg, err := q.dispatch.CallMethod(
		"PeekNext",
		options.wantDestinationQueue,
		options.wantBody,
		options.timeout,
		options.wantConnectorType)
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
	open, err := q.IsOpen()
	if err != nil {
		return Message{}, fmt.Errorf("go-msmq: failed to peek next message by lookup id: %d: %w", id, err)
	}

	if !open {
		return Message{}, fmt.Errorf("go-msmq: failed to peek next message by lookup id: %d: %w", id, errors.New("Exception occurred. (The queue is not open or might not exist. )"))
	}

	options := &peekByLookupIDOptions{
		wantDestinationQueue: false,
		wantBody:             true,
		wantConnectorType:    false,
	}

	for _, o := range opts {
		o.set(options)
	}

	msg, err := q.dispatch.CallMethod(
		"PeekNextByLookupID",
		id,
		options.wantDestinationQueue,
		options.wantBody,
		options.wantConnectorType,
	)
	if err != nil {
		return Message{}, fmt.Errorf("go-msmq: PeekNextByLookupID(%d) failed to peek next message by lookup id: %w", id, err)
	}

	return Message{
		dispatch: msg.ToIDispatch(),
	}, nil
}

func (q *Queue) Receive() (Message, error) {
	msg, err := q.dispatch.CallMethod("Receive")
	if err != nil {
		return Message{}, err
	}

	return Message{
		dispatch: msg.ToIDispatch(),
	}, nil
}

func (q *Queue) IsOpen() (bool, error) {
	res, err := q.dispatch.GetProperty("IsOpen2")
	if err != nil {
		return false, fmt.Errorf("go-msmq: IsOpen() failed to get IsOpen2: %w", err)
	}

	return res.Value().(bool), err
}
