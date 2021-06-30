// +build windows

package msmq

import (
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

func (q *Queue) Peek() (Message, error) {
	msg, err := q.dispatch.CallMethod("Peek")
	if err != nil {
		return Message{}, err
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
