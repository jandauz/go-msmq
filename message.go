package msmq

import (
	"fmt"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type Message struct {
	dispatch *ole.IDispatch
}

func NewMessage() (Message, error) {
	unknown, err := oleutil.CreateObject("MSMQ.MSMQMessage")
	if err != nil && err.Error() == "Invalid class string" {
		return Message{}, ErrMSMQNotInstalled
	}

	dispatch, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return Message{}, err
	}

	msg := Message{
		dispatch: dispatch,
	}

	return msg, nil
}

// Send sends a message to the queue. An option can be specified to indicate
// whether the message is sent as a transaction.
func (m *Message) Send(queue *Queue, opts ...SendOption) error {
	options := &sendOptions{
		level: MTS,
	}
	for _, o := range opts {
		o.set(options)
	}

	_, err := m.dispatch.CallMethod("Send", queue.dispatch, int(options.level))
	if err != nil {
		return fmt.Errorf("go-msmq: Send() failed to send message: %w", err)
	}

	return nil
}

// SendOption represents an option to send messages to a queue.
type SendOption struct {
	set func(o *sendOptions)
}

// sendOptions contains all the options to send messages to a queue.
type sendOptions struct {
	level TransactionLevel
}

// SendWithTransaction returns a SendOption that configures sending messages
// to a queue with the specified level value.
//
// The default is MTS.
func SendWithTransaction(level TransactionLevel) SendOption {
	return SendOption{
		set: func(o *sendOptions) {
			o.level = level
		},
	}
}

func (m *Message) Body() (string, error) {
	res, err := m.dispatch.GetProperty("Body")
	if err != nil {
		return "", err
	}

	switch {
	// Applications using win32 API to communicate with MSMQ set message
	// body type to VT_EMPTY by default. The COM implementation interprets
	// this as an array of bytes. Since go-ole.VARIANT.Value() does not
	// support array of bytes, we need to include a check to see if the
	// variant type contains the VT_ARRAY bit flag, and if it does we
	// first convert to SafeArray and then to byte array.
	//
	// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/msmq/ms701459%28v%3dvs.85%29
	case res.VT&ole.VT_ARRAY != 0:
		return string(res.ToArray().ToByteArray()), nil
	default:
		return res.Value().(string), nil
	}
}

func (m *Message) SetBody(s string) error {
	_, err := m.dispatch.PutProperty("Body", s)
	if err != nil {
		return err
	}

	return nil
}
