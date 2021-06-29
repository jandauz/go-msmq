package msmq

import (
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

	return Message{
		dispatch: dispatch,
	}, nil
}

func (m *Message) Send(queue *Queue) error {
	_, err := m.dispatch.CallMethod("Send", queue.dispatch)
	if err != nil {
		return err
	}

	return nil
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
