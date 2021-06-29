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

	return res.Value().(string), nil
}

func (m *Message) SetBody(s string) error {
	_, err := m.dispatch.PutProperty("Body", s)
	if err != nil {
		return err
	}

	return nil
}
