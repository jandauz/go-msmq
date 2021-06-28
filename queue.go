package msmq

import "github.com/go-ole/go-ole"

type Queue struct {
	dispatch *ole.IDispatch
}

func (m *Queue) Peek() (Message, error) {
	msg, err := m.dispatch.CallMethod("Peek")
	if err != nil {
		return Message{}, err
	}

	return Message{
		dispatch: msg.ToIDispatch(),
	}, nil
}

func (m *Queue) Receive() (Message, error) {
	msg, err := m.dispatch.CallMethod("Receive")
	if err != nil {
		return Message{}, err
	}

	return Message{
		dispatch: msg.ToIDispatch(),
	}, nil
}
