package msmq_test

import (
	"testing"

	"github.com/jandauz/go-msmq"
)

func TestMessage_Send(t *testing.T) {
	queueInfo, err := msmq.NewQueueInfo()
	if err != nil {
		t.Errorf("NewQueueInfo() returned unexpected error: %v", err)
	}

	const path = `DIRECT=OS:.\private$\go-msmq`
	err = queueInfo.SetFormatName(path)
	if err != nil {
		t.Errorf("SetFormatName(%s) returned unexpected error: %v", path, err)
	}

	sendQueue, err := queueInfo.Open(msmq.Send, msmq.DenyNone)
	if err != nil {
		t.Errorf("Open(%v, %v) returned unexpected error: %v", msmq.Receive, msmq.DenyNone, err)
	}

	msg, err := msmq.NewMessage()
	if err != nil {
		t.Errorf("NewMessage() returned unexpected error: %v", err)
	}

	want := "Hello"
	err = msg.SetBody(want)
	if err != nil {
		t.Errorf("SetBody(%s) returned unexpected error: %v", want, err)
	}

	err = msg.Send(sendQueue)
	if err != nil {
		t.Errorf("Send(%+v) returned unexpected error: %v", sendQueue, err)
	}

	receiveQueue, err := queueInfo.Open(msmq.Receive, msmq.DenyNone)
	if err != nil {
		t.Errorf("Open(%v, %v) returned unexpected error: %v", msmq.Receive, msmq.DenyNone, err)
	}

	msg, err = receiveQueue.Receive()
	if err != nil {
		t.Errorf("Receive() returned unexpected error: %v", err)
	}

	got, err := msg.Body()
	if err != nil {
		t.Errorf("Body() returned unexpected error: %v", err)
	}

	if got != want {
		t.Errorf("got: %s, want: %s", got, want)
	}
}
