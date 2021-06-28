package msmq_test

import (
	"testing"

	"github.com/jandauz/go-msmq"
)

func TestQueueInfo_Open(t *testing.T) {
	queueInfo, err := msmq.NewQueueInfo()
	if err != nil {
		t.Errorf("NewQueueInfo() returned unexpected error: %v", err)
	}

	const path = `DIRECT=OS:.\private$\go-msmq`
	err = queueInfo.SetFormatName(path)
	if err != nil {
		t.Errorf("SetFormatName(%s) returned unexpected error: %v", path, err)
	}

	_, err = queueInfo.Open(msmq.Receive, msmq.DenyNone)
	if err != nil {
		t.Errorf("Open(%v, %v) returned unexpected error: %v", msmq.Receive, msmq.DenyNone, err)
	}
}

func TestQueueInfo_FormatName(t *testing.T) {
	queueInfo, err := msmq.NewQueueInfo()
	if err != nil {
		t.Errorf("NewQueueInfo() returned unexpected error: %v", err)
	}

	const want = `DIRECT=OS:.\private$\go-msmq`
	err = queueInfo.SetFormatName(want)
	if err != nil {
		t.Errorf("SetFormatName(%s) returned unexpected error: %v", want, err)
	}

	got, err := queueInfo.FormatName()
	if err != nil {
		t.Errorf("FormatName() returned unexpected error: %v", err)
	}

	if got != want {
		t.Errorf("got: %s, want: %s", got, want)
	}
}

// func TestDial(t *testing.T) {
// 	const path = `DIRECT=OS:.\private$\go-msmq`
// 	opts := msmq.Options{
// 		AccessMode: msmq.Peek,
// 		ShareMode:  msmq.DenyNone,
// 	}
// 	_, err := msmq.Open(path, opts)
// 	if err != nil {
// 		t.Errorf("Dial(%s, %+v) returned err: %v", path, opts, err)
// 	}
// }

// func TestPeek(t *testing.T) {
// 	const path = `DIRECT=OS:.\private$\go-msmq`
// 	opts := msmq.Options{
// 		AccessMode: msmq.Peek,
// 		ShareMode:  msmq.DenyNone,
// 	}
// 	queue, err := msmq.Open(path, opts)
// 	if err != nil {
// 		t.Errorf("Dial(%s, %+v) returned error: %v", path, opts, err)
// 	}

// 	_, err = queue.Peek()
// 	if err != nil {
// 		t.Errorf("Peek() returned error: %v", err)
// 	}
// }

// func TestSend(t *testing.T) {
// 	const path = `DIRECT=OS:.\private$\go-msmq`
// 	opts := msmq.Options{
// 		AccessMode: msmq.Send,
// 		ShareMode:  msmq.DenyNone,
// 	}
// 	queue, err := msmq.Open(path, opts)
// 	if err != nil {
// 		t.Errorf("Dial(%s, %+v) returned error: %v", path, opts, err)
// 	}

// 	s := "Hello"
// 	err = queue.Send(s)
// 	if err != nil {
// 		t.Errorf("Send(%s) returned error: %v", s, err)
// 	}
// }
