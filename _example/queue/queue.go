package main

import (
	"log"
	"strconv"

	"github.com/jandauz/go-msmq"
)

func main() {
	// Creating QueueInfo
	opts := []msmq.QueueInfoOption{
		msmq.WithPathName(`.\private$\go-msmq`),
	}
	queueInfo, err := msmq.NewQueueInfo(opts...)
	if err != nil {
		log.Fatal(err)
	}

	// Create queue
	err = queueInfo.Create()
	if err != nil {
		log.Fatal(err)
	}

	// Refresh queue to retrieve auto-generated FormatName
	err = queueInfo.Refresh()
	if err != nil {
		log.Fatal(err)
	}

	// Open queue
	log.Println("Opening queue")
	queue, err := queueInfo.Open(msmq.Peek, msmq.DenyNone)
	if err != nil {
		log.Fatal(err)
	}
	b, err := queue.IsOpen()
	if err != nil {
		log.Fatal(err)
	}
	log.Printf("IsOpen: %v", b)

	// Close queue
	log.Println("Closing queue")
	err = queue.Close()
	if err != nil {
		log.Fatal(err)
	}
	b, err = queue.IsOpen()
	if err != nil {
		log.Fatal(err)
	}
	log.Printf("IsOpen: %v", b)

	// Peek
	{
		// Send message to queue
		sendMessages(queueInfo, msmq.NoTransaction)

		queue, err = queueInfo.Open(msmq.Peek, msmq.DenyNone)
		if err != nil {
			log.Fatal(err)
		}

		opts := []msmq.PeekOption{
			msmq.PeekWithWantDestinationQueue(true),
			msmq.PeekWithWantBody(true),
			msmq.PeekWithTimeout(1),
			msmq.PeekWithWantConnectorType(true),
		}
		msg, err := queue.Peek(opts...)
		if err != nil {
			log.Fatal(err)
		}
		s, err := msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Peek: %s", s)

		msg, err = queue.PeekCurrent(opts...)
		if err != nil {
			log.Fatal(err)
		}
		s, err = msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Peek current: %s", s)

		msg, err = queue.PeekNext(opts...)
		if err != nil {
			log.Fatal(err)
		}
		s, err = msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Peek next: %s", s)

		err = queue.Close()
		if err != nil {
			log.Fatal(err)
		}
	}

	// Peek by lookup id
	{
		queue, err = queueInfo.Open(msmq.Peek, msmq.DenyNone)
		if err != nil {
			log.Fatal(err)
		}
		msg, err := queue.Peek()
		if err != nil {
			log.Fatal(err)
		}
		s, err := msg.LookupID()
		if err != nil {
			log.Fatal(err)
		}
		id, err := strconv.ParseUint(s, 10, 64)
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Msg lookup id: %d", id)

		opts := []msmq.PeekByLookupIDOption{
			msmq.PeekByLookupIDWithWantDestinationQueue(true),
			msmq.PeekByLookupIDWithWantBody(true),
			msmq.PeekByLookupIDWithWantConnectorType(true),
		}
		msg, err = queue.PeekByLookupID(id, opts...)
		if err != nil {
			log.Fatal(err)
		}
		s, err = msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Peek by lookup id: %s", s)

		msg, err = queue.PeekFirstByLookupID(opts...)
		if err != nil {
			log.Fatal(err)
		}
		s, err = msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Peek first by lookup id: %s", s)

		msg, err = queue.PeekLastByLookupID(opts...)
		if err != nil {
			log.Fatal(err)
		}
		s, err = msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Peek last by lookup id: %s", s)

		msg, err = queue.PeekNextByLookupID(id, opts...)
		if err != nil {
			log.Fatal(err)
		}
		s, err = msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Peek next by lookup id: %s", s)

		s, err = msg.LookupID()
		if err != nil {
			log.Fatal(err)
		}
		id, err = strconv.ParseUint(s, 10, 64)
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Msg lookup id: %d", id)

		msg, err = queue.PeekPreviousByLookupID(id, opts...)
		if err != nil {
			log.Fatal(err)
		}
		s, err = msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Peek previous by lookup id: %s", s)

		err = queue.Close()
		if err != nil {
			log.Fatal(err)
		}
	}

	// Purge
	{
		queue, err = queueInfo.Open(msmq.Receive, msmq.DenyNone)
		if err != nil {
			log.Fatal(err)
		}

		err = queue.Purge()
		if err != nil {
			log.Fatal(err)
		}

		opts := []msmq.PeekOption{
			msmq.PeekWithWantDestinationQueue(true),
			msmq.PeekWithWantBody(true),
			msmq.PeekWithTimeout(1),
			msmq.PeekWithWantConnectorType(true),
		}
		msg, err := queue.Peek(opts...)
		if err != nil {
			log.Fatal(err)
		}

		if (msmq.Message{}) != msg {
			log.Fatal("Queue not purged")
		} else {
			log.Println("Queue is purged")
		}

		err = queue.Close()
		if err != nil {
			log.Fatal(err)
		}
	}

	// Receive
	{
		// Send message to queue
		sendMessages(queueInfo, msmq.MTS)

		queue, err = queueInfo.Open(msmq.Receive, msmq.DenyNone)
		if err != nil {
			log.Fatal(err)
		}

		opts := []msmq.ReceiveOption{
			msmq.ReceiveWithTransaction(msmq.NoTransaction),
			msmq.ReceiveWithWantDestinationQueue(true),
			msmq.ReceiveWithWantBody(true),
			msmq.ReceiveWithTimeout(1),
			msmq.ReceiveWithWantConnectorType(true),
		}
		msg, err := queue.Receive(opts...)
		if err != nil {
			log.Fatal(err)
		}
		s, err := msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Receive: %s", s)

		msg, err = queue.ReceiveCurrent(opts...)
		if err != nil {
			log.Fatal(err)
		}
		s, err = msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Receive current: %s", s)

		err = queue.Purge()
		if err != nil {
			log.Fatal(err)
		}

		err = queue.Close()
		if err != nil {
			log.Fatal(err)
		}
	}

	// Receive by look up id
	{
		// Send message to queue
		sendMessages(queueInfo, msmq.MTS)

		queue, err = queueInfo.Open(msmq.Peek, msmq.DenyNone)
		if err != nil {
			log.Fatal(err)
		}
		msg, err := queue.Peek()
		if err != nil {
			log.Fatal(err)
		}
		s, err := msg.LookupID()
		if err != nil {
			log.Fatal(err)
		}
		id, err := strconv.ParseUint(s, 10, 64)
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Msg lookup id: %d", id)

		err = queue.Close()
		if err != nil {
			log.Fatal(err)
		}

		queue, err = queueInfo.Open(msmq.Receive, msmq.DenyNone)
		if err != nil {
			log.Fatal(err)
		}

		opts := []msmq.ReceiveByLookupIDOption{
			msmq.ReceiveByLookupIDWithTransaction(msmq.NoTransaction),
			msmq.ReceiveByLookupIDWithWantDestinationQueue(true),
			msmq.ReceiveByLookupIDWithWantBody(true),
			msmq.ReceiveByLookupIDWithWantConnectorType(true),
		}
		msg, err = queue.ReceiveByLookupID(id, opts...)
		if err != nil {
			log.Fatal(err)
		}

		s, err = msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Receive by lookup id: %s", s)

		msg, err = queue.ReceiveFirstByLookupID(opts...)
		if err != nil {
			log.Fatal(err)
		}

		s, err = msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Receive by first by lookup id: %s", s)

		msg, err = queue.ReceiveLastByLookupID(opts...)
		if err != nil {
			log.Fatal(err)
		}

		s, err = msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Receive by last by lookup id: %s", s)

		msg, err = queue.PeekFirstByLookupID()
		if err != nil {
			log.Fatal(err)
		}

		s, err = msg.LookupID()
		if err != nil {
			log.Fatal(err)
		}
		id, err = strconv.ParseUint(s, 10, 64)
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Msg lookup id: %d", id)

		msg, err = queue.ReceiveNextByLookupID(id, opts...)
		if err != nil {
			log.Fatal(err)
		}

		s, err = msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Receive by next by lookup id: %s", s)

		msg, err = queue.PeekLastByLookupID()
		if err != nil {
			log.Fatal(err)
		}

		s, err = msg.LookupID()
		if err != nil {
			log.Fatal(err)
		}
		id, err = strconv.ParseUint(s, 10, 64)
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Msg lookup id: %d", id)

		msg, err = queue.ReceivePreviousByLookupID(id, opts...)
		if err != nil {
			log.Fatal(err)
		}

		s, err = msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Receive by previous by lookup id: %s", s)

		err = queue.Reset()
		if err != nil {
			log.Fatal(err)
		}

		msg, err = queue.Receive(msmq.ReceiveWithTimeout(1))
		if err != nil {
			log.Fatal(err)
		}

		s, err = msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Receive after reset: %s", s)

		err = queue.Close()
		if err != nil {
			log.Fatal(err)
		}
	}

	// Receive transactional
	{
		err = queueInfo.Delete()
		if err != nil {
			log.Fatal(err)
		}

		err = queueInfo.Create(msmq.CreateQueueWithTransactional(true))
		if err != nil {
			log.Fatal(err)
		}

		// Send message to queue
		sendMessages(queueInfo, msmq.SingleMessage)

		queue, err = queueInfo.Open(msmq.Receive, msmq.DenyNone)
		if err != nil {
			log.Fatal(err)
		}

		opts := []msmq.ReceiveOption{
			msmq.ReceiveWithTransaction(msmq.SingleMessage),
			msmq.ReceiveWithWantDestinationQueue(true),
			msmq.ReceiveWithWantBody(true),
			msmq.ReceiveWithTimeout(1),
			msmq.ReceiveWithWantConnectorType(true),
		}
		msg, err := queue.Receive(opts...)
		if err != nil {
			log.Fatal(err)
		}
		s, err := msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Receive (transactional): %s", s)

		msg, err = queue.ReceiveCurrent(opts...)
		if err != nil {
			log.Fatal(err)
		}
		s, err = msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Receive current (transactional): %s", s)

		err = queue.Purge()
		if err != nil {
			log.Fatal(err)
		}

		err = queue.Close()
		if err != nil {
			log.Fatal(err)
		}
	}

	// Receive by look up id transactional
	{
		// Send message to queue
		sendMessages(queueInfo, msmq.SingleMessage)

		queue, err = queueInfo.Open(msmq.Peek, msmq.DenyNone)
		if err != nil {
			log.Fatal(err)
		}
		msg, err := queue.Peek()
		if err != nil {
			log.Fatal(err)
		}
		s, err := msg.LookupID()
		if err != nil {
			log.Fatal(err)
		}
		id, err := strconv.ParseUint(s, 10, 64)
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Msg lookup id: %d", id)

		err = queue.Close()
		if err != nil {
			log.Fatal(err)
		}

		queue, err = queueInfo.Open(msmq.Receive, msmq.DenyNone)
		if err != nil {
			log.Fatal(err)
		}

		opts := []msmq.ReceiveByLookupIDOption{
			msmq.ReceiveByLookupIDWithTransaction(msmq.SingleMessage),
			msmq.ReceiveByLookupIDWithWantDestinationQueue(true),
			msmq.ReceiveByLookupIDWithWantBody(true),
			msmq.ReceiveByLookupIDWithWantConnectorType(true),
		}
		msg, err = queue.ReceiveByLookupID(id, opts...)
		if err != nil {
			log.Fatal(err)
		}

		s, err = msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Receive by lookup id (transactional): %s", s)

		msg, err = queue.ReceiveFirstByLookupID(opts...)
		if err != nil {
			log.Fatal(err)
		}

		s, err = msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Receive by first by lookup id (transactional): %s", s)

		msg, err = queue.ReceiveLastByLookupID(opts...)
		if err != nil {
			log.Fatal(err)
		}

		s, err = msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Receive by last by lookup id (transactional): %s", s)

		msg, err = queue.PeekFirstByLookupID()
		if err != nil {
			log.Fatal(err)
		}

		s, err = msg.LookupID()
		if err != nil {
			log.Fatal(err)
		}
		id, err = strconv.ParseUint(s, 10, 64)
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Msg lookup id: %d", id)

		msg, err = queue.ReceiveNextByLookupID(id, opts...)
		if err != nil {
			log.Fatal(err)
		}

		s, err = msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Receive by next by lookup id (transactional): %s", s)

		msg, err = queue.PeekLastByLookupID()
		if err != nil {
			log.Fatal(err)
		}

		s, err = msg.LookupID()
		if err != nil {
			log.Fatal(err)
		}
		id, err = strconv.ParseUint(s, 10, 64)
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Msg lookup id: %d", id)

		msg, err = queue.ReceivePreviousByLookupID(id, opts...)
		if err != nil {
			log.Fatal(err)
		}

		s, err = msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Receive by previous by lookup id (transactional): %s", s)

		err = queue.Reset()
		if err != nil {
			log.Fatal(err)
		}

		msg, err = queue.Receive(
			msmq.ReceiveWithTransaction(msmq.SingleMessage),
			msmq.ReceiveWithTimeout(1),
		)
		if err != nil {
			log.Fatal(err)
		}

		s, err = msg.Body()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Receive after reset (transactional): %s", s)

		err = queue.Close()
		if err != nil {
			log.Fatal(err)
		}
	}

	am, err := queue.Access()
	if err != nil {
		log.Fatal(err)
	}
	log.Printf("Access Mode: %v", am)

	i, err := queue.Handle()
	if err != nil {
		log.Fatal(err)
	}
	log.Printf("Handle: %d", i)

	b, err = queue.IsOpen()
	if err != nil {
		log.Fatal(err)
	}
	log.Printf("IsOpen: %v", b)

	queueInfo, err = queue.QueueInfo()
	if err != nil {
		log.Fatal(err)
	}
	s, err := queueInfo.FormatName()
	if err != nil {
		log.Fatal(err)
	}
	log.Printf("QueueInfo.FormatName: %s", s)

	sm, err := queue.ShareMode()
	if err != nil {
		log.Fatal(err)
	}
	log.Printf("Share Mode: %v", sm)

	// Delete queue
	err = queueInfo.Delete()
	if err != nil {
		log.Fatal(err)
	}
}

func sendMessages(queueInfo *msmq.QueueInfo, level msmq.TransactionLevel) {
	queue, err := queueInfo.Open(msmq.Send, msmq.DenyNone)
	if err != nil {
		log.Fatal(err)
	}

	msgs := []string{
		"Hello",
		"world",
		"Lorem",
		"ipsum",
		"dolor",
		"sit",
	}

	msg, err := msmq.NewMessage()
	if err != nil {
		log.Fatal(err)
	}
	for _, m := range msgs {
		err = msg.SetBody(m)
		if err != nil {
			log.Fatal(err)
		}
		err = msg.Send(queue, msmq.SendWithTransaction(level))
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Send msg: %s", m)
	}

	err = queue.Close()
	if err != nil {
		log.Fatal(err)
	}
}
