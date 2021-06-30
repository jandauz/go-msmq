// +build windows

package main

import (
	"log"

	"github.com/jandauz/go-msmq"
)

func main() {
	// Creating QueueInfo
	opts := []msmq.QueueInfoOption{
		msmq.WithAuthenticate(false),
		msmq.WithBasePriority(0),
		msmq.WithPathName(`.\private$\go-msmq`),
		msmq.WithJournal(true),
		msmq.WithJournalQuota(100),
		msmq.WithLabel("go-msmq"),
		msmq.WithMulticastAddress(""),
		msmq.WithPrivacyLevel(msmq.NonPrivate),
		msmq.WithQuota(8_000_000),
		msmq.WithServiceTypeGUID("{12345678-1234-1234-1234-123456789ABC}"),
	}
	queueInfo, err := msmq.NewQueueInfo(opts...)
	if err != nil {
		log.Fatal(err)
	}

	// Create queue
	err = queueInfo.Create(msmq.CreateQueueWithTransactional(true), msmq.CreateQueueWithWorldReadable(true))
	if err != nil {
		log.Fatal(err)
	}

	// Refresh queue to retrieve auto-generated FormatName
	err = queueInfo.Refresh()
	if err != nil {
		log.Fatal(err)
	}

	// Open queue
	_, err = queueInfo.Open(msmq.Peek, msmq.DenyNone)
	if err != nil {
		log.Fatal(err)
	}

	// Print QueueInfo properties
	{
		s, err := queueInfo.ADsPath()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("ADsPath: %s", s)

		b, err := queueInfo.Authenticate()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Authenticate: %v", b)

		i, err := queueInfo.BasePriority()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("BasePriority: %d", i)

		t, err := queueInfo.CreateTime()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("CreateTime: %s", t.String())

		s, err = queueInfo.FormatName()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("FormatName: %s", s)

		b, err = queueInfo.IsTransactional()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("IsTransactional: %v", b)

		b, err = queueInfo.IsWorldReadable()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("IsWorldReadable: %v", b)

		s, err = queueInfo.PathName()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("PathName: %s", s)

		b, err = queueInfo.Journal()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Journal: %v", b)

		i, err = queueInfo.JournalQuota()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("JournalQuota: %dkb", i)

		s, err = queueInfo.Label()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Label: %s", s)

		t, err = queueInfo.ModifyTime()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("ModifyTime: %s", t.String())

		s, err = queueInfo.MulticastAddress()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("MulticastAddress: %s", s)

		s, err = queueInfo.PathName()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("PathName: %s", s)

		s, err = queueInfo.PathNameDNS()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("PathNameDNS: %s", s)

		pl, err := queueInfo.PrivacyLevel()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("PrivacyLevel: %v", pl)

		s, err = queueInfo.QueueGUID()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("QueueGUID: %s", s)

		i, err = queueInfo.Quota()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Quota: %dkb", i)

		s, err = queueInfo.ServiceTypeGUID()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("ServiceTypeGUID: %s", s)
	}

	// Simulate second user updating queue and first user refreshing QueueInfo
	{
		queueInfo2, err := msmq.NewQueueInfo(msmq.WithPathName(`.\private$\go-msmq`))
		if err != nil {
			log.Fatal(err)
		}

		_, err = queueInfo2.Open(msmq.Send, msmq.DenyNone)
		if err != nil {
			log.Fatal(err)
		}

		err = queueInfo2.SetAuthenticate(true)
		if err != nil {
			log.Fatal(err)
		}

		err = queueInfo2.Update()
		if err != nil {
			log.Fatal(err)
		}

		err = queueInfo.Refresh()
		if err != nil {
			log.Fatal(err)
		}

		b, err := queueInfo.Authenticate()
		if err != nil {
			log.Fatal(err)
		}
		log.Printf("Authenticate: %v", b)
	}

	// Delete queue
	err = queueInfo.Delete()
	if err != nil {
		log.Fatal(err)
	}
}
