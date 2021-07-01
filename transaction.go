// +build windows

package msmq

// TransactionLevel defines transaction levels for message transactions with a queue.
type TransactionLevel int

const (
	// NoTransaction specifies that the message is not part of a transaction.
	// This level cannot be used to send or receive a message from a
	// transactional queue.
	NoTransaction TransactionLevel = iota

	// MTS specifies that MSMQ will determine if the application
	// is running in the context of a COM+ (Component Services) transaction,
	// then the message is sent or received within the current COM+
	// transaction. Otherwise, the message is sent or received outside of a
	// transaction.
	MTS

	// XA specifies that the message is part of an externally
	// coordinated XA transaction.
	XA

	// SingleMessage specifies that the message is sent or received in a
	// single-message transaction. Messages in a single message transaction
	// must be sent or received from a transactional queue.
	SingleMessage
)
