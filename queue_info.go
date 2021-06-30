// +build windows

package msmq

import (
	"errors"
	"fmt"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// QueueInfo provides queue management for a single queue. It provides methods
// for creating a queue (either a transactional or non-transactional queue),
// opening a queue, changing or retrieving properties of a queue, and deleting
// a queue.
type QueueInfo struct {
	dispatch *ole.IDispatch
}

// NewQueueInfo returns a pointer to a QueueInfo. The FormatName or PathName
// must be set before interacting with a queue.
// This can be done through options:
//   queueInfo, err := msmq.NewQueueInfo(msmq.WithFormatName(name))
// Alternatively, it can be done through the QueueInfo.SetFormatName() function:
//   err := queueInfo.SetFormatName(name)
func NewQueueInfo(opts ...QueueInfoOption) (*QueueInfo, error) {
	unknown, err := oleutil.CreateObject("MSMQ.MSMQQueueInfo")
	if err != nil && err.Error() == "Invalid class string" {
		return nil, ErrMSMQNotInstalled
	}

	dispatch, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return nil, err
	}

	queueInfo := &QueueInfo{
		dispatch: dispatch,
	}

	for _, o := range opts {
		err = o.set(queueInfo)
		if err != nil {
			return nil, fmt.Errorf("msmq: failed to create new QueueInfo: %w", err)
		}

	}

	return queueInfo, nil
}

// QueueInfoOption represents an option to configure QueueInfo.
type QueueInfoOption struct {
	set func(qi *QueueInfo) error
}

// WithAuthenticate returns a QueueInfoOption that configures QueueInfo with the
// specified Authenticate value.
func WithAuthenticate(authenticate bool) QueueInfoOption {
	return QueueInfoOption{
		set: func(qi *QueueInfo) error {
			return qi.SetAuthenticate(authenticate)
		},
	}
}

// WithBasePriority returns a QueueInfoOption that configures QueueInfo with the
// specified BasePriority value.
func WithBasePriority(priority int32) QueueInfoOption {
	return QueueInfoOption{
		set: func(qi *QueueInfo) error {
			return qi.SetBasePriority(priority)
		},
	}
}

// WithFormatName returns a QueueInfoOption that configures QueueInfo with the
// specified FormatName value.
func WithFormatName(name string) QueueInfoOption {
	return QueueInfoOption{
		set: func(qi *QueueInfo) error {
			return qi.SetFormatName(name)
		},
	}
}

// WithJournal returns a QueueInfoOption that configures QueueInfo with the
// specified Journal value.
func WithJournal(enabled bool) QueueInfoOption {
	return QueueInfoOption{
		set: func(qi *QueueInfo) error {
			return qi.SetJournal(enabled)
		},
	}
}

// WithJournalQuota returns a QueueInfoOption that configures QueueInfo with
// the specified JournalQuota value.
func WithJournalQuota(size int32) QueueInfoOption {
	return QueueInfoOption{
		set: func(qi *QueueInfo) error {
			return qi.SetJournalQuota(size)
		},
	}
}

// WithLabel returns a QueueInfoOption that configures QueueInfo with the
// specified Label value.
func WithLabel(label string) QueueInfoOption {
	return QueueInfoOption{
		set: func(qi *QueueInfo) error {
			return qi.SetLabel(label)
		},
	}
}

// WithMulticastAddress returns a QueueInfoOption that configures QueueInfo with the
// specified MulticastAddress value.
func WithMulticastAddress(address string) QueueInfoOption {
	return QueueInfoOption{
		set: func(qi *QueueInfo) error {
			return qi.SetMulticastAddress(address)
		},
	}
}

// WithPathName returns a QueueInfoOption that configures QueueInfo with the
// specified PathName value.
func WithPathName(name string) QueueInfoOption {
	return QueueInfoOption{
		set: func(qi *QueueInfo) error {
			return qi.SetPathName(name)
		},
	}
}

// WithPrivacyLevel returns a QueueInfoOption that configures QueueInfo with the
// specified PrivacyLevel value.
func WithPrivacyLevel(level PrivLevel) QueueInfoOption {
	return QueueInfoOption{
		set: func(qi *QueueInfo) error {
			return qi.SetPrivacyLevel(level)
		},
	}
}

// WithQuota returns a QueueInfoOption that configures QueueInfo with
// the specified Quota value.
func WithQuota(size int32) QueueInfoOption {
	return QueueInfoOption{
		set: func(qi *QueueInfo) error {
			return qi.SetQuota(size)
		},
	}
}

// WithServiceTypeGUID returns a QueueInfoOption that configures QueueInfo with
// the specified ServiceTypeGUID value.
func WithServiceTypeGUID(guid string) QueueInfoOption {
	return QueueInfoOption{
		set: func(qi *QueueInfo) error {
			return qi.SetServiceTypeGUID(guid)
		},
	}
}

// ErrMSMQNotInstalled is returned when trying to interact with MSMQ but it is
// not installed.
var ErrMSMQNotInstalled = errors.New("go-msmq: message queuing has not been installed on this computer")

// Create creates a public or private queue based on the options set on QueueInfo.
//
// The PathName option must be set on QueueInfo before calling Create.
//   queueInfo, err := msmq.NewQueueInfo()
//   if err != nil {
//	     log.Error(err)
//   }
//   err = queueInfo.SetPathName(name)
//   if err != nil {
//	     log.Error(err)
//   }
//   err = queueInfo.Create()
//   if err != nil {
//	     log.Error(err)
//   }
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms703983(v=vs.85)
func (qi *QueueInfo) Create(opts ...CreateQueueOption) error {
	s, err := qi.PathName()
	if err != nil {
		return fmt.Errorf("go-msmq: failed to create queue: %w", err)
	}

	if s == "" {
		return fmt.Errorf("go-msmq: failed to create queue: %w", errors.New("Exception occurred. (The queue path name is not set. )"))
	}

	options := &createQueueOptions{
		transactional: false,
		worldReadable: false,
	}
	for _, o := range opts {
		o.set(options)
	}

	_, err = qi.dispatch.CallMethod("Create", options.transactional, options.worldReadable)
	if err != nil {
		return fmt.Errorf("go-msmq: Create(%v, %v) failed to create queue: %w", options.transactional, options.worldReadable, err)
	}
	return nil
}

// CreateQueueOption represents an option to configure the creation of a queue.
type CreateQueueOption struct {
	set func(opts *createQueueOptions)
}

// createQueueOptions contains all the options for creating a queue.
type createQueueOptions struct {
	transactional bool
	worldReadable bool
}

// CreateQueueWithTransactional returns a CreateQueueOption that configures
// the queue with the specified transactional value.
func CreateQueueWithTransactional(transactional bool) CreateQueueOption {
	return CreateQueueOption{
		set: func(opts *createQueueOptions) {
			opts.transactional = transactional
		},
	}
}

// CreateQueueWithWorldReadable returns a CreateQueueOption that configures
// the queue with the specified worldReadable value.
func CreateQueueWithWorldReadable(worldReadable bool) CreateQueueOption {
	return CreateQueueOption{
		set: func(opts *createQueueOptions) {
			opts.worldReadable = worldReadable
		},
	}
}

// Delete deletes the queue that is managed by QueueInfo.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms706050(v=vs.85)
func (qi *QueueInfo) Delete() error {
	_, err := qi.dispatch.CallMethod("Delete")
	if err != nil {
		return fmt.Errorf("go-msmq: Delete() failed to delete queue: %w", err)
	}

	return nil
}

// Open opens a queue for sending, peeking at, retrieving, or purging messages
// and creates a cursor for navigating the queue if the queue is being opened
// for retrieving messages.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms707027(v=vs.85)
func (qi *QueueInfo) Open(accessMode AccessMode, shareMode ShareMode) (*Queue, error) {
	queue, err := qi.dispatch.CallMethod("Open", int(accessMode), int(shareMode))
	if err != nil {
		return nil, fmt.Errorf("go-msmq: Open(%v, %v) failed to open queue: %w", accessMode, shareMode, err)
	}

	return &Queue{
		dispatch: queue.ToIDispatch(),
	}, nil
}

// AccessMode defines access modes for accessing messages within a queue. The
// access mode cannot be changed while a queue is open.
type AccessMode int

const (
	// Receive grants permissions to read, peek, and delete messages from a local queue.
	Receive AccessMode = 1

	// Send grants permissions to insert new messages into a queue.
	Send AccessMode = 2

	// Peek grants permissions to peek but not delete messages from a local queue.
	Peek AccessMode = 32

	// admin specifies that a remote queue is to be opened.
	admin AccessMode = 128

	// PeekAndAdmin grants Peek permissions to a remote queue.
	PeekAndAdmin AccessMode = Peek | admin

	// ReceiveAndAdmin grants Receive permissions to a remote queue.
	ReceiveAndAdmin AccessMode = Receive | admin
)

// ShareMode defines the exclusivity level when accessing a queue. Default
// value is DenyNone.
type ShareMode int

const (
	// DenyNone indicates that accessing a queue is available to all members
	// of the EVERYONE group.
	DenyNone ShareMode = 0

	// DenyReceive limits access to other processes.
	DenyReceive ShareMode = 1
)

// Refresh updates the properties of QueueInfo. For example, if user 1 locates
// the queue and then user 2 modifies the queue's properties, user 1 needs to
// call QueueInfo.Refresh() to sync up with user 2's changes.
//
// All queue properties can be updated. However, you can retrieve the properties
// of private queues only if they are located on your local computer.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms703265(v=vs.85)
func (qi *QueueInfo) Refresh() error {
	_, err := qi.dispatch.CallMethod("Refresh")
	if err != nil {
		return fmt.Errorf("go-msmq: Refresh() failed to retrieve updated properties: %w", err)
	}

	return nil
}

// Update updates the properties of the queue represented by QueueInfo with
// its current property values. It can only be called after a queue has been
// created or before the queue is deleted.
//
// Update can only update the properties of a public queue or a local private
// queue. Additionally, there are some properties that Update cannot update.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms705153(v=vs.85)
func (qi *QueueInfo) Update() error {
	_, err := qi.dispatch.CallMethod("Update")
	if err != nil {
		return fmt.Errorf("go-msmq: Update() failed to update queue: %w", err)
	}
	return nil
}

// ADsPath returns the Active Directory Domain Services (AD DS) path to the
// public queue.
func (qi *QueueInfo) ADsPath() (string, error) {
	res, err := qi.dispatch.GetProperty("ADsPath")
	if err != nil {
		return "", fmt.Errorf("go-msmq: failed to get AD path: %w", err)
	}

	return res.Value().(string), nil
}

// Authenticate returns authenticate.
func (qi *QueueInfo) Authenticate() (bool, error) {
	res, err := qi.dispatch.GetProperty("Authenticate")
	if err != nil {
		return false, fmt.Errorf("go-msmq: failed to get Authenticate: %w", err)
	}

	i := res.Value().(int32)
	return i != 0, nil
}

// SetAuthenticate sets authenticate. Authenticate specifies whether the queue
// only accepts authenticated messages.
//
// The default value is false. The queue accepts authenticated and non-authenticated
// messages.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms703976(v=vs.85)
func (qi *QueueInfo) SetAuthenticate(authenticate bool) error {
	i := 0
	if authenticate {
		i = 1
	}

	_, err := qi.dispatch.PutProperty("Authenticate", i)
	if err != nil {
		return fmt.Errorf("go-msmq: SetAuthenticate(%v) failed to set Authenticate: %w", i, err)
	}

	return nil
}

// BasePriority returns the base priority.
func (qi *QueueInfo) BasePriority() (int32, error) {
	res, err := qi.dispatch.GetProperty("BasePriority")
	if err != nil {
		return 0, fmt.Errorf("go-msmq: failed to get BasePriority: %w", err)
	}

	return res.Value().(int32), nil
}

// SetBasePriority sets base prioirty. Base priority specifies the base priority
// for all messages sent to a public queue.
//
// The default value is 0.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms701847(v=vs.85)
func (qi *QueueInfo) SetBasePriority(priority int32) error {
	_, err := qi.dispatch.PutProperty("BasePriority", priority)
	if err != nil {
		return fmt.Errorf("go-msmq: SetBasePriority(%d) failed to set BasePriority: %w", priority, err)
	}

	return nil
}

// CreateTime returns when the public queue or private queue was created. The
// the value is automatically converted to the local system time and system date.
func (qi *QueueInfo) CreateTime() (time.Time, error) {
	res, err := qi.dispatch.GetProperty("CreateTime")
	if err != nil {
		return time.Time{}, fmt.Errorf("go-msmq: failed to get CreateTime: %w", err)
	}

	return res.Value().(time.Time), nil
}

// FormatName returns the format name.
func (qi *QueueInfo) FormatName() (string, error) {
	res, err := qi.dispatch.GetProperty("FormatName")
	if err != nil {
		return "", fmt.Errorf("go-msmq: failed to get FormatName: %w", err)
	}

	return res.Value().(string), nil
}

// SetFormatName sets the format name. Format names are used to reference public
// or private queues without accessing directory service.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms705703(v=vs.85)
func (qi *QueueInfo) SetFormatName(name string) error {
	_, err := qi.dispatch.PutProperty("FormatName", name)
	if err != nil {
		return fmt.Errorf("go-msmq: SetFormatName(%s) failed to set FormatName: %w", name, err)
	}

	return nil
}

// IsTransactional indicates whether the queue supports transactions.
func (qi *QueueInfo) IsTransactional() (bool, error) {
	res, err := qi.dispatch.GetProperty("IsTransactional2")
	if err != nil {
		return false, fmt.Errorf("go-msmq: failed to get IsTransactional2: %w", err)
	}

	return res.Value().(bool), nil
}

// IsWorldReadable indicates whether all members of the Everyone group can
// read the messages in the queue.
func (qi *QueueInfo) IsWorldReadable() (bool, error) {
	res, err := qi.dispatch.GetProperty("IsWorldReadable2")
	if err != nil {
		return false, fmt.Errorf("go-msmq: failed to get IsWorldReadable: %w", err)
	}

	return res.Value().(bool), nil
}

// Journal returns whether messages retrieved from the queue are stored in the
// journal of the queue.
func (qi *QueueInfo) Journal() (bool, error) {
	res, err := qi.dispatch.GetProperty("Journal")
	if err != nil {
		return false, fmt.Errorf("go-msmq: failed to get Journal: %w", err)
	}

	i := res.Value().(int32)
	return i != 0, nil
}

// SetJournal specifies whether the messages retrieved from the queue are stored
// in the journal of the queue.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms701492(v=vs.85)
func (qi *QueueInfo) SetJournal(enabled bool) error {
	i := 0
	if enabled {
		i = 1
	}

	_, err := qi.dispatch.PutProperty("Journal", i)
	if err != nil {
		return fmt.Errorf("go-msmq: SetJournal(%v) failed to set Journal: %w", enabled, err)
	}

	return nil
}

// JournalQuota returns the maximum size (in kilobytes) of the queue journal.
func (qi *QueueInfo) JournalQuota() (int32, error) {
	res, err := qi.dispatch.GetProperty("JournalQuota")
	if err != nil {
		return 0, fmt.Errorf("go-msmq: failed to get JournalQuota: %w", err)
	}

	return res.Value().(int32), nil
}

// SetJournalQuota specifies the maximum size (in kilobytes) of the queue journal.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms700230(v=vs.85)
func (qi *QueueInfo) SetJournalQuota(size int32) error {
	_, err := qi.dispatch.PutProperty("JournalQuota", size)
	if err != nil {
		return fmt.Errorf("go-msmq: SetJournalQuota(%d) failed to set JournalQuota: %w", size, err)
	}

	return nil
}

// Label returns the description of the queue.
func (qi *QueueInfo) Label() (string, error) {
	res, err := qi.dispatch.GetProperty("Label")
	if err != nil {
		return "", fmt.Errorf("go-msmq: failed to get Label: %w", err)
	}

	return res.Value().(string), nil
}

// SetLabel sets the description of the queue.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms701520(v=vs.85)
func (qi *QueueInfo) SetLabel(label string) error {
	_, err := qi.dispatch.PutProperty("Label", label)
	if err != nil {
		return fmt.Errorf("go-msmq: SetLabel(%s) failed to set Label: %w", label, err)
	}

	return nil
}

// ModifyTime returns when the public queue or private queue was last updated. The
// the value is automatically converted to the local system time and system date.
func (qi *QueueInfo) ModifyTime() (time.Time, error) {
	res, err := qi.dispatch.GetProperty("ModifyTime")
	if err != nil {
		return time.Time{}, fmt.Errorf("go-msmq: failed to get ModifyTime: %w", err)
	}

	return res.Value().(time.Time), nil
}

// MulticastAddress returns the multicast address associated with the queue.
func (qi *QueueInfo) MulticastAddress() (string, error) {
	res, err := qi.dispatch.GetProperty("MulticastAddress")
	if err != nil {
		return "", fmt.Errorf("go-msmq: failed to get MulticastAddress: %w", err)
	}

	return res.Value().(string), nil
}

// SetMulticastAddress sets the multicast address of the queue. The value of
// address should be in the form:
//   <address>:<port>
// An empty string can also be specified to indicate that the queue is not
// associated with a multicast address.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms704978(v=vs.85)
func (qi *QueueInfo) SetMulticastAddress(address string) error {
	_, err := qi.dispatch.PutProperty("MulticastAddress", address)
	if err != nil {
		return fmt.Errorf("go-msmq: SetMulticastAddress(%s) failed to set MulticastAddress: %w", address, err)
	}

	return nil
}

// PathName returns the path name.
func (qi *QueueInfo) PathName() (string, error) {
	res, err := qi.dispatch.GetProperty("PathName")
	if err != nil {
		return "", fmt.Errorf("go-msmq: failed to get PathName: %w", err)
	}

	return res.Value().(string), nil
}

// SetPathName sets the path name which specifies the name of the computer where
// the messages in the queue will be stored, an optional PRIVATE$ keyword that
// indicates whether the queue is a private queue, and the name of the queue.
//
// Path name syntax can be any of:
//   ComputerName\QueueName
//   ComputerName\PRIVATE$\QueueName
//   .\QueueName
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms706083(v=vs.85)
func (qi *QueueInfo) SetPathName(name string) error {
	_, err := qi.dispatch.PutProperty("PathName", name)
	if err != nil {
		return fmt.Errorf("go-msmq: SetPathName(%s) failed to set PathName: %w", name, err)
	}

	return nil
}

// PathNameDNS returns the DNS path name of the queue.
func (qi *QueueInfo) PathNameDNS() (string, error) {
	res, err := qi.dispatch.GetProperty("PathNameDNS")
	if err != nil {
		return "", fmt.Errorf("go-msmq: failed to get PathNameDNS: %w", err)
	}

	return res.Value().(string), nil
}

// PrivLevel returns the privacy level.
func (qi *QueueInfo) PrivacyLevel() (PrivLevel, error) {
	res, err := qi.dispatch.GetProperty("PrivLevel")
	if err != nil {
		return 0, fmt.Errorf("go-msmq: failed to get PrivLevel: %w", err)
	}

	return PrivLevel(res.Value().(int32)), nil
}

// SetPrivacyLevel sets the privacy level of the queue. The default value is
// PrivLevel.Optional.
//
// If the privacy level of a message does not correspond to the privacy level
// of the queue, the message is rejected by the queue.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms701989(v=vs.85)
func (qi *QueueInfo) SetPrivacyLevel(level PrivLevel) error {
	_, err := qi.dispatch.PutProperty("PrivLevel", int(level))
	if err != nil {
		return fmt.Errorf("go-msmq: SetPrivacyLevel(%v) failed to set PrivLevel: %w", level, err)
	}

	return nil
}

// PrivLevel defines the privacy level of the queue. Default value is Optional.
type PrivLevel int

const (
	// NonPrivate specifies that the queue accepts only non-private (clear) messages.
	NonPrivate PrivLevel = 0

	// OptionalPrivate specifies that the queue does not force privacy. It accepts
	// private (encrypted) messages and non-private (clear) messages.
	OptionalPrivate PrivLevel = 1

	// OnlyPrivate specifies that the queue accepts only private (encrypted) messages.
	OnlyPrivate PrivLevel = 2
)

// QueueGUID returns GUID of the public queue in the form:
//   {12345678-1234-1234-1234-123456789ABC}
func (qi *QueueInfo) QueueGUID() (string, error) {
	res, err := qi.dispatch.GetProperty("QueueGuid")
	if err != nil {
		return "", fmt.Errorf("go-msmq: failed to get QueueGuid : %w", err)
	}

	return res.Value().(string), nil
}

// Quota returns the maximum size (in kilobytes) of the queue.
func (qi *QueueInfo) Quota() (int32, error) {
	res, err := qi.dispatch.GetProperty("Quota")
	if err != nil {
		return 0, fmt.Errorf("go-msmq: failed to get Quota: %w", err)
	}

	return res.Value().(int32), nil
}

// SetQuota specifies the maximum size (in kilobytes) of the queue. The default
// is INFINITE - this is limited only by the available disk space on the local
// computer or the computer quota.
//
// When the quota of the queue is changed, the new quota affects only arriving
// messages; it does not affect messages already in the queue.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms707016(v=vs.85)
func (qi *QueueInfo) SetQuota(size int32) error {
	_, err := qi.dispatch.PutProperty("Quota", size)
	if err != nil {
		return fmt.Errorf("go-msmq: SetQuota(%d) failed to set Quota: %w", size, err)
	}

	return nil
}

// ServiceTypeGUID returns the GUID that specifies the type of service provided
// by the queue in the form:
//   {12345678-1234-1234-1234-123456789ABC}
func (qi *QueueInfo) ServiceTypeGUID() (string, error) {
	res, err := qi.dispatch.GetProperty("ServiceTypeGuid")
	if err != nil {
		return "", fmt.Errorf("go-msmq: failed to get ServiceTypeGUID: %w", err)
	}

	return res.Value().(string), nil
}

// SetServiceTypeGUID specifies the type of service provided by the queue. It is
// used to identify the queue by its type of service. This identifier can be used
// to locate public queues registered in the directory service.
//
// See: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms703206(v=vs.85)
func (qi *QueueInfo) SetServiceTypeGUID(guid string) error {
	_, err := qi.dispatch.PutProperty("ServiceTypeGuid", guid)
	if err != nil {
		return fmt.Errorf("go-msmq: SetServiceTypeGUID(%s) failed to set ServiceTypeGuid: %w", guid, err)
	}

	return nil
}
