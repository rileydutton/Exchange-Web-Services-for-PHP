<?php

/**
 * Exchangeclient class.
 *
 * @author Riley Dutton
 * @author Rudolf Leermakers
 */
class ExchangeClient {

    private $wsdl;
    private $client;
    private $user;
    private $pass;

    /**
     * The last error that occurred when communicating with the Exchange server.
     *
     * @var mixed
     * @access public
     */
    public $lastError;
    private $impersonate;
    private $delegate;

    /**
     * Initialize the class. This could be better as a __construct, but for CodeIgniter compatibility we keep it separate.
     *
     * @access public
     * @param string $user (the username of the mailbox account you want to use on the Exchange server)
     * @param string $pass (the password of the account)
     * @param string $delegate. (the email address you would like to access...the account you are logging in as must be an administrator account.
     * @param string $wsdl. (The path to the WSDL file. If you put them in the same directory as the Exchangeclient.php script, you can leave this alone. default: "Services.wsdl")
     * @return void
     */
    public function init($user, $pass, $delegate = NULL, $wsdl = "Services.wsdl") {
        $this->wsdl = $wsdl;
        $this->user = $user;
        $this->pass = $pass;
        $this->delegate = $delegate;

        $this->setup();

        $this->client = new NTLMSoapClient($this->wsdl, array(
            'trace' => 1,
            'exceptions' => true,
            'login' => $user,
            'password' => $pass
        ));

        $this->teardown();
    }

    /**
     * Create an event in the user's calendar. Does not currently support sending invitations to other users. Times must be passed as ISO date format.
     *
     * @access public
     * @param string $subject
     * @param string $start (start time of event in ISO date format e.g. "2010-09-21T16:00:00Z"
     * @param string $end (ISO date format)
     * @param string $location
     * @param bool $isallday. (default: false)
     * @return bool $success (true if the message was created, false if there was an error)
     */
    public function create_event($subject, $start, $end, $location, $isallday = false) {
        $this->setup();

        $CreateItem->SendMeetingInvitations = "SendToNone";
        $CreateItem->SavedItemFolderId->DistinguishedFolderId->Id = "calendar";
        if($this->delegate != NULL) {
            $FindItem->SavedItemFolderId->DistinguishedFolderId->Mailbox->EmailAddress = $this->delegate;
        }
        $CreateItem->Items->CalendarItem->Subject = $subject;
        $CreateItem->Items->CalendarItem->Start = $start; #e.g. "2010-09-21T16:00:00Z"; # ISO date format. Z denotes UTC time
        $CreateItem->Items->CalendarItem->End = $end;
        $CreateItem->Items->CalendarItem->IsAllDayEvent = $isallday;
        $CreateItem->Items->CalendarItem->LegacyFreeBusyStatus = "Busy";
        $CreateItem->Items->CalendarItem->Location = $location;

        $response = $this->client->CreateItem($CreateItem);

        $this->teardown();

        if($response->ResponseMessages->CreateItemResponseMessage->ResponseCode == "NoError")
            return true;
        else {
            $this->lastError = $response->ResponseMessages->CreateItemResponseMessage->ResponseCode;
            return false;
        }
    }

    public function get_events($start, $end) {
        $this->setup();

        $FindItem->Traversal = "Shallow";
        $FindItem->ItemShape->BaseShape = "IdOnly";
        $FindItem->ParentFolderIds->DistinguishedFolderId->Id = "calendar";

        if($this->delegate != NULL) {
            $FindItem->ParentFolderIds->DistinguishedFolderId->Mailbox->EmailAddress = $this->delegate;
        }

        $FindItem->CalendarView->StartDate = $start;
        $FindItem->CalendarView->EndDate = $end;

        $response = $this->client->FindItem($FindItem);

        if($response->ResponseMessages->FindItemResponseMessage->ResponseCode != "NoError") {
            $this->lastError = $response->ResponseMessages->FindItemResponseMessage->ResponseCode;
            return false;
        }

        $items = $response->ResponseMessages->FindItemResponseMessage->RootFolder->Items->CalendarItem;

        $i = 0;
        $events = array();

        if(count($items) == 0)
            return false; //we didn't get anything back!

        if(!is_array($items)) //if we only returned one event, then it doesn't send it as an array, just as a single object. so put it into an array so that everything works as expected.
            $items = array($items);

        foreach($items as $item) {
            $GetItem->ItemShape->BaseShape = "Default";
            $GetItem->ItemIds->ItemId = $item->ItemId;
            $response = $this->client->GetItem($GetItem);

            if($response->ResponseMessages->GetItemResponseMessage->ResponseCode != "NoError") {
                $this->lastError = $response->ResponseMessages->GetItemResponseMessage->ResponseCode;
                return false;
            }

            $eventobj = $response->ResponseMessages->GetItemResponseMessage->Items->CalendarItem;

            $newevent = null;
            $newevent->id = $eventobj->ItemId->Id;
            $newevent->changekey = $eventobj->ItemId->ChangeKey;
            $newevent->subject = $eventobj->Subject;
            $newevent->start = strtotime($eventobj->Start);
            $newevent->end = strtotime($eventobj->End);
            $newevent->location = $eventobj->Location;

            $organizer = null;
            $organizer->name = $eventobj->Organizer->Mailbox->Name;
            $organizer->email = $eventobj->Organizer->Mailbox->EmailAddress;

            $people = array();
            $required = $eventobj->RequiredAttendees->Attendee;

            if(!is_array($required))
                $required = array($required);

            foreach($required as $r) {
                $o = null;
                $o->name = $r->Mailbox->Name;
                $o->email = $r->Mailbox->EmailAddress;
                $people[] = $o;
            }

            $newevent->organizer = $organizer;
            $newevent->people = $people;
            $newevent->allpeople = array_merge(array($organizer), $people);

            $events[] = $newevent;
        }

        $this->teardown();

        return $events;
    }

    /**
     * Get the messages for a mailbox.
     *
     * @access public
     * @param int $limit. (How many messages to get? default: 50)
     * @param bool $onlyunread. (Only get unread messages? default: false)
     * @param string $folder. (default: "inbox", other options include "sentitems")
     * @param bool $folderIdIsDistinguishedFolderId. (default: true, is $folder a DistinguishedFolderId or a simple FolderId)
     * @return array $messages (an array of objects representing the messages)
     */
    public function get_messages($limit = 50, $onlyunread = false, $folder = "inbox", $folderIdIsDistinguishedFolderId = true) {
        $this->setup();

        $FindItem = new stdClass();
        $FindItem->Traversal = "Shallow";

        $FindItem->ItemShape = new stdClass();
        $FindItem->ItemShape->BaseShape = "IdOnly";

        $FindItem->ParentFolderIds = new stdClass();

        if($folderIdIsDistinguishedFolderId) {
            $FindItem->ParentFolderIds->DistinguishedFolderId = new stdClass();
            $FindItem->ParentFolderIds->DistinguishedFolderId->Id = $folder;
        } else {
            $FindItem->ParentFolderIds->FolderId = new stdClass();
            $FindItem->ParentFolderIds->FolderId->Id = $folder;
        }

        if($this->delegate != NULL) {
            $FindItem->ParentFolderIds->DistinguishedFolderId->Mailbox->EmailAddress = $this->delegate;
        }

        $response = $this->client->FindItem($FindItem);

        if($response->ResponseMessages->FindItemResponseMessage->ResponseCode != "NoError") {
            $this->lastError = $response->ResponseMessages->FindItemResponseMessage->ResponseCode;
            return false;
        }

        $items = $response->ResponseMessages->FindItemResponseMessage->RootFolder->Items->Message;

        $i = 0;
        $messages = array();

        if(count($items) == 0)
            return false; //we didn't get anything back!

        if(!is_array($items)) //if we only returned one message, then it doesn't send it as an array, just as a single object. so put it into an array so that everything works as expected.
            $items = array($items);

        foreach($items as $item) {
            $GetItem = new stdClass();
            $GetItem->ItemShape = new stdClass();

            $GetItem->ItemShape->BaseShape = "Default";
            $GetItem->ItemShape->IncludeMimeContent = "true";

            $GetItem->ItemIds = new stdClass();
            $GetItem->ItemIds->ItemId = $item->ItemId;

            $response = $this->client->GetItem($GetItem);

            if($response->ResponseMessages->GetItemResponseMessage->ResponseCode != "NoError") {
                $this->lastError = $response->ResponseMessages->GetItemResponseMessage->ResponseCode;
                return false;
            }

            $messageobj = $response->ResponseMessages->GetItemResponseMessage->Items->Message;

            if($onlyunread && $messageobj->IsRead)
                continue;

            $newmessage = new stdClass();
            $newmessage->source = base64_decode($messageobj->MimeContent->_);
            $newmessage->bodytext = $messageobj->Body->_;
            $newmessage->bodytype = $messageobj->Body->BodyType;
            $newmessage->isread = $messageobj->IsRead;
            $newmessage->ItemId = $item->ItemId;
            $newmessage->from = $messageobj->From->Mailbox->EmailAddress;
            $newmessage->from_name = $messageobj->From->Mailbox->Name;

            $newmessage->to_recipients = array();

            if(!is_array($messageobj->ToRecipients->Mailbox)) {
                $messageobj->ToRecipients->Mailbox = array($messageobj->ToRecipients->Mailbox);
            }

            foreach($messageobj->ToRecipients->Mailbox as $mailbox) {
                $newmessage->to_recipients[] = $mailbox;
            }

            $newmessage->cc_recipients = array();

            if(isset($messageobj->CcRecipients->Mailbox)) {
                if(!is_array($messageobj->CcRecipients->Mailbox)) {
                    $messageobj->CcRecipients->Mailbox = array($messageobj->CcRecipients->Mailbox);
                }

                foreach($messageobj->CcRecipients->Mailbox as $mailbox) {
                    $newmessage->cc_recipients[] = $mailbox;
                }
            }

            $newmessage->time_sent = $messageobj->DateTimeSent;
            $newmessage->time_created = $messageobj->DateTimeCreated;
            $newmessage->subject = $messageobj->Subject;
            $newmessage->attachments = array();

            if($messageobj->HasAttachments == 1) {
                if(property_exists($messageobj->Attachments, 'FileAttachment')) {
                    if(!is_array($messageobj->Attachments->FileAttachment)) {
                        $messageobj->Attachments->FileAttachment = array($messageobj->Attachments->FileAttachment);
                    }

                    foreach($messageobj->Attachments->FileAttachment as $attachment) {
                        $newmessage->attachments[] = $this->get_attachment($attachment->AttachmentId);
                    }
                }
            }

            $messages[] = $newmessage;

            if(++$i > $limit) {
                break;
            }
        }

        $this->teardown();

        return $messages;
    }

    private function get_attachment($AttachmentID) {
        $GetAttachment = new stdClass();
        $GetAttachment->AttachmentIds = new stdClass();

        $GetAttachment->AttachmentIds->AttachmentId = $AttachmentID;

        $response = $this->client->GetAttachment($GetAttachment);

        if($response->ResponseMessages->GetAttachmentResponseMessage->ResponseCode != "NoError") {
            $this->lastError = $response->ResponseMessages->GetAttachmentResponseMessage->ResponseCode;
            return false;
        }

        $attachmentobj = $response->ResponseMessages->GetAttachmentResponseMessage->Attachments->FileAttachment;

        return $attachmentobj;
    }

    /**
     * Send a message through the Exchange server as the currently logged-in user.
     *
     * @access public
     * @param mixed $to (the email address or an array of email address to send the message to)
     * @param string $subject
     * @param string $content
     * @param string $bodytype. (default: "Text", "HTML" for HTML emails)
     * @param bool $saveinsent. (Save in the user's sent folder after sending? default: true)
     * @param bool $markasread. (Mark as read after sending? This currently does nothing. default: true)
     * @param array $attachments. (Array of files to attach. Each array item should be the full path to the file you want to attach)
     * @param mixed $cc (the email address or an array of email address of recipients to receive a carbon copy (cc) of the e-mail message)
     * @param mixed $bcc (the email address or an array of email address of recipients to receive a blind carbon copy (Bcc) of the e-mail message)
     * @return bool $success. (True if the message was sent, false if there was an error).
     */
    public function send_message($to, $subject, $content, $bodytype = "Text", $saveinsent = true, $markasread = true, $attachments = false, $cc = false, $bcc = false) {
        $this->setup();

        $CreateItem = new stdClass();

        if($attachments && !is_array($attachments))
            $attachments = false;

        if($attachments) {
            $CreateItem->MessageDisposition = "SaveOnly";

            $CreateItem->SavedItemFolderId = new stdClass();
            $CreateItem->SavedItemFolderId->DistinguishedFolderId = new stdClass();

            $CreateItem->SavedItemFolderId->DistinguishedFolderId->Id = 'drafts';
        } else {
            if($saveinsent) {
                $CreateItem->SavedItemFolderId = new stdClass();
                $CreateItem->SavedItemFolderId->DistinguishedFolderId = new stdClass();

                $CreateItem->SavedItemFolderId->DistinguishedFolderId->Id = "sentitems";
                $CreateItem->MessageDisposition = "SendAndSaveCopy";
            } else {
                $CreateItem->MessageDisposition = "SendOnly";
            }
        }

        $CreateItem->Items = new stdClass();
        $CreateItem->Items->Message = new stdClass();
        $CreateItem->Items->Message->Body = new stdClass();

        $CreateItem->Items->Message->ItemClass = "IPM.Note";
        $CreateItem->Items->Message->Subject = $subject;
        $CreateItem->Items->Message->Body->BodyType = $bodytype;
        $CreateItem->Items->Message->Body->_ = $content;

        if(is_array($to)) {
            $recipients = array();
            foreach($to as $EmailAddress) {
                $Mailbox = new stdClass();
                $Mailbox->EmailAddress = $EmailAddress;
                $recipients[] = $Mailbox;
            }

            $CreateItem->Items->Message->ToRecipients = new stdClass();

            $CreateItem->Items->Message->ToRecipients->Mailbox = $recipients;
        } else {
            $CreateItem->Items->Message->ToRecipients = new stdClass();
            $CreateItem->Items->Message->ToRecipients->Mailbox = new stdClass();

            $CreateItem->Items->Message->ToRecipients->Mailbox->EmailAddress = $to;
        }

        if($cc) {
            if(is_array($cc)) {
                $recipients = array();
                foreach($cc as $EmailAddress) {
                    $Mailbox = new stdClass();
                    $Mailbox->EmailAddress = $EmailAddress;
                    $recipients[] = $Mailbox;
                }

                $CreateItem->Items->Message->CcRecipients = new stdClass();

                $CreateItem->Items->Message->CcRecipients->Mailbox = $recipients;
            } else {
                $CreateItem->Items->Message->CcRecipients = new stdClass();
                $CreateItem->Items->Message->CcRecipients->Mailbox = new stdClass();

                $CreateItem->Items->Message->CcRecipients->Mailbox->EmailAddress = $cc;
            }
        }

        if($bcc) {
            if(is_array($bcc)) {
                $recipients = array();
                foreach($bcc as $EmailAddress) {
                    $Mailbox = new stdClass();
                    $Mailbox->EmailAddress = $EmailAddress;
                    $recipients[] = $Mailbox;
                }

                $CreateItem->Items->Message->BccRecipients = new stdClass();

                $CreateItem->Items->Message->BccRecipients->Mailbox = $recipients;
            } else {
                $CreateItem->Items->Message->BccRecipients = new stdClass();
                $CreateItem->Items->Message->BccRecipients->Mailbox = new stdClass();

                $CreateItem->Items->Message->BccRecipients->Mailbox->EmailAddress = $bcc;
            }
        }

        if($markasread) {
            $CreateItem->Items->Message->IsRead = "true";
        }

        if($this->delegate != NULL) {
            $CreateItem->Items->Message->From = new stdClass();
            $CreateItem->Items->Message->From->Mailbox = new stdClass();

            $CreateItem->Items->Message->From->Mailbox->EmailAddress = $this->delegate;
        }

        $response = $this->client->CreateItem($CreateItem);

        if($response->ResponseMessages->CreateItemResponseMessage->ResponseCode != "NoError") {
            $this->lastError = $response->ResponseMessages->CreateItemResponseMessage->ResponseCode;
            $this->teardown();
            return false;
        }

        if($attachments && $response->ResponseMessages->CreateItemResponseMessage->ResponseCode == "NoError") {
            $itemId = $response->ResponseMessages->CreateItemResponseMessage->Items->Message->ItemId->Id;
            $itemChangeKey = $response->ResponseMessages->CreateItemResponseMessage->Items->Message->ItemId->ChangeKey;

            foreach($attachments AS $attachment) {
                if(!file_exists($attachment))
                    continue;

                $attachmentMime = "";

                $fileExtension = pathinfo($attachment, PATHINFO_EXTENSION);

                if($fileExtension == "xls")
                    $attachmentMime = "application/vnd.ms-excel";

                if(!strlen($attachmentMime)) {
                    $fileInfo = finfo_open(FILEINFO_MIME_TYPE);
                    $attachmentMime = finfo_file($fileInfo, $attachment);
                    finfo_close($fileInfo);
                }

                $fileInfo = finfo_open(FILEINFO_MIME_TYPE);
                $attachmentMime = finfo_file($fileInfo, $attachment);
                finfo_close($fileInfo);

                $filename = pathinfo($attachment, PATHINFO_BASENAME);

                $FileAttachment = new stdClass();
                $FileAttachment->Content = file_get_contents($attachment);
                $FileAttachment->ContentType = $attachmentMime;
                $FileAttachment->Name = $filename;

                $CreateAttachment = new stdClass();
                $CreateAttachment->Attachments = new stdClass();
                $CreateAttachment->ParentItemId = new stdClass();

                $CreateAttachment->Attachments->FileAttachment = $FileAttachment;
                $CreateAttachment->ParentItemId->Id = $itemId;
                $CreateAttachment->ParentItemId->ChangeKey = $itemChangeKey;

                $response = $this->client->CreateAttachment($CreateAttachment);

                if($response->ResponseMessages->CreateAttachmentResponseMessage->ResponseCode != "NoError") {
                    $this->lastError = $response->ResponseMessages->CreateAttachmentResponseMessage->ResponseCode;
                    return false;
                }

                $itemId = $response->ResponseMessages->CreateAttachmentResponseMessage->Attachments->FileAttachment->AttachmentId->RootItemId;
                $itemChangeKey = $response->ResponseMessages->CreateAttachmentResponseMessage->Attachments->FileAttachment->AttachmentId->RootItemChangeKey;
            }

            $SendItem = new stdClass();
            $SendItem->ItemIds = new stdClass();
            $SendItem->ItemIds->ItemId = new stdClass();

            $SendItem->ItemIds->ItemId->Id = $itemId;
            $SendItem->ItemIds->ItemId->ChangeKey = $itemChangeKey;

            if($saveinsent) {
                $SendItem->SaveItemToFolder = true;

                $SendItem->SavedItemFolderId = new stdClass();
                $SendItem->SavedItemFolderId->DistinguishedFolderId = new stdClass();

                $SendItem->SavedItemFolderId->DistinguishedFolderId->Id = "sentitems";
            } else {
                $SendItem->SaveItemToFolder = false;
            }

            $response = $this->client->SendItem($SendItem);

            if($response->ResponseMessages->SendItemResponseMessage->ResponseCode != "NoError") {
                $this->lastError = $response->ResponseMessages->SendItemResponseMessage->ResponseCode;
                $this->teardown();
                return false;
            }
        }

        $this->teardown();

        return true;
    }

    /**
     * Deletes a message in the mailbox of the current user.
     *
     * @access public
     * @param ItemId $ItemId (such as one returned by get_messages)
     * @param string $deletetype. (default: "HardDelete")
     * @return bool $success (true: message was deleted, false: message failed to delete)
     */
    public function delete_message($ItemId, $deletetype = "HardDelete") {
        $this->setup();

        $DeleteItem->DeleteType = $deletetype;
        $DeleteItem->ItemIds->ItemId = $ItemId;

        $response = $this->client->DeleteItem($DeleteItem);

        $this->teardown();

        if($response->ResponseMessages->DeleteItemResponseMessage->ResponseCode == "NoError") {
            return true;
        } else {
            $this->lastError = $response->ResponseMessages->DeleteItemResponseMessage->ResponseCode;
            return false;
        }
    }

    /**
     * Moves a message to a different folder.
     *
     * @access public
     * @param ItemId $ItemId (such as one returned by get_messages, has Id and ChangeKey)
     * @return ItemID $ItemId The new ItemId (such as one returned by get_messages, has Id and ChangeKey)
     */
    public function move_message($ItemId, $FolderId) {
        $this->setup();

        $MoveItem = new stdClass();
        $MoveItem->ToFolderId = new stdClass();
        $MoveItem->ToFolderId->FolderId = new stdClass();
        $MoveItem->ItemIds = new stdClass();

        $MoveItem->ToFolderId->FolderId->Id = $FolderId;
        $MoveItem->ItemIds->ItemId = $ItemId;

        $response = $this->client->MoveItem($MoveItem);

        if($response->ResponseMessages->MoveItemResponseMessage->ResponseCode == "NoError") {
            return $response->ResponseMessages->MoveItemResponseMessage->Items->Message->ItemId;
        } else {
            $this->lastError = $response->ResponseMessages->MoveItemResponseMessage->ResponseCode;
        }
    }

    /**
     * Mark message(s) as read
     *
     * @access public
     * @param array $messages (the message or an array of messages to mark as read)
     * @return boolean $response. (Returns true if successful and false on error).
     */
    public function mark_message_as_read($messages) {
        $this->setup();

        if(!is_array($messages)) {
            $messages = array($messages);
        }

        $request = new stdClass();
        $request->MessageDisposition = 'SaveOnly';
        $request->ConflictResolution = 'AlwaysOverwrite';
        $request->ItemChanges = array();

        foreach($messages as $message) {
            $change = new stdClass();
            $change->ItemId = new stdClass();
            $change->ItemId->Id = $message->ItemId->Id;
            $change->ItemId->ChangeKey = $message->ItemId->ChangeKey;

            $field = new stdClass();
            $field->FieldURI = new stdClass();
            $field->FieldURI->FieldURI = 'message:IsRead';
            $field->Message = new stdClass();
            $field->Message->IsReadSpecified = true;
            $field->Message->IsRead = true;

            $change->Updates->SetItemField[] = $field;
            $request->ItemChanges[] = $change;
        }

        $response = $this->client->UpdateItem($request);
        $this->teardown();

        if($response->ResponseMessages->UpdateItemResponseMessage->ResponseCode == 'NoError') {
            return true;
        } else {
            $this->lastError = $response->ResponseMessages->UpdateItemResponseMessage->ResponseCode;
            return false;
        }
    }

    public function getFolder($regex, $parent = 'inbox') {
        foreach($this->get_subfolders($parent) as $folder) {
            if(preg_match(sprintf('#%s#', $regex), $folder->DisplayName))
                return $folder;
        }

        return false;
    }

    /**
     * Get all subfolders of a single folder.
     *
     * @access public
     * @param string $ParentFolderId string representing the folder id of the parent folder, defaults to "inbox"
     * @param bool $Distinguished Defines whether or not its a distinguished folder name or not
     * @return object $response the response containing all the folders
     */
    public function get_subfolders($ParentFolderId = "inbox", $Distinguished = TRUE) {
        $this->setup();

        $FolderItem = new stdClass();
        $FolderItem->FolderShape = new stdClass();
        $FolderItem->ParentFolderIds = new stdClass();

        $FolderItem->FolderShape->BaseShape = "Default";
        $FolderItem->Traversal = "Shallow";

        if($Distinguished) {
            $FolderItem->ParentFolderIds->DistinguishedFolderId = new stdClass();
            $FolderItem->ParentFolderIds->DistinguishedFolderId->Id = $ParentFolderId;
        } else {
            $FolderItem->ParentFolderIds->FolderId = new stdClass();
            $FolderItem->ParentFolderIds->FolderId->Id = $ParentFolderId;
        }

        $response = $this->client->FindFolder($FolderItem);

        if($response->ResponseMessages->FindFolderResponseMessage->ResponseCode == "NoError") {
            $folders = array();

            if(!is_array($response->ResponseMessages->FindFolderResponseMessage->RootFolder->Folders->Folder)) {
                $folders[] = $response->ResponseMessages->FindFolderResponseMessage->RootFolder->Folders->Folder;
            } else {
                $folders = $response->ResponseMessages->FindFolderResponseMessage->RootFolder->Folders->Folder;
            }

            return $folders;
        } else {
            $this->lastError = $response->ResponseMessages->FindFolderResponseMessage->ResponseCode;
        }
    }

    /**
     * Sets up strream handling. Internally used.
     *
     * @access private
     * @return void
     */
    private function setup() {
        if($this->impersonate != NULL) {
            $impheader = new ImpersonationHeader($this->impersonate);
            $header = new SoapHeader("http://schemas.microsoft.com/exchange/services/2006/messages", "ExchangeImpersonation", $impheader, false);
            $this->client->__setSoapHeaders($header);
        }

        ExchangeNTLMStream::setCredentials($this->user, $this->pass);

        stream_wrapper_unregister('http');
        stream_wrapper_unregister('https');

        if(!stream_wrapper_register('http', 'ExchangeNTLMStream')) {
            throw new Exception("Failed to register protocol");
        }

        if(!stream_wrapper_register('https', 'ExchangeNTLMStream')) {
            throw new Exception("Failed to register protocol");
        }
    }

    /**
     * Tears down stream handling. Internally used.
     *
     * @access private
     * @return void
     */
    private function teardown() {
        stream_wrapper_restore('http');
        stream_wrapper_restore('https');
    }

}

class ImpersonationHeader {

    public $ConnectingSID;

    function __construct($email) {
        $this->ConnectingSID->PrimarySmtpAddress = $email;
    }

}
