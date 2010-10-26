<?php
/**
 * Exchangeclient class.
 *
 * @author Riley Dutton
 */
class Exchangeclient {  

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

	/**
	 * Initialize the class. This could be better as a __construct, but for CodeIgniter compatibility we keep it separate.
	 * 
	 * @access public
	 * @param string $user (the username of the mailbox account you want to use on the Exchange server)
	 * @param string $pass (the password of the account)
	 * @param string $impersonate. (the email address of someone you want to impoersonate on the server [NOTE: Your server must have Impersonation enabled (most don't) and you're account must have impersonation permissions (most don't), otherwise this will cause everything to fail! If you don't know what it is, leave it alone :-)] default: NULL)
	 * @param string $wsdl. (The path to the WSDL file. If you put them in the same directory as the Exchangeclient.php script, you can leave this alone. default: "Services.wsdl")
	 * @return void
	 */
	function init($user, $pass, $impersonate=NULL, $wsdl="Services.wsdl") {
		$this->wsdl = $wsdl;
		$this->user = $user;
		$this->pass = $pass;
		
		$this->client = new ExchangeNTLMSoapClient($this->wsdl);
		$this->client->user = $this->user;
		$this->client->password = $this->pass;
	
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
   function create_event($subject, $start, $end, $location, $isallday=false) {
		
		$this->setup();
		
		$CreateItem->SendMeetingInvitations = "SendToNone";
		$CreateItem->SavedItemFolderId->DistinguishedFolderId->Id = "calendar";
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
	
	/**
	 * Get the messages for a mailbox.
	 * 
	 * @access public
	 * @param int $limit. (How many messages to get? default: 50)
	 * @param bool $onlyunread. (Only get unread messages? default: false)
	 * @param string $folder. (default: "inbox", other options include "sentitems", this must be a DistinguishedFolderId)
	 * @return array $messages (an array of objects representing the messages)
	 */
	function get_messages($limit=50, $onlyunread=false, $folder="inbox") {
		
		$this->setup();
		
		$FindItem->Traversal = "Shallow";
		$FindItem->ItemShape->BaseShape = "IdOnly";
		$FindItem->ParentFolderIds->DistinguishedFolderId->Id = "inbox";
		
		$response = $this->client->FindItem($FindItem);
		
		if($response->ResponseMessages->FindItemResponseMessage->ResponseCode != "NoError") {
			$this->lastError = $response->ResponseMessages->FindItemResponseMessage->ResponseCode;
			return false;
		}
		
		$items = $response->ResponseMessages->FindItemResponseMessage->RootFolder->Items->Message;
		
		$i = 0;
		$messages = array();
		
		foreach($items as $item) {
			$GetItem->ItemShape->BaseShape = "Default";
			$GetItem->ItemShape->IncludeMimeContent = "true";
			$GetItem->ItemIds->ItemId = $item->ItemId;
			$response = $this->client->GetItem($GetItem);
			
			if($response->ResponseMessages->GetItemResponseMessage->ResponseCode != "NoError") {
				$this->lastError = $response->ResponseMessages->GetItemResponseMessage->ResponseCode;
				return false;
			}
			
			$messageobj = $response->ResponseMessages->GetItemResponseMessage->Items->Message;
			if($onlyunread && $messageobj->IsRead)
				continue;

			$newmessage = null;
			$newmessage->bodytext = $messageobj->Body->_;
			$newmessage->bodytype = $messageobj->Body->BodyType;
			$newmessage->isread = $messageobj->IsRead;
			$newmessage->ItemId = $item->ItemId;
			$newmessage->from = $messageobj->From->Mailbox->EmailAddress;
			$newmessage->from_name = $messageobj->From->Mailbox->Name;
			$newmessage->subject = $messageobj->Subject;
			
			$messages[] = $newmessage;
			
			$i++;
			if($i > $limit)
				break;
		}
		
		$this->teardown();
		
		return $messages;
		
	}
	
	/**
	 * Send a message through the Exchange server as the currently logged-in user.
	 * 
	 * @access public
	 * @param string $to (the email address to send the message to)
	 * @param string $subject
	 * @param string $content
	 * @param string $bodytype. (default: "Text", "HTML" for HTML emails)
	 * @param bool $saveinsent. (Save in the user's sent folder after sending? default: true)
	 * @param bool $markasread. (Mark as read after sending? This currently does nothing. default: true)
	 * @return bool $success. (True if the message was sent, false if there was an error).
	 */
	function send_message($to, $subject, $content, $bodytype="Text", $saveinsent=true, $markasread=true) {
		$this->setup();
		
		if($saveinsent) {
			$CreateItem->MessageDisposition = "SendOnly";
			$CreateItem->SavedItemFolderId->DistinguishedFolderId->Id = "sentitems";
		}
		else
			$CreateItem->MessageDisposition = "SendOnly";
		
		$CreateItem->Items->Message->ItemClass = "IPM.Note";
		$CreateItem->Items->Message->Subject = $subject;
		$CreateItem->Items->Message->Body->BodyType = $bodytype;
		$CreateItem->Items->Message->Body->_ = $content;
		$CreateItem->Items->Message->ToRecipients->Mailbox->EmailAddress = $to;
		
		if($markasread)
			$CreateItem->Items->Message->IsRead = "true";
		
		$response = $this->client->CreateItem($CreateItem);
		
		$this->teardown();
		
		if($response->ResponseMessages->CreateItemResponseMessage->ResponseCode == "NoError")
			return true;
		else {
			$this->lastError = $response->ResponseMessages->CreateItemResponseMessage->ResponseCode;
			return false;
		}
		
	}
	
	/**
	 * Deletes a message in the mailbox of the current user.
	 * 
	 * @access public
	 * @param ItemId $ItemId (such as one returned by get_messages)
	 * @param string $deletetype. (default: "HardDelete")
	 * @return bool $success (true: message was deleted, false: message failed to delete)
	 */
	function delete_message($ItemId, $deletetype="HardDelete") {
		$this->setup();
		
		$DeleteItem->DeleteType = $deletetype;
		$DeleteItem->ItemIds->ItemId = $ItemId;
		
		$response = $this->client->DeleteItem($DeleteItem);
		
		$this->teardown();
		
		if($response->ResponseMessages->DeleteItemResponseMessage->ResponseCode == "NoError")
			return true;
		else {
			$this->lastError = $response->ResponseMessages->DeleteItemResponseMessage->ResponseCode;
			return false;
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
			
		stream_wrapper_unregister('http');
		stream_wrapper_register('http', 'ExchangeNTLMStream') or die("Failed to register protocol");
		
	}
	
	/**
	 * Tears down stream handling. Internally used.
	 * 
	 * @access private
	 * @return void
	 */
	function teardown() {
		stream_wrapper_restore('http');
	}
}

class ImpersonationHeader {

	var $ConnectingSID;

	function __construct($email) {
		$this->ConnectingSID->PrimarySmtpAddress = $email;
	}

}

class NTLMSoapClient extends SoapClient {
	function __doRequest($request, $location, $action, $version) {
		//print_r($request);
		//print($location);
		$headers = array(
			'Method: POST',
			'Connection: Keep-Alive',
			'User-Agent: PHP-SOAP-CURL',
			'Content-Type: text/xml; charset=utf-8',
			'SOAPAction: "'.$action.'"',
		);  
		$this->__last_request_headers = $headers;
		$ch = curl_init($location);
		curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
		curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);
		curl_setopt($ch, CURLOPT_POST, true );
		curl_setopt($ch, CURLOPT_POSTFIELDS, $request);
		curl_setopt($ch, CURLOPT_HTTP_VERSION, CURL_HTTP_VERSION_1_1);
		curl_setopt($ch, CURLOPT_HTTPAUTH, CURLAUTH_NTLM);
		curl_setopt($ch, CURLOPT_USERPWD, $this->user.':'.$this->password);
		curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
		curl_setopt($ch, CURLOPT_SSL_VERIFYHOST, false);
		curl_setopt($ch, CURLOPT_VERBOSE, true);
		$response = curl_exec($ch);
		//print("RESPONSE: ");
		//print_r($response);
		return $response;
	}   
	function __getLastRequestHeaders() {
		return implode("n", $this->__last_request_headers)."n";
	}   
}

class ExchangeNTLMSoapClient extends NTLMSoapClient {
	public $user = '';
	public $password = '';
}

class NTLMStream {
	private $path;
	private $mode;
	private $options;
	private $opened_path;
	private $buffer;
	private $pos;

	public function stream_open($path, $mode, $options, $opened_path) {
		echo "[NTLMStream::stream_open] $path , mode=$mode n";
		$this->path = $path;
		$this->mode = $mode;
		$this->options = $options;
		$this->opened_path = $opened_path;
		$this->createBuffer($path);
		return true;
	}

	public function stream_close() {
		echo "[NTLMStream::stream_close] n";
		curl_close($this->ch);
	}

	public function stream_read($count) {
		echo "[NTLMStream::stream_read] $count n";
		if(strlen($this->buffer) == 0) {
			return false;
		}
		$read = substr($this->buffer,$this->pos, $count);
		$this->pos += $count;
		return $read;
	}

	public function stream_write($data) {
		echo "[NTLMStream::stream_write] n";
		if(strlen($this->buffer) == 0) {
			return false;
		}
		return true;
	}

	public function stream_eof() {
		echo "[NTLMStream::stream_eof] ";
		if($this->pos > strlen($this->buffer)) {
			echo "true n";
			return true;
		}
		echo "false n";
		return false;
	}

	/* return the position of the current read pointer */
	public function stream_tell() {
		echo "[NTLMStream::stream_tell] n";
		return $this->pos;
	}

	public function stream_flush() {
		echo "[NTLMStream::stream_flush] n";
		$this->buffer = null;
		$this->pos = null;
	}

	public function stream_stat() {
		echo "[NTLMStream::stream_stat] n";
		$this->createBuffer($this->path);
		$stat = array(
			'size' => strlen($this->buffer),
		);
		return $stat;
	}

	public function url_stat($path, $flags) {
		echo "[NTLMStream::url_stat] n";
		$this->createBuffer($path);
		$stat = array(
			'size' => strlen($this->buffer),
		);
		return $stat;
	}

	/* Create the buffer by requesting the url through cURL */
	private function createBuffer($path) {
		if($this->buffer) {
			return;
		}
		echo "[NTLMStream::createBuffer] create buffer from : $pathn";
		$this->ch = curl_init($path);
		curl_setopt($this->ch, CURLOPT_RETURNTRANSFER, true);
		curl_setopt($this->ch, CURLOPT_HTTP_VERSION, CURL_HTTP_VERSION_1_1);
		curl_setopt($this->ch, CURLOPT_HTTPAUTH, CURLAUTH_NTLM);
		curl_setopt($this->ch, CURLOPT_USERPWD, $this->user.':'.$this->password);
		echo $this->buffer = curl_exec($this->ch);
		echo "[NTLMStream::createBuffer] buffer size : ".strlen($this->buffer)."bytesn";
		$this->pos = 0;
	}
}

class ExchangeNTLMStream extends NTLMStream {
	//protected $user = '';
	//protected $password = '';
}
