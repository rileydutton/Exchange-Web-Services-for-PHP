<?php
 
class NTLMSoapClient extends SoapClient {
	private $username;
	private $password;

	public function __construct($wsdl, array $options = array()) { 
		parent::__construct($wsdl, $options);

		if (array_key_exists('login', $options))
			$this->username = $options['login'];

		if (array_key_exists('password', $options))
			$this->password = $options['password'];
	}
	
	public function __doRequest($request, $location, $action, $version, $one_way = 0) {
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
		curl_setopt($ch, CURLOPT_HTTPAUTH, CURLAUTH_NTLM | CURLAUTH_BASIC);
		curl_setopt($ch, CURLOPT_USERPWD, $this->username.':'.$this->password);
		curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
		curl_setopt($ch, CURLOPT_SSL_VERIFYHOST, false);

    return curl_exec($ch);
	}   

	public function __getLastRequestHeaders() {
		return implode("n", $this->__last_request_headers)."n";
	}   
}

