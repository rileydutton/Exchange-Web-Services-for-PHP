<?php

class ExchangeNTLMStream extends NTLMStream {
  protected function createBuffer($path) { 
    parent::createBuffer($path);

    if (!preg_match('#Services\.wsdl$#', $path))
      return;

    $xml = simplexml_load_string($this->buffer);

    if (!$xml instanceof SimpleXMLElement)
      return;

    $service = $xml->addChild('wsdl:service');
    $service->addAttribute('name', 'ExchangeServices');
    
    $port = $service->addChild('wsdl:port');
    $port->addAttribute('name', 'ExchangeServicePort');
    $port->addAttribute('binding', 'tns:ExchangeServiceBinding');

    $address = $port->addChild('soap:address', null, 'soap');
    $address->addAttribute('location', str_replace('Services.wsdl', 'Exchange.asmx', $path));

    $this->buffer = str_replace('xmlns:soap="soap" ', '', $xml->asXML());
  }
}

