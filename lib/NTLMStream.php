<?php

/*
 * Copyright (c) 2008 Invest-In-France Agency http://www.invest-in-france.org
 *
 * Author : Thomas Rabaix 
 *
 * Permission to use, copy, modify, and distribute this software for any
 * purpose with or without fee is hereby granted, provided that the above
 * copyright notice and this permission notice appear in all copies.
 *
 * THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
 * WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
 * MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
 * ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
 * WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
 * ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
 * OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
 */ 
class NTLMStream {
	private $path;
	private $mode;
	private $options;
	private $opened_path;
	private $pos;

	protected $buffer;
	protected static $username;
	protected static $password;

	public static function setCredentials($username, $password) {
		self::$username = $username;
		self::$password = $password;
	}

	public function stream_open($path, $mode, $options, $opened_path) {
		$this->path = $path;
		$this->mode = $mode;
		$this->options = $options;
		$this->opened_path = $opened_path;
		$this->createBuffer($path);

		return true;
	}

	public function stream_close() {
		curl_close($this->ch);
	}

	public function stream_read($count) {
		if(strlen($this->buffer) == 0) {
			return false;
		}

		$read = substr($this->buffer,$this->pos, $count);
		$this->pos += $count;

		return $read;
	}

	public function stream_write($data) {
		if(strlen($this->buffer) == 0) {
			return false;
		}

		return true;
	}

	public function stream_eof() {
    return $this->pos > strlen($this->buffer);
	}

	public function stream_tell() {
		return $this->pos;
	}

	public function stream_flush() {
		$this->buffer = null;
		$this->pos = null;
	}

	public function stream_stat() {
		return $this->url_stat($this->path, 0);
	}

	public function url_stat($path, $flags) {
		$this->createBuffer($path);

		$stat = array(
			'size' => strlen($this->buffer),
		);

		return $stat;
	}

	/* Create the buffer by requesting the url through cURL */
	protected function createBuffer($path) {
		if($this->buffer) {
			return;
		}

		$this->ch = curl_init($path);
		curl_setopt($this->ch, CURLOPT_RETURNTRANSFER, true);
		curl_setopt($this->ch, CURLOPT_HTTP_VERSION, CURL_HTTP_VERSION_1_1);
		curl_setopt($this->ch, CURLOPT_HTTPAUTH, CURLAUTH_NTLM | CURLAUTH_BASIC);
		curl_setopt($this->ch, CURLOPT_USERPWD, self::$username.':'.self::$password);

    $this->buffer = curl_exec($this->ch);
		$this->pos = 0;
	}
}

