<?php

/*

Copyright (c) 2009 Dimas Begunoff, http://www.farinspace.com/

Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

*/

class Google_Spreadsheet_Exception extends Exception{
	
}

class Google_Spreadsheet{
	
	protected 
		$_client,
		$_spreadsheet = null,
		$_spreadsheet_id = null,
		$_worksheet = 'Sheet1',
		$_worksheet_id;

	public function __construct($user=null, $pass=null, $ss=null, $ws=null){
		if (!is_null($user) && !is_null($pass)) $this->login($user, $pass);
		if (!is_null($ss)) $this->useSpreadsheet($ss);
		if (!is_null($ws)) $this->useWorksheet($ws);
	}
	
	public function useSpreadsheet($ss, $ws=null){
		$this->setSpreadsheetId(null);
		$this->_spreadsheet = $ss;
		if (!is_null($ws)) $this->useWorksheet($ws);
	}
	
	public function setSpreadsheetId($id){
		$this->_spreadsheet_id = $id;
	}

	public function useWorksheet($ws){
		$this->_worksheet = $ws;
		$this->setWorksheetId(null);
	}

	public function setWorksheetId($id){
		$this->_worksheet_id = $id;
	}
	
	public function addRow($row){
		
		if ($this->_client instanceof Zend_Gdata_Spreadsheets){
			
			$ss_id = $this->_getSpreadsheetId($this->_spreadsheet);
			
			$ws_id = $this->_getWorksheetId($ss_id, $this->_worksheet);
			
			$insert_row = array();
			
			foreach ($row as $k => $v) $insert_row[$this->_cleanKey($k)] = $v;
			
			$entry = $this->_client->insertRow($insert_row, $ss_id, $ws_id);
			
			if ($entry instanceof Zend_Gdata_Spreadsheets_ListEntry) return true;
		}
		
		throw new Google_Spreadsheet_Exception('Unable to add row to the spreadsheet');
	}

	// http://code.google.com/apis/spreadsheets/docs/2.0/reference.html#ListParameters
	public function updateRow($row, $search){
		if ($this->_client instanceof Zend_Gdata_Spreadsheets AND $search){
			$feed = $this->_findRows($search);
			
			if ($feed->entries){
				foreach($feed->entries as $entry){
					
					if ($entry instanceof Zend_Gdata_Spreadsheets_ListEntry){
						$update_row = array();
						
						$customRow = $entry->getCustom();
						foreach ($customRow as $customCol){
							$update_row[$customCol->getColumnName()] = $customCol->getText();
						}
						
						// overwrite with new values
						foreach ($row as $k => $v){
							$update_row[$this->_cleanKey($k)] = $v;
						}
						
						// update row data, then save
						$entry = $this->_client->updateRow($entry, $update_row);
						if ( ! ($entry instanceof Zend_Gdata_Spreadsheets_ListEntry)) return false;
					}
				}
				
				return true;
			}
		}
		
		return false;
	}

	// http://code.google.com/apis/spreadsheets/docs/2.0/reference.html#ListParameters
	function getRows($search=false){
		$rows = array();
		
		if ($this->_client instanceof Zend_Gdata_Spreadsheets){
			$feed = $this->_findRows($search);
			
			if ($feed->entries){
				foreach($feed->entries as $entry){
					if ($entry instanceof Zend_Gdata_Spreadsheets_ListEntry){
						$row = array();
						
						$customRow = $entry->getCustom();
						foreach ($customRow as $customCol){
							$row[$customCol->getColumnName()] = $customCol->getText();
						}
						
						$rows[] = $row;
					}
				}
			}
		}
		
		return $rows;
	}

	// user contribution by dmon (6/10/2009)
	public function deleteRow($search){
		if ($this->_client instanceof Zend_Gdata_Spreadsheets AND $search){
			$feed = $this->_findRows($search);
			
			if ($feed->entries){
				foreach($feed->entries as $entry){
					
					if ($entry instanceof Zend_Gdata_Spreadsheets_ListEntry){
						$this->_client->deleteRow($entry);
						
						if ( ! ($entry instanceof Zend_Gdata_Spreadsheets_ListEntry)) return false;
					}
				}
				return true;
			}
		}
		return false;
	}

	public function getColumnNames(){
		$query = new Zend_Gdata_Spreadsheets_ListQuery();
		$query->setSpreadsheetKey($this->_getSpreadsheetId());
		$query->setWorksheetId($this->_getWorksheetId());
		$query->setMaxResults(1);
		$query->setStartIndex(1);
		
		$feed = $this->_client->getListFeed($query);
		
		$data = array();
		
		if ($feed->entries){
			foreach($feed->entries as $entry){
				if ($entry instanceof Zend_Gdata_Spreadsheets_ListEntry){
					$customRow = $entry->getCustom();
					
					foreach ($customRow as $customCol){
						array_push($data,$customCol->getColumnName());
					}
				}
			}
		}

		return $data;
	}

	public function login($user, $pass){
		// Zend Gdata package required
		// http://framework.zend.com/download/gdata
		
		require_once 'Zend/Loader.php';
		Zend_Loader::loadClass('Zend_Http_Client');
		Zend_Loader::loadClass('Zend_Gdata');
		Zend_Loader::loadClass('Zend_Gdata_ClientLogin');
		Zend_Loader::loadClass('Zend_Gdata_Spreadsheets');
		
		$service = Zend_Gdata_Spreadsheets::AUTH_SERVICE_NAME;
		$http = Zend_Gdata_ClientLogin::getHttpClient($user, $pass, $service);
		
		$this->_client = new Zend_Gdata_Spreadsheets($http);
		
		if ($this->_client instanceof Zend_Gdata_Spreadsheets) return true;
		
		throw new Google_Spreadsheet_Exception('Login failed, incorrect credentials?');
	}

	private function _findRows($search=false){
		$query = new Zend_Gdata_Spreadsheets_ListQuery();
		$query->setSpreadsheetKey($this->_getSpreadsheetId());
		$query->setWorksheetId($this->_getWorksheetId());
		
		if ($search) $query->setSpreadsheetQuery($search);
		
		$feed = $this->_client->getListFeed($query);
		
		return $feed;
	}

	/**
	 * Lookup spreadsheet id, if $this->_spreadsheet_id is already set return $this->_spreadsheet_id
	 *
	 * @param string $ss 
	 * @return string $ss_id
	 * @author Koen Punt
	 */
	private function _getSpreadsheetId($ss = null){
		if ($this->_spreadsheet_id) return $this->_spreadsheet_id;
		
		$ss = is_null($ss) ? $this->_spreadsheet : $ss;
		
		$ss_id = false;
		
		$feed = $this->_client->getSpreadsheetFeed();
		
		foreach($feed->entries as $entry){
			if ($entry->title->text == $ss){
				$ss_id = array_pop(explode("/", $entry->id->text));
				
				$this->_spreadsheet_id = $ss_id;
				
				break;
			}
		}
		if (!$ss_id) throw new Google_Spreadsheet_Exception('Unable to find spreadsheet by name: "' . $this->_spreadsheet . '", confirm the name of the spreadsheet');
		return $ss_id;
	}

	/**
	 * Lookup worksheet id, if $this->_worksheet_id is already set return $this->_worksheet_id
	 *
	 * @param string $ss 
	 * @param string $ws 
	 * @return string $ws_id
	 * @author Koen Punt
	 */
	private function _getWorksheetId($ss_id = null, $ws = null){
		if ($this->_worksheet_id) return $this->_worksheet_id;
		
		$ss_id = is_null($ss_id) ? $this->_spreadsheet_id : $ss_id;
		
		$ws = is_null($ws) ? $this->_worksheet : $ws;
		
		$ws_id = false;
		
		if ($ss_id and $ws){
			$query = new Zend_Gdata_Spreadsheets_DocumentQuery();
			$query->setSpreadsheetKey( $ss_id );
			$feed = $this->_client->getWorksheetFeed( $query );
			
			foreach($feed->entries as $entry){
				if ($entry->title->text == $ws){
					$ws_id = array_pop(explode("/", $entry->id->text));
					
					$this->_worksheet_id = $ws_id;
					
					break;
				}
			}
		}
		if (!$ws_id) throw new Google_Spreadsheet_Exception('Unable to find worksheet by name: "' . $this->_worksheet . '", confirm the name of the worksheet');
		return $ws_id;
	}

	private function _cleanKey($k){
		return strtolower(preg_replace('/[^A-Za-z0-9\-\.]+/','',$k));
	}
}
