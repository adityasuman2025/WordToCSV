<?php
 //set headers to download file rather than displayed
	$input_file_name = "File 3.docx";
    $output_filename = "output3" . ".csv";

    header('Content-Type: text/csv');
    header('Content-Disposition: attachment; filename="' . $output_filename . '";');

  	class DocxConversion
  	{
	    private $filename;

	    public function __construct($filePath) {
	        $this->filename = $filePath;
	    }

	    private function read_doc() {
	        $fileHandle = fopen($this->filename, "r");
	        $line = @fread($fileHandle, filesize($this->filename));   
	        $lines = explode(chr(0x0D),$line);
	        $outtext = "";
	        foreach($lines as $thisline)
	          {
	            $pos = strpos($thisline, chr(0x00));
	            if (($pos !== FALSE)||(strlen($thisline)==0))
	              {
	              } else {
	                $outtext .= $thisline." ";
	              }
	          }
	         $outtext = preg_replace("/[^a-zA-Z0-9\s\,\.\-\n\r\t@\/\_\(\)]/","",$outtext);
	        return $outtext;
	    }

	    private function read_docx(){
	    //1st algo	
	        // $striped_content = '';
	        // $content = '';

	        // $zip = zip_open($this->filename);

	        // if (!$zip || is_numeric($zip)) return false;

	        // while ($zip_entry = zip_read($zip)) {

	        //     if (zip_entry_open($zip, $zip_entry) == FALSE) continue;

	        //     if (zip_entry_name($zip_entry) != "word/document.xml") continue;

	        //     $content = $content . "" . zip_entry_read($zip_entry, zip_entry_filesize($zip_entry));

	        //     zip_entry_close($zip_entry);
	        // }// end while

	        // zip_close($zip);

	        // $content = str_replace('</w:r></w:p></w:tc><w:tc>', " ", $content);
	        // $content = str_replace('</w:r></w:p>', "\r\n", $content);
	        // $striped_content = strip_tags($content);

	        // return $striped_content;

	    //2nd algo	
			$docx = $this->filename;
	        $xml_filename = "word/document.xml"; //content file name
	        $zip_handle = new ZipArchive;
	        
	        $output_text = "";

	        if(true === $zip_handle->open($docx))
	        {
	            if(($xml_index = $zip_handle->locateName($xml_filename)) !== false)
	            {
	                $xml_datas = trim($zip_handle->getFromIndex($xml_index));
	                $replace_newlines = trim(preg_replace('/<w:p w[0-9-Za-z]+:[a-zA-Z0-9]+="[a-zA-z"0-9 :="]+">/',"\n\r",$xml_datas));
	                $replace_tableRows = trim(preg_replace('/<w:tr>/',"\n\r",$replace_newlines));
	                $replace_tab = trim(preg_replace('/<w:tab\/>/',"\t",$replace_tableRows));
	                $replace_paragraphs = trim(preg_replace('/<\/w:p>/',"\n\r",$replace_tab));
	                $replace_other_Tags = trim(strip_tags($replace_paragraphs));

	                $output_text = trim($replace_other_Tags);
	            }
	            else
	            {
	                $output_text .="";
	            }

	            $zip_handle->close();
	        }
	        else
	        {
	        	$output_text .="";
	        }
	      
	        return $output_text;
	    }

	    public function convertToText() 
	    {
	        if(isset($this->filename) && !file_exists($this->filename)) {
	            return "File Not exists";
	        }

	        $fileArray = pathinfo($this->filename);
	        $file_ext  = $fileArray['extension'];
	        if($file_ext == "doc" || $file_ext == "docx" || $file_ext == "xlsx" || $file_ext == "pptx")
	        {
	            if($file_ext == "doc") {
	                return $this->read_doc();
	            } elseif($file_ext == "docx") {
	                return $this->read_docx();
	            } elseif($file_ext == "xlsx") {
	                return $this->xlsx_to_text();
	            }elseif($file_ext == "pptx") {
	                return $this->pptx_to_text();
	            }
	        } else {
	            return "Invalid File Type";
	        }
	    }

	}

	$docObj = new DocxConversion($input_file_name);
	$docText= $docObj->convertToText();

//breaking each line ans storing in an array
	$broken = explode("\n", $docText);
	
//trimming/cleaning each elements of array and creating group array
	$final_array = array();

	$temp_array = array();
	foreach ($broken as $key => $value) 
	{
		$val = trim($value);
		if($val != "")
		{
			if($val == "\n" || $val == "\t")
				continue;
			else
			{
				array_push($temp_array, $val);

				if($val == "RELATED STUDENTS")
				{
				//if any group contains addaress too then reomving it	
					if(in_array("HOME", $temp_array))
					{
						$len = sizeof($temp_array);
						array_splice($temp_array, $len-3, 2);

						$temp_house = trim(strtolower($temp_array[$len - 4]));
						if(strpos($temp_house, 'household') !== false) 
						{
							// echo 'true';
						}
						else
						{
							$len = sizeof($temp_array);
							array_splice($temp_array, $len-2, 1);
						}
					}

					array_push($final_array, $temp_array);
					$temp_array = array();
				}
			}
		}
	}
	array_push($final_array, $temp_array);

//creating group array
	$group_arr_size = sizeof($final_array);

	foreach ($final_array as $key => $array) 
	{
	//getting the last two elements of this group array
		$this_arr_len = sizeof($array);
		$comp_name = $array[$this_arr_len-2];
		$rel_stud = $array[$this_arr_len-1];

	//adding the last two elements of this group array, at front of the next group array
		if($key + 1 < $group_arr_size)
		{
			array_unshift($final_array[$key+1], $rel_stud);
			array_unshift($final_array[$key+1], $comp_name);
		}
		
	//removing last two elements of this group array except for the last group array
		if($key != $group_arr_size - 1)
		{
			$new_len = sizeof($final_array[$key]);
			unset($final_array[$key][$new_len -2]);
			unset($final_array[$key][$new_len -1]);	
		}
	}

//making associateive_array
	$assoc_group_arr = array();
	foreach ($final_array as $key => $array) 
	{
		$arr_len = sizeof($array);

		$temp_array = array();
		if($arr_len > 0)
		{
			$temp_array['comp_name'] = $array[0];
			$temp_array['many_name'] = $array[2];

			if($arr_len > 3) //if name 1 exists
			{
				$temp_array['name_1'] = $array[3];	

				if($arr_len > 7) //if email 1 exists
					$temp_array['email_1'] = $array[7];
			}

			if($arr_len > 8) //if name 2 exists
			{
				$temp_array['name_2'] = $array[8];

				if($arr_len > 12) //if email 2 exists
				{
					$temp_array['email_2'] = $array[12];
				}
			}
			array_push($assoc_group_arr, $temp_array);
		}
	}

	// echo "<pre>";
	// print_r($assoc_group_arr);
	// echo "</pre>";

//creating csv file
	$delimiter = ",";

//create a file pointer
	$f = fopen('php://memory', 'w');

//set column headers
	// $fields = array('emp_id', 'name', 'email', 'person_type', 'consent', 'comment', "contrbuted_on");
	// fputcsv($f, $fields, $delimiter);

//output each row of the data, format line as csv and write to file pointer
	foreach ($assoc_group_arr as $key => $temp_array) 
	{
		if(array_key_exists("comp_name", $temp_array))
		{
			$comp_name = $temp_array['comp_name'];

			if(array_key_exists("name_1", $temp_array))
				$name_1 = $temp_array['name_1'];
			else
				$name_1 = "--";

			if(array_key_exists("email_1", $temp_array))
				$email_1 = $temp_array['email_1'];
			else
				$email_1 = "--";

			if(array_key_exists("name_2", $temp_array))
				$name_2 = $temp_array['name_2'];
			else
				$name_2 = "--";

			if(array_key_exists("email_2", $temp_array))
				$email_2 = $temp_array['email_2'];
			else
				$email_2 = "--";

			$lineData = @array($comp_name, $name_1, $email_1, $name_2, $email_2);
			@fputcsv($f, $lineData, $delimiter);
		}
	}

//move back to beginning of file
	fseek($f, 0);

//output all remaining data on a file pointer
	fpassthru($f);
?>