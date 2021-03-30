<?php
/** 
 * Converting document microsoft word journal to various template university in indonesia
 * using COM (Component Object Model) extension in php for connecting word file with php, for 
 * reference this project :
 * @link https://docs.microsoft.com/en-us/office/vba/api/overview/ VBA Reference
 * @link https://www.php.net/manual/en/class.com.php COM php extension
 * @link https://bettersolutions.com/word/styles/vba-code.htm another vba code
 * 
 * @author arif iskandar <aripdevel@gmail.com>
 * 
 */

/**
 * Default Configuration
 */

 $config['font-name']="Times New Roman";

$word=new COM('word.application');

/** 
 * For development turn on
 */

$word->Visible=1;

/** 
 * Open document sample
 */

$word->Documents->Open(getcwd()."/sample.doc");

/**
 * Set default font name for writing
 */

$word->Selection->Font->Name=$config['font-name'];

/**
 * Get Title for required standar journal
 */

$title=$word->ActiveDocument->Styles("Title");

$title->Font->Size="16";
$title->Font->Bold=True;
$title->Font->Name=$config['font-name'];
$title->Borders->Enable=False;
$title->ParagraphFormat->Alignment=1;








