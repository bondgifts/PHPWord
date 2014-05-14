<?php

include_once 'Sample_Header.php';

// 1440 TWIP == 1 Inch
$twip = 1440;

// 914400 EMU = 1 Inch
$emu = 914400;

// 914400 EMU = 1440 TWIP
// 635 EMU = 1 TWIP
function twip2emu($val)
{
	return $val * 635;
}

// New Word Document
echo date('H:i:s') , " Create new PhpWord bond test object" , EOL;
$phpWord = new \PhpOffice\PhpWord\PhpWord();

$section = $phpWord->createSection(array(
	'pageSizeW' => $twip * 5,
	'pageSizeH' => $twip * 7,
	'marginTop' => $twip * 0.25,
	'marginRight' => $twip * 0.25,
	'marginBottom' => $twip * 0.25,
	'marginLeft' => $twip * 0.25
));

//Textbox needs to go here
//Goal is to have two pieces of text next to each other (start with 2 columsn basically)



$textboxOne = $section->addTextbox(array(
	'offsetX' => twip2emu($twip * 0),
	'offsetY' => twip2emu($twip * 0),
	'width' => twip2emu($twip * 2),
	'height' => twip2emu($twip * 2)
));

$textRunBold = $textboxOne->createTextRun(array('align' => 'justify'));
$textRunBold->addText('Text in the top left corner! Text in the top left corner!', array('name' => 'astJohnsonCheryl', 'size' => 20, 'bold' => true));
// $textRunBold->addTextBreak(2);
// $textRunBold->addText('This is the rest of this text', array('name' => 'astJohnsonCheryl', 'size' => 14, 'italic' => true));

$textboxTwo = $section->addTextbox(array(
	'offsetX' => twip2emu($twip * 3),
	'offsetY' => twip2emu($twip * 0),
	'width' => twip2emu($twip * 2),
	'height' => twip2emu($twip * 2)
));
$text = $textboxTwo->addText('lots and lots and lots of Text in the top right corner!', array('name' => 'astDunn', 'size' => 15 ), array('align' => 'justify'));

$textboxThree = $section->addTextbox(array(
	'offsetX' => twip2emu($twip * 0),
	'offsetY' => twip2emu($twip * 5),
	'width' => twip2emu($twip * 2),
	'height' => twip2emu($twip * 2)
));
$text = $textboxThree->addText('Text in the bottom left corner! Text in the bottom left corner!', array('name' => 'astRossi', 'size' => 30), array('align' => 'left'));

$textboxFour = $section->addTextbox(array(
	'offsetX' => twip2emu($twip * 3),
	'offsetY' => twip2emu($twip * 5),
	'width' => twip2emu($twip * 2),
	'height' => twip2emu($twip * 2)
));
$text = $textboxFour->addText('Text in the bottom right corner! Text in the bottom right corner!', array('name' => 'astRaymond', 'size' => 20), array('align' => 'right'));

$textboxFive = $section->addTextbox(array(
	'offsetX' => twip2emu($twip * 1),
	'offsetY' => twip2emu($twip * 2.5),
	'width' => twip2emu($twip * 3),
	'height' => twip2emu($twip * 1)
));
$textboxFive->addText('Magic Man', array('name' => 'astRossi', 'size' => 35), array('align' => 'center'));
$textRunBold->addTextBreak(2);
$textboxFive->addText('Middle Text!', array('name' => 'astRossi', 'size' => 35, 'italic' => true), array('align' => 'center'));

//Each textbox needs to have its own width, height, and offset from left and top of section

// $textRun->addText('Each textrun can contain native text, link elements or an image.');

// Save file
echo write($phpWord, basename(__FILE__, '.php'), $writers);
$doc = basename(__FILE__, '.php').'.docx';

copy('/Users/larrylaski/Sites/PHPWord/samples/results/'.$doc, '/Users/larrylaski/Downloads/'.$doc);
copy('/Users/larrylaski/Downloads/'.$doc, '/Users/larrylaski/Downloads/'.str_replace('docx', 'zip', $doc));
if (!CLI) {
    include_once 'Sample_Footer.php';
}
