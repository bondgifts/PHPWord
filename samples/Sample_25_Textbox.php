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
	'pageSizeW' => $twip * 4,
	'pageSizeH' => $twip * 6,
	'marginTop' => $twip * 0.25,
	'marginRight' => $twip * 0.25,
	'marginBottom' => $twip * 0.25,
	'marginLeft' => $twip * 0.25
));

//Textbox needs to go here
//Goal is to have two pieces of text next to each other (start with 2 columsn basically)



$textboxOne = $section->addTextbox(array(
	'offsetX' => 0,
	'offsetY' => twip2emu($twip * 1),
	'width' => twip2emu($twip * 2),
	'height' => twip2emu($twip * 2)
));

$textRunBold = $textboxOne->createTextRun();
$textRunBold->addText('Hey there!!', array('name' => 'Times New Roman', 'size' => 20, 'bold' => true));
$textRunBold->addTextBreak(2);
$textRunBold->addText('This is the rest of this text', array('name' => 'Times New Roman', 'size' => 14, 'italic' => true));

//need to fix doc part and docpart id
// pp($section);
// pp($textboxOne);
// pp($text);
// exit;
// pp($textboxOne);
// exit;
$textboxTwo = $section->addTextbox(array(
	'offsetX' => twip2emu($twip * 2),
	'offsetY' => twip2emu($twip * 1),
	'width' => twip2emu($twip * 2),
	'height' => twip2emu($twip * 2)
));
$text = $textboxTwo->addText('And Lonny has cookies??');

$textboxThree = $section->addTextbox(array(
	'offsetX' => twip2emu($twip * 3),
	'offsetY' => twip2emu($twip * 2),
	'width' => twip2emu($twip * 2),
	'height' => twip2emu($twip * 2)
));
$text = $textboxThree->addText('Dessert overload!');
//Each textbox needs to have its own width, height, and offset from left and top of section


// $textRun->addText('Each textrun can contain native text, link elements or an image.');


// Save file
echo write($phpWord, basename(__FILE__, '.php'), $writers);
$doc = basename(__FILE__, '.php').'.docx';
copy('/Users/larrylaski/Sites/word/samples/results/'.$doc, '/Users/larrylaski/Downloads/'.$doc);
copy('/Users/larrylaski/Downloads/'.$doc, '/Users/larrylaski/Downloads/'.str_replace('docx', 'zip', $doc));
if (!CLI) {
    include_once 'Sample_Footer.php';
}
