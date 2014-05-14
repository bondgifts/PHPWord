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
	'pageSizeW' => $twip * 7, 
	'pageSizeH' => $twip * 5,
	'marginTop' => $twip * 0.25, // a la border-box model
	'marginRight' => $twip * 0.25,
	'marginBottom' => $twip * 0.25,
	'marginLeft' => $twip * 0.25
));

$actualWidth = $section->getSettings()->getPageSizeW() - $twip * 0.50;
$actualHeight = $section->getSettings()->getPageSizeH() - $twip * 0.50;

$textboxOne = $section->addTextbox(array(
	'offsetX' => twip2emu($twip * 0),
	'offsetY' => twip2emu($twip * 0),
	'width' => twip2emu($twip * 2),
	'height' => twip2emu($twip * 1)
));

$textRun = $textboxOne->createTextRun(array('align' => 'left'));
$textRun->addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit. In ultrices leo metus, id porta augue blandit id.', array('name' => 'astJohnsonCheryl', 'size' => 34));




$textboxTwo = $section->addTextbox(array(
	'offsetX' => twip2emu($actualWidth) - twip2emu($twip * 5), //right
	'offsetY' => twip2emu($actualHeight) - twip2emu($twip * 1), //bottom
	'width' => twip2emu($twip * 5),
	'height' => twip2emu($twip * 1)
));

$textRun = $textboxTwo->createTextRun(array('align' => 'right'));
$textRun->addText('In sed mattis dui, ut iaculis lectus. Quisque tincidunt bibendum sodales. Cras et ullamcorper diam, eu fermentum dui. Cras bibendum magna eleifend tellus cursus sollicitudin.', array('name' => 'astUribe', 'size' => 24));




$textboxThree = $section->addTextbox(array(
	'offsetX' => twip2emu($twip * 1),
	'offsetY' => twip2emu($twip * 1),
	'width' => twip2emu($twip * 5),
	'height' => twip2emu($twip * 3)
));

$textRun = $textboxThree->createTextRun(array('align' => 'center'));
$textRun->addText('Cras et ullamcorper diam.', array('name' => 'astDunn', 'size' => 48));

// Save file
echo write($phpWord, basename(__FILE__, '.php'), $writers);
$doc = basename(__FILE__, '.php').'.docx';

copy('/Users/larrylaski/Sites/PHPWord/samples/results/'.$doc, '/Users/larrylaski/Downloads/'.$doc);
copy('/Users/larrylaski/Downloads/'.$doc, '/Users/larrylaski/Downloads/'.str_replace('docx', 'zip', $doc));
if (!CLI) {
    include_once 'Sample_Footer.php';
}
