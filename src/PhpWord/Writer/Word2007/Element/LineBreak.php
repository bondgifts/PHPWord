<?php
/**
 * PHPWord
 *
 * @link        https://github.com/PHPOffice/PHPWord
 * @copyright   2014 PHPWord
 * @license     http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt LGPL
 */

namespace PhpOffice\PhpWord\Writer\Word2007\Element;

use PhpOffice\PhpWord\Writer\Word2007\Style\Font as FontStyleWriter;
use PhpOffice\PhpWord\Writer\Word2007\Style\Paragraph as ParagraphStyleWriter;

/**
 * LineBreak element writer
 *
 * @since 0.10.0
 */
class LineBreak extends Element
{
	/**
	 * Write text element
	 */
	public function write()
	{
		if (!$this->withoutP) {
			$styleWriter = new ParagraphStyleWriter($this->xmlWriter, $this->element->getParagraphStyle());
			$styleWriter->setIsInline(true);

			$this->xmlWriter->startElement('w:p');
			$styleWriter->write();
		}
		$styleWriter = new FontStyleWriter($this->xmlWriter, $this->element->getFontStyle());
		$styleWriter->setIsInline(true);

		$styleWriter->write();
		$this->xmlWriter->writeAttribute('xml:space', 'preserve');
		if (!$this->withoutP) {
			$this->xmlWriter->endElement(); // w:p
		}
	}
}
