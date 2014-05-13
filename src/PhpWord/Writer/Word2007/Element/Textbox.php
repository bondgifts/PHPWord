<?php
/**
 * PHPWord
 *
 * @link        https://github.com/PHPOffice/PHPWord
 * @copyright   2014 PHPWord
 * @license     http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt LGPL
 */

namespace PhpOffice\PhpWord\Writer\Word2007\Element;

// use Rhumsaa\Uuid\Uuid; //move this out  TODO

/**
 * TextRun element writer
 *
 * @since 0.10.0
 */
class Textbox extends Element
{
    protected $settings;

    /**
     * Write textrun element
     */
    public function write()
    {
        $settings = $this->element->getSettings();
        $this->settings = $settings;

        $this->xmlWriter->startElement('mc:AlternateContent');
        $this->xmlWriter->startElement('mc:Choice');
        $this->xmlWriter->writeAttribute('Requires', 'wps');
        $this->xmlWriter->startElement('w:drawing');

        $this->xmlWriter->startElement('wp:anchor');
        $this->xmlWriter->writeAttributes(array(
            'distT' => '114300', //distance from text on top edge (emu (default is .125 inches))
            'distB' => '114300', //distance from text on bottom edge
            'distL' => '114300', //distance from text on left edge
            'distR' => '114300', //distance from text on right edge
            'simplePos' => '0', //page positioning - starting coordinates - if this is 1 then the starting coordinates come from wp:simplePos
            'relativeHeight' => '251659264', //z-index, not sure why this is the default
            'behindDoc' => '0', //display behind document text (boolean via 0 or 1)
            'locked' => '0',
            'layoutInCell' => '1',
            'allowOverlap' => '1' //allow textboxes to overlap
        ));

        $this->writePositioning();
        $this->writeExtentSize();
        $this->writeEffectExtent();

        //Wrap text around virtual rectangle bounding the textbox
        $this->xmlWriter->startElement('wp:wrapSquare');
        $this->xmlWriter->writeAttribute('wrapText', 'bothSides');
        $this->xmlWriter->endElement(); // wp:wrapSquare

        $this->writeDocInfo();

        $this->xmlWriter->startElement('wp:cNvGraphicFramePr');
        $this->xmlWriter->endElement(); // wp:cNvGraphicFramePr

        //Graphics
        $this->xmlWriter->startElement('a:graphic');
        $this->xmlWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');

            $this->xmlWriter->startElement('a:graphicData');
            $this->xmlWriter->writeAttribute('uri', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape');

                $this->xmlWriter->startElement('wps:wsp');

                    $this->xmlWriter->startElement('wps:cNvSpPr');
                    $this->xmlWriter->writeAttribute('txBox', '1');
                    $this->xmlWriter->endElement(); // wps:cNvSpPr

                        $this->xmlWriter->startElement('wps:spPr');

                            $this->xmlWriter->startElement('a:xfrm');

                                $this->xmlWriter->startElement('a:off');
                                $this->xmlWriter->writeAttributes(array(
                                    'x' => '0',
                                    'y' => '0'
                                ));
                                $this->xmlWriter->endElement(); // a:off

                                $this->xmlWriter->startElement('a:ext');
                                $this->xmlWriter->writeAttributes(array(
                                    'cx' => $settings->getWidth(),
                                    'cy' => $settings->getHeight()
                                ));
                                $this->xmlWriter->endElement(); // a:ext

                            $this->xmlWriter->endElement(); // a:xfrm

                        $this->xmlWriter->startElement('a:prstGeom');
                        $this->xmlWriter->writeAttributes(array(
                            'prst' => 'rect'
                        ));
                            $this->xmlWriter->startElement('a:avLst');
                            $this->xmlWriter->endElement(); // a:avLst
                        $this->xmlWriter->endElement(); // a:prstGeom

                            $this->xmlWriter->startElement('a:extLst');
                                $this->xmlWriter->startElement('a:ext');
                                $this->xmlWriter->writeAttributes(array(
                                    'uri' => '{C572A759-6A51-4108-AA02-DFA0A04FC94B}'
                                ));

                                    $this->xmlWriter->startElement('ma14:wrappingTextBoxFlag');
                                    $this->xmlWriter->writeAttributes(array(
                                        'xmlns:ma14' => 'http://schemas.microsoft.com/office/mac/drawingml/2011/main'
                                    ));
                                    $this->xmlWriter->endElement(); // ma14:wrappingTextBoxFlag

                                $this->xmlWriter->endElement(); // a:ext
                            $this->xmlWriter->endElement(); // a:extLst

                        $this->xmlWriter->endElement(); // wps:spPr
        
                        $this->xmlWriter->startElement('wps:style');

                            $this->xmlWriter->startElement('a:lnRef');
                            $this->xmlWriter->writeAttribute('idx', '0');
                            $this->xmlWriter->startElement('a:schemeClr');
                            $this->xmlWriter->writeAttribute('val', 'accent1');
                            $this->xmlWriter->endElement(); // a:schemeClr
                            $this->xmlWriter->endElement(); // a:lnRef

                            $this->xmlWriter->startElement('a:fillRef');
                            $this->xmlWriter->writeAttribute('idx', '0');
                            $this->xmlWriter->startElement('a:schemeClr');
                            $this->xmlWriter->writeAttribute('val', 'accent1');
                            $this->xmlWriter->endElement(); // a:schemeClr
                            $this->xmlWriter->endElement(); // a:fillRef

                            $this->xmlWriter->startElement('a:effectRef');
                            $this->xmlWriter->writeAttribute('idx', '0');
                            $this->xmlWriter->startElement('a:schemeClr');
                            $this->xmlWriter->writeAttribute('val', 'accent1');
                            $this->xmlWriter->endElement(); // a:schemeClr
                            $this->xmlWriter->endElement(); // a:effectRef

                            $this->xmlWriter->startElement('a:fontRef');
                            $this->xmlWriter->writeAttribute('idx', 'minor');
                            $this->xmlWriter->startElement('a:schemeClr');
                            $this->xmlWriter->writeAttribute('val', 'dk1');
                            $this->xmlWriter->endElement(); // a:schemeClr
                            $this->xmlWriter->endElement(); // a:fontRef

                        $this->xmlWriter->endElement(); // wps:style

                        //Text in textbox
                        $this->xmlWriter->startElement('wps:txbx');
                            $this->xmlWriter->startElement('w:txbxContent');
                            $this->parentWriter->writeContainerElements($this->xmlWriter, $this->element);
                            $this->xmlWriter->endElement(); // w:txbxContent
                        $this->xmlWriter->endElement(); // wps:txbx

                        $this->writeBodyParagraph();

        $this->xmlWriter->endElement(); // wps:wsp
        $this->xmlWriter->endElement(); // a:graphicData
        $this->xmlWriter->endElement(); // a:graphic


        $this->xmlWriter->endElement(); // wp:anchor
        $this->xmlWriter->endElement(); // w:drawing
        $this->xmlWriter->endElement(); // mc:Choice
        $this->xmlWriter->endElement(); // mc:AlternateContent
    }

    private function writePositioning() 
    {
        //Page positioning, must be enabled from wp:anchor
        //- will override offsets below (wp:positionH, wp:positionV)
        $this->xmlWriter->startElement('wp:simplePos');
        $this->xmlWriter->writeAttributes(array(
            'x' => '0',
            'y' => '0'
        ));
        $this->xmlWriter->endElement(); // wp:simplePos

        //Offsets
        $this->xmlWriter->startElement('wp:positionH');
        $this->xmlWriter->writeAttribute('relativeFrom', 'column'); //specifies what to calculate the offset relative to
        $this->xmlWriter->startElement('wp:posOffset');
        $this->xmlWriter->writeRaw($this->settings->getOffsetX());
        $this->xmlWriter->endElement(); // wp:posOffset
        $this->xmlWriter->endElement(); // wp:positionH

        $this->xmlWriter->startElement('wp:positionV');
        $this->xmlWriter->writeAttribute('relativeFrom', 'paragraph');
        $this->xmlWriter->startElement('wp:posOffset');
        $this->xmlWriter->writeRaw($this->settings->getOffsetY());
        $this->xmlWriter->endElement(); // wp:posOffset
        $this->xmlWriter->endElement(); // wp:positionV
    }

    private function writeExtentSize() 
    {
        //Canvas Size - Specifies final height and width
        $this->xmlWriter->startElement('wp:extent');
        $this->xmlWriter->writeAttributes(array(
            'cx' => $this->settings->getWidth(),
            'cy' => $this->settings->getHeight()
        ));
        $this->xmlWriter->endElement(); // wp:extent
    }

    private function writeEffectExtent() 
    {
        //Additional shape effects - none
        $this->xmlWriter->startElement('wp:effectExtent');
        $this->xmlWriter->writeAttributes(array(
            'l' => '0',
            't' => '0',
            'r' => '0',
            'b' => '0',
        ));
        $this->xmlWriter->endElement(); // wp:effectExtent
    }

    private function writeDocInfo()
    {
        //Unique Identifier & name of object
        //Duplicates do not break the xml apparently
        //- not sure we will need worry about this
        $this->xmlWriter->startElement('wp:docPr');
        $this->xmlWriter->writeAttributes(array(
            'id' => '1',
            'name' => 'Text Box '.$this->element->getElementId(),
        ));
        $this->xmlWriter->endElement(); // wp:docPr
    }

    private function writeBodyParagraph()
    {
        //Body Paragraph
        $this->xmlWriter->startElement('wps:bodyPr');
        $this->xmlWriter->writeAttributes(array(
            'rot' => '0',
            'spcFirstLastPara' => '0',
            'vertOverflow' => 'overflow',
            'horzOverflow' => 'overflow',
            'vert' => 'horz',
            'wrap' => 'square',
            'lIns' => '91440',
            'tIns' => '45720',
            'lIns' => '91440',
            'bIns' => '45720',
            'numCol' => '1',
            'spcCol' => '0',
            'rtlCol' => '0',
            'fromWordArt' => '0',
            'anchor' => 't',
            'anchorCtr' => '0',
            'forceAA' => '0',
            'compatLnSpc' => '1'
        ));
        $this->xmlWriter->startElement('a:prstTxWarp');
        $this->xmlWriter->writeAttributes(array(
            'prst' => 'textNoShape'
        ));
        $this->xmlWriter->startElement('a:avLst');
        $this->xmlWriter->endElement(); // a:avLst
        $this->xmlWriter->endElement(); // a:prstTxWarp
        $this->xmlWriter->startElement('a:noAutofit');
        $this->xmlWriter->endElement(); // a:noAutofit
        $this->xmlWriter->endElement(); // wps:bodyPr
    }

   
    // $this->xmlWriter->writeRaw('<mc:Fallback>
                //         <w:pict>
                //             <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m0,0l0,21600,21600,21600,21600,0xe">
                //                 <v:stroke joinstyle="miter" />
                //                 <v:path gradientshapeok="t" o:connecttype="rect" /></v:shapetype>
                //             <v:shape id="Text Box 1" o:spid="_x0000_s1026" type="#_x0000_t202" style="position:absolute;margin-left:108pt;margin-top:0;width:234pt;height:162pt;z-index:251659264;visibility:visible;mso-wrap-style:square;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:absolute;mso-position-horizontal-relative:text;mso-position-vertical:absolute;mso-position-vertical-relative:text;v-text-anchor:top" o:gfxdata="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&#xA;UEsFBgAAAAAEAAQA8wAAAC0GAAAAAA==&#xA;" filled="f" stroked="f">
                //                 <v:textbox>
                //                     <w:txbxContent>
                //                         <w:p w:rsidR="000E1FF9" w:rsidRDefault="000E1FF9">
                //                             <w:r>
                //                                 <w:t>Testing Bond Positioning</w:t>
                //                             </w:r>
                //                             <w:bookmarkStart w:id="1" w:name="_GoBack" />
                //                             <w:bookmarkEnd w:id="1" /></w:p>
                //                     </w:txbxContent>
                //                 </v:textbox>
                //                 <w10:wrap type="square" />
                //             </v:shape>
                //         </w:pict>
                //     </mc:Fallback>');
                //     
}
