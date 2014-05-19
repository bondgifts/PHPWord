<?php
/**
 * PHPWord
 *
 * @link        https://github.com/PHPOffice/PHPWord
 * @copyright   2014 PHPWord
 * @license     http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt LGPL
 */

namespace PhpOffice\PhpWord\Style;

/**
 * Section settings
 */
class Textbox extends AbstractStyle
{
    /**
     * Page default constants
     *
     * @const int|float
     */
    const DEFAULT_WIDTH = 914440; // In EMU
    const DEFAULT_HEIGHT = 914440; // In EMU
    const DEFAULT_OFFSET_X = 0; // In EMU
    const DEFAULT_OFFSET_Y = 0; // In EMU
    const DEFAULT_ROTATION = 0;

    private $width = self::DEFAULT_WIDTH;
    private $height = self::DEFAULT_HEIGHT;
    private $offsetX = self::DEFAULT_OFFSET_X;
    private $offsetY = self::DEFAULT_OFFSET_Y;
    private $rotation = self::DEFAULT_ROTATION;

    /**
     * Set Setting Value
     *
     * @param string $key
     * @param string $value
     * @return self
     */
    public function setSettingValue($key, $value)
    {
        return $this->setStyleValue($key, $value);
    }

    /**
     * Get Width
     *
     * @return int|float
     */
    public function getWidth()
    {
        return $this->width;
    }

    /**
     * Set Width
     * @param string $value [description]
     */
    public function setWidth($value = '')
    {
        $this->width = $this->setNumericVal($value, self::DEFAULT_WIDTH);

        return $this;
    }

    /**
     * Get Height
     *
     * @return int|float
     */
    public function getHeight()
    {
        return $this->height;
    }

    public function setHeight($value = '')
    {
        $this->height = $this->setNumericVal($value, self::DEFAULT_HEIGHT);

        return $this;
    }

    public function getOffsetX()
    {
        return $this->offsetX;
    }

    public function setOffsetX($value = '')
    {
        $this->offsetX = $this->setNumericVal($value, self::DEFAULT_OFFSET_X);

        return $this;
    }

    public function getOffsetY()
    {
        return $this->offsetY;
    }

    public function setOffsetY($value = '')
    {
        $this->offsetY = $this->setNumericVal($value, self::DEFAULT_OFFSET_Y);

        return $this;
    }

    /**
     * Get Rotation
     *
     * @return int|float
     */
    public function getRotation()
    {
        return $this->rotation;
    }

    public function setRotation($value = '')
    {
        $this->rotation = $this->setNumericVal($value, self::DEFAULT_ROTATION);

        return $this;
    }

}
