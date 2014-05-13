<?php
/**
 * PHPWord
 *
 * @link        https://github.com/PHPOffice/PHPWord
 * @copyright   2014 PHPWord
 * @license     http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt LGPL
 */

namespace PhpOffice\PhpWord\Element;

use PhpOffice\PhpWord\Exception\Exception;
use PhpOffice\PhpWord\Style\Textbox as TextboxSettings;

/**
 * Section
 */
class Textbox extends AbstractContainer
{
    /**
     * Section settings
     *
     * @var \PhpOffice\PhpWord\Style\Section
     */
    private $settings;

    /**
     * Create new instance
     *
     * @param int $sectionCount
     * @param array $settings
     */
    public function __construct($settings = null)
    {
        $this->container = 'textbox';
        // $this->setDocPart($this->container, $this->sectionId);
        $this->settings = new TextboxSettings();
        $this->setSettings($settings);
    }

    /**
     * Set section settings
     *
     * @param array $settings
     */
    public function setSettings($settings = null)
    {

        if (!is_null($settings) && is_array($settings)) {
            foreach ($settings as $key => $value) {
                if (is_null($value)) {
                    continue;
                }
                $this->settings->setSettingValue($key, $value);
            }
        }
    }

    /**
     * Get Section Settings
     *
     * @return \PhpOffice\PhpWord\Style\Section
     */
    public function getSettings()
    {
        return $this->settings;
    }


}
