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
use Rhumsaa\Uuid\Uuid;

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

    protected $uri;

    /**
     * Create new instance
     *
     * @param int $sectionCount
     * @param array $settings
     */
    public function __construct($settings = null)
    {
        $this->container = 'textbox';
        $this->setDocPart($this->container);
        $this->settings = new TextboxSettings();
        $this->setSettings($settings);
        $this->setUri();
        
    }

    private function setUri()
    {
        $this->uri = '{'.Uuid::uuid4().'}';
    }

    public function getUri()
    {
        return $this->uri;
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
