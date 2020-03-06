<?php

/**
 * Class WordReader
 *
 * this class can read docx file.
 * This file has ability to fetch images embedded in this word file and save it in specified folder.
 */
class WordReader {

    private $content;
    private $picTag = 'pic:cNvPr';
    private $html = true; // true to get html, false to get string
    private $imagePath = '/tmp/'; // image path to write images embedded in docx file
    private $imagePathUrl = 'http://domain.com/public'; // image path url to get url for image
    private $imageFiles = []; // hold image files embedded in docx file

    public function __construct($returnHtml = true)
    {
        $this->html = $returnHtml;
    }

    /**
     * read file content
     *
     * @param $filename
     * @return array|bool|string
     */
    public function read ($filename)
    {
        // check file exists
        if (!$filename || !file_exists($filename)) return false; // return false if file doesn't exists
        $zip = zip_open($filename);
        if (!$zip || is_numeric($zip)) return false;
        $this->content = '';
        while ($zip_entry = zip_read($zip)) {
            if (zip_entry_open($zip, $zip_entry) == FALSE) continue;

            $zip_name_entry = zip_entry_name($zip_entry);
            $zip_file_content = zip_entry_read($zip_entry, zip_entry_filesize($zip_entry));

            // word/media
            if (strpos($zip_name_entry, 'word/media') !== false) {
                $name = explode('/', $zip_name_entry);
                $name = end($name);
                $this->imageFiles[] = strstr($name, '.', true);
                file_put_contents($this->imagePath . '/' . $name, $zip_file_content);
                unset($name);
            }
            if ($zip_name_entry != "word/document.xml") continue;
            $this->content .= $zip_file_content;
            unset($zip_file_content);
            zip_entry_close($zip_entry);
        }
        zip_close($zip);
        $this->content = str_replace('</w:r></w:p></w:tc><w:tc>', " ", $this->content);
        // $content = str_replace('</w:r></w:p>', "\r\n", $content);
        $this->content = explode('</w:p>', $this->content);

        // $striped_content = strip_tags($content);
        // $striped_content = $content;
        foreach($this->content as $key => $val) {
            $innerContent = explode('</w:r>', $val);
            if (!empty($innerContent)) {
                foreach($innerContent as $iKey => $iVal) {
                    $innerContent[$iKey] = $this->getContent($iVal);
                }
                $this->content[$key] = implode(' ', $innerContent);
            } else {
                $this->content[$key] = $this->getContent($val);
            }
        }

        return $this->html ? '<p>' . implode('</p><p>', $this->content) . '</p>' : $this->content;
    }

    /**
     * read content of tag
     *
     * @param $content
     * @return string
     */
    private function getContent($content)
    {
        if (strpos($content, $this->picTag) !== false) {
            // pic found in this line
            $content = $this->getImage($content);
        } else {
            $content = strip_tags($content);
        }
        return trim($content);
    }

    /**
     * get image from content and parse it properly
     *
     * @param $content
     * @return string
     */
    protected function getImage ($content)
    {
        $pic = '';
        if (strpos($content, $this->picTag) !== false) {
            // pic found in this line
            $start_deli = '<' . $this->picTag;
            $end_deli = '</' . $this->picTag . '>';
            $pic = $this->getStringBetweenDeli($content, $start_deli, $end_deli);
            if (!empty($pic)) {
                $pic = '<pic' . $pic . '</pic>';
                $arr = $this->xmlToArray($pic);
                $image = [
                    'src' => $arr['@attributes']['descr'],
                    'name' => $arr['@attributes']['name']
                ];
                $pic = $this->getImageHtml($image);
            }
        }
        return trim($pic);
    }

    /**
     * get image src and name in image tag or in string
     *
     * @param $image
     * @return string
     */
    private function getImageHtml($image)
    {
        if ($this->html && !empty($image['src'])) {
            return sprintf('<image src="%s" alt="%s" />', $image['src'], $image['name']);
        } else if ($this->html) {
            return sprintf('<image src="%s" alt="%s" />', $this->imagePathUrl . '/' . $image['src'], $image['name']);
        } else {
            return $image['name'];
        }
    }

    /**
     * Get string between delimiters
     *
     * @param $string
     * @param $start_deli
     * @param $end_deli
     * @return bool|string
     */
    private function getStringBetweenDeli($string, $start_deli, $end_deli)
    {
        $ini = strpos($string, $start_deli);
        if ($ini == 0) return '';
        $ini += strlen($start_deli);
        $len = strpos($string, $end_deli, $ini) - $ini;
        $content = substr($string, $ini, $len);
        return $content;
    }

    /**
     * convert xml to array format
     *
     * @param $xml
     * @return array
     */
    private function xmlToArray($xml)
    {
        return (array) simplexml_load_string($xml);
    }
}
