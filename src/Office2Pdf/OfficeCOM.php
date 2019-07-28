<?php
/**
 * Class OfficeCOM | Office/OfficeCOM.php
 *
 * @source    OfficeDOM.php
 * @package   Office
 * @author    AlanAlbert <alan1766447919@gmail.com>
 * @version   v1.0.0	Sunday, July 28th, 2019.
 * @copyright Copyright (c) 2019, AlanAlbert
 * @license   MIT License
 */

namespace Office2Pdf;

/**
 * OfficeCOM
 * 
 * @property resource|null $word
 * @property resource|null $excel
 * @property resource|null $ppt
 * 
 * @method boolean      run($type = 'all')
 * @method boolean      close($type = 'all')
 * @method int|false    getPageNumFromDoc($file)
 * @method int|false    getPageNumFromPpt($file)
 * @method int|false    getPageNumFromExcel($file)
 * @method int|false    getPageNumFromPdf($file)
 * @method boolean      word2Pdf($source_file, $output_file)
 * @method boolean      excel2Pdf($source_file, $output_file)
 * @method boolean      ppt2Pdf($source_file, $output_file)
 * 
 */
class OfficeCOM
{
    /**
     * @var		resource|null	$word   Word文件
     */
    private static $word = null;
    /**
     * @var		resource|null	$ppt    PPT文件
     */
    private static $ppt = null;
    /**
     * @var		resource|null	$excel  Excel文件
     */
    private static $excel = null;

    /**
     * 运行office相关软件
     * @param   string $type     office type: all | word | ppt | excel
     * @return  boolean         true | false
     */
    public function run($type = 'all')
    {
        if (!class_exists('COM')) {
            echo "Office COM interface not exists", PHP_EOL;
            return false;
        }
        try {
            // 打开word程序
            if (($type == 'word' || $type == 'all') && self::$word == null)
                self::$word = new \COM("word.application");

            // 打开ppt程序
            if (($type == 'ppt' || $type == 'all') && self::$ppt == null)
                self::$ppt = new \COM("powerpoint.application");

            // 打开excel程序
            if (($type == 'excel' || $type == 'all') && self::$excel == null)
                self::$excel = new \COM("excel.application");

            return true;
        } catch (Exception $e) {
            echo "open office faild: ", $e->getMessage(), PHP_EOL;
            return false;
        }
        
    }

    /**
     * 关闭office相关软件
     * @param  string $type     office type: all | word | ppt | excel
     * @return boolean          true | false
     */
    public function close($type = 'all')
    {
        try {
            if (($type == 'word' || $type == 'all') && self::$word != null)
                self::$word->Quit();
            if (($type == 'ppt' || $type == 'all') && self::$ppt != null)
                self::$ppt->Quit();
            if (($type == 'excel' || $type == 'all') && self::$excel != null)
                self::$excel->Quit();
            return true;
        } catch (Exception $e) {
            echo "close office faild: ", $e->getMessage(), PHP_EOL;
            return false;
        }
        
    }

    /**
     * 获取word文件的页数
     * @param  string $file     文件路径            file path
     * @return int|false        页数，失败为false    page number or false
     */
    public function getPageNumFromDoc($file) 
    {
        if (!file_exists($file)){
            echo "file(", iconv('gbk', 'utf-8', $file), ") not exists", PHP_EOL;
            return false;
        }
        try {
            self::$word->Visible = 0;
            $document = self::$word->Documents->Open($file, false, false, false, "1", "1", true);
            $document->Repaginate;  // 将文档重新分页（重要），直接获取页数可能存在偏差
            $page_num = $document->BuiltInDocumentProperties(14)->Value;// 获取页数，14表示wdPropertyPages
            $document->Close();

            return $page_num;
        } catch (\Exception $e) {
            // office的错误信息编码为GBK，需要转化为utf8
            echo iconv('GBK', 'UTF-8', $e->getMessage()), PHP_EOL;
            return false;
        }
    }

    /**
     * 获取ppt的页数 
     * @param  string $file     文件路径            file path
     * @return int|false        页数，失败为false    page number or false
     */
    public function getPageNumFromPpt($file) 
    {
        if (!file_exists($file)){
            echo "file(", iconv('gbk', 'utf-8', $file), ") not exists";
            return false;
        }
        try {
            $presentation = self::$ppt
                ->Presentations
                ->Open($file, false, false, false);
            $page_num = $presentation->Slides->Count;
            $presentation->Close();
            return $page_num;
        } catch (\Exception $e) {
            echo iconv('GBK', 'utf-8', $e->getMessage()), PHP_EOL;
            return false;
        }
    }

    /**
     * 获取Excel页数（转pdf）
     * @param  string $file     文件路径            file path
     * @return int|false        页数，失败返回false  page number or false
     */
    public function getPageNumFromExcel($file) 
    {
        // 注释部分为获取的excel文件中表sheet的数量
        // try {
        //     if(!file_exists($file)){
        //         return 0;
        //     }
        //     $excel = new \COM("excel.application") or die("Unable to instantiate excel");
        //     $workbook = $excel
        //         ->Workbooks
        //         ->Open($file, null, false, null, "1", "1", true);
        //     $page_num = $workbook->Sheets->Count;
        //     $workbook->Close();
        //     $excel->Quit();
        //     return $page_num;
        // } catch (\Exception $e) {
        //     echo iconv('gb2312', 'utf-8', $e->getMessage());
        //     if (method_exists($excel, "Quit")){
        //         $excel->Quit();
        //     }
        //     return -1;
        // }
        if (!file_exists($file)){
            echo "file(", iconv('gbk', 'utf-8', $file), ") not exists";
            return false;
        }
        try {
            $pdf_file = $file . '.pdf';
            if ($this->excel2Pdf($file, $pdf_file)) {
                $page_num = $this->getPageNumFromPdf($pdf_file);
                if (file_exists($pdf_file))
                    unlink($pdf_file);
                return $page_num;
            }
            return false;
        } catch (Exception $e) {
            echo iconv('GBK', 'utf-8', $e->getMessage()), PHP_EOL;
            return false;
        }
    }

    /**
     * 获取PDF页数
     * Get page number of PDF
     * @param  string $file     文件路径            file path
     * @return int|false        页数，失败为false    page number or false
     */
    public function getPageNumFromPdf($file){
        if (!file_exists($file)){
            echo "file(", iconv('gbk', 'utf-8', $file), ") not exists", PHP_EOL;
            return false;
        }
        if (!$fp = @fopen($file,"r")) {
            echo "open file(", $file, ") faild", PHP_EOL;
            return false;
        }
        $page = 0;
        while(!feof($fp)) {
            $line = fgets($fp,255);
            if (preg_match('/\/Count [0-9]+/', $line, $matches)){
                preg_match('/[0-9]+/',$matches[0], $matches2);
                if ($page<$matches2[0]) $page=$matches2[0];
            }
        }
        fclose($fp);
        return $page;
    }

    /**
     * Excel转PDF
     * Excel to PDF
     * @param  string $source_file  源文件      source file path
     * @param  string $output_file  目标文件    object file path
     * @return boolean              结果       true | false
     */
    public function excel2Pdf($source_file, $output_file) {
        if (!file_exists($source_file)){
            echo "file(", iconv('gbk', 'utf-8', $source_file), ") not exists", PHP_EOL;
            return false;
        }
        try {
            $workbook = self::$excel->Workbooks->Open($source_file, null, false, null, "1", "1", true);
            $workbook->ExportAsFixedFormat(0, $output_file);
            $workbook->Close();
            return true;
        } catch (\Exception $e) {
            echo iconv('GBK', 'utf-8', $e->getMessage()), PHP_EOL;
            return false;
        }
    }

    /**
     * ppt转pdf
     * @param  string $source_file  源文件路径      source file path
     * @param  string $output_file  目标文件路径    object file path
     * @return boolean              结果           true | false
     */
    public function ppt2Pdf($source_file, $output_file) {
        if (!file_exists($source_file)){
            echo "file(", iconv('gbk', 'utf-8', $source_file), ") not exists", PHP_EOL;
            return false;
        }
        try {
            $presentation = self::$ppt->Presentations->Open($source_file, false, false, false);
            $presentation->SaveAs($output_file,32,1);
            $presentation->Close();
            return true;
        } catch (\Exception $e) {
            echo iconv('GBK', 'utf-8', $e->getMessage()), PHP_EOL;
            return false;
        }
    }

    /**
     * word转pdf
     * @param  string $source_file  源文件路径          source file path
     * @param  string $output_file  目标文件路径        object file path
     * @return boolean              结果               true | false
     */
    public function word2Pdf($source_file, $output_file) {
        if (!file_exists($source_file)){
            echo "file(", iconv('gbk', 'utf-8', $source_file), ") not exists", PHP_EOL;
            return false;
        }
        try {
            self::$word->Visible = 0;
            $document = self::$word->Documents->Open($source_file, false, false, false, "1", "1", true);

            $document->final = false;
            $document->Saved = true;
            $document->ExportAsFixedFormat(
                $output_file,
                17,                         // wdExportFormatPDF
                false,                      // open file after export
                0,                          // wdExportOptimizeForPrint
                3,                          // wdExportFromTo
                1,                          // begin page
                5000,                       // end page
                7,                          // wdExportDocumentWithMarkup
                true,                       // IncludeDocProps
                true,                       // KeepIRM
                1                           // WdExportCreateBookmarks
            );
            $document->Close();
            return true;
        } catch (\Exception $e) {
            echo iconv('GBK', 'utf-8', $e->getMessage()), PHP_EOL;
            return false;
        }
    }
}
