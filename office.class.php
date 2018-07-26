<?php

/**
 * @File:	office.class.php
 * @Author: Alan_Albert
 * @Email:  1766447919@qq.com
 * @Date:	2018-07-25 12:43:04
 * @Last Modified by:   Alan_Albert
 * @Last Modified time: 2018-07-25 16:12:56
 */

// office开发文档URL：https://msdn.microsoft.com/vba/office-vba-reference

class Office
{
    private static $word = null;
    private static $ppt = null;
    private static $excel = null;

    public function __construct() {}
    public function __destruct() {}

    /**
     * 运行office相关软件
     * @param  string $type office软件: all, word, ppt, excel
     * @return  boolean      结果: true, false
     */
    public function run($type = 'all')
    {
        if (!class_exists('COM')) {
            echo "Office COM interface not exists";
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
            echo "open office faild: ", $e->getMessage();
            return false;
        }
        
    }

    /**
     * 关闭office相关软件
     * @param  string $type office软件: all, word, ppt, excel
     * @return boolean       结果: true, false
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
            echo "close office faild: ", $e->getMessage();
            return false;
        }
        
    }

    /**
     * 获取word文件的页数
     * @param  string $file 文件路径
     * @return int|false       页数，失败为false
     */
    public function getPageNumFromDoc($file) 
    {
        if (!file_exists($file)){
            echo "file(", iconv('gbk', 'utf-8', $file), ") not exists";
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
            echo iconv('GBK', 'UTF-8', $e->getMessage());
            return false;
        }
    }

    /**
     * 获取ppt的页数 
     * @param  string $file 文件路径
     * @return int|false       页数，失败为false
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
            echo iconv('GBK', 'utf-8', $e->getMessage());
            return false;
        }
    }

    /**
     * 获取Excel页数（转pdf）
     * @param  string $file 文件路径
     * @return int|false       页数，失败返回false
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
            echo $e->getMessage();
            return false;
        }
    }

    /**
     * 获取PDF页数
     * @param  string $file 文件路径
     * @return int|false       页数，失败为false
     */
    public function getPageNumFromPdf($file){
        if (!file_exists($file)){
            echo "file(", iconv('gbk', 'utf-8', $file), ") not exists";
            return false;
        }
        if (!$fp = @fopen($file,"r")) {
            echo "open file(", $file, ") faild";
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
     * @param  string $source_file 源文件
     * @param  string $output_file 目标文件
     * @return boolean              结果
     */
    public function excel2Pdf($source_file, $output_file) {
        if (!file_exists($source_file)){
            echo "file(", iconv('gbk', 'utf-8', $source_file), ") not exists";
            return false;
        }
        try {
            $workbook = self::$excel->Workbooks->Open($source_file, null, false, null, "1", "1", true);
            $workbook->ExportAsFixedFormat(0, $output_file);
            $workbook->Close();
            return true;
        } catch (\Exception $e) {
            echo iconv('GBK', 'utf-8', $e->getMessage());
            return false;
        }
    }

    /**
     * ppt转pdf
     * @param  string $source_file 源文件路径
     * @param  string $output_file 目标文件路径
     * @return boolean              结果
     */
    function ppt2Pdf($source_file, $output_file) {
        if (!file_exists($source_file)){
            echo "file(", iconv('gbk', 'utf-8', $source_file), ") not exists";
            return false;
        }
        try {
            $presentation = self::$ppt->Presentations->Open($source_file, false, false, false);
            $presentation->SaveAs($output_file,32,1);
            $presentation->Close();
            return true;
        } catch (\Exception $e) {
            echo iconv('GBK', 'utf-8', $e->getMessage());
            return false;
        }
    }

    /**
     * word转pdf
     * @param  string $source_file 源文件路径
     * @param  string $output_file 目标文件路径
     * @return boolean              结果
     */
    function word2Pdf($source_file, $output_file) {
        if (!file_exists($source_file)){
            echo "file(", iconv('gbk', 'utf-8', $source_file), ") not exists";
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
            echo iconv('GBK', 'utf-8', $e->getMessage());
            return false;
        }
    }
}
