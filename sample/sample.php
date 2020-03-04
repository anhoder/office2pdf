<?php

/**
 * @File:	sample.php
 * @Author: Alan_Albert
 * @Email:  1766447919@qq.com
 * @Date:	2018-07-25 12:53:40
 * @Last Modified by:   Alan_Albert
 */

require '../vendor/autoload.php';

$source_path_utf8 = "E:/测试.doc";                           // Word文件路径
$source_path = iconv('UTF-8', 'GBK', $source_path_utf8);	// 转为GBK，防止中文乱码而找不到文件
$output_path = $source_path . '.pdf';                       // PDF目标文件路径

$office = new Alan\Office2Pdf\OfficeCOM();
$office->run('word');

var_dump($office->word2Pdf($source_path, $output_path));

$office->close('word');