<?php

/**
 * @File:	test.php
 * @Author: Alan_Albert
 * @Email:  1766447919@qq.com
 * @Date:	2018-07-25 12:53:40
 * @Last Modified by:   Alan_Albert
 * @Last Modified time: 2018-07-25 16:13:34
 */

require 'office.class.php';

$source_path_utf8 = "E:/测试.doc";
$source_path = iconv('UTF-8', 'GBK', $source_path_utf8);	// 转为GBK，防止中文乱码而找不到文件

$output_path = $source_path . '.pdf';

$office = new Office();
$office->run('word');

var_dump($office->word2Pdf($source_path, $output_path));

$office->close('word');