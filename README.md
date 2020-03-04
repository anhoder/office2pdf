
Office转PDF

 > 该库基于Office(或WPS) COM组件，因此只支持Windows系统，其他系统可以考虑使用其他库，例如Open Office、PHPOffice。

## 安装

 * Microsoft Office或WPS
 * PHP配置开启`com.allow_dcom = true`
 * PHP配置`extension=php_com_dotnet.dll`（PHP 5.4.5之前的版本自带，不需要开启）


## 使用

1. `composer require alanalbert/office2pdf`
2. 包含composer的自动加载文件`autoload.php`
3. 编写具体业务逻辑，如下（注意：Windows使用的中文编码为GBK，因此需要使用`iconv`函数进行编码转换）：


```php
require '../vendor/autoload.php';

$source_path_utf8 = "E:/测试.doc";                           // Word文件路径
$source_path = iconv('UTF-8', 'GBK', $source_path_utf8);	// 转为GBK，防止中文乱码而找不到文件
$output_path = $source_path . '.pdf';                       // PDF目标文件路径

$office = new Alan\Office2Pdf\OfficeCOM();
$office->run('word');
var_dump($office->word2Pdf($source_path, $output_path));
$office->close('word');
```

## 方法

```php
/*
 * @method boolean      run($type = 'all')                      运行应用, 参数: all|word|excel|ppt
 * @method boolean      close($type = 'all')                    关闭应用, 参数: all|word|excel|ppt
 * @method int|false    getPageNumFromDoc($file)                获取Word文档页数
 * @method int|false    getPageNumFromPpt($file)                获取PPT文档页数
 * @method int|false    getPageNumFromExcel($file)              获取Excel文档页数
 * @method int|false    getPageNumFromPdf($file)                获取PDF文件页数
 * @method boolean      word2Pdf($source_file, $output_file)    Word转PDF
 * @method boolean      excel2Pdf($source_file, $output_file)   Excel转PDF
 * @method boolean      ppt2Pdf($source_file, $output_file)     PPT转PDF
 */
```

## 更多

如果需要使用更多功能，请参考[微软office开发文档](https://msdn.microsoft.com/vba/office-vba-reference)
（适用于WPS）。
