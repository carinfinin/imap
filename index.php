<?
require_once($_SERVER['DOCUMENT_ROOT'] . "/bitrix/modules/main/include/prolog_before.php");
require_once ($_SERVER['DOCUMENT_ROOT'].'/excel/Classes/PHPExcel/IOFactory.php');
require_once ($_SERVER['DOCUMENT_ROOT'].'/imap/SaveAttachment.php');
require_once ($_SERVER['DOCUMENT_ROOT'].'/imap/WareHouse.php');

$env_imap = require_once($_SERVER['DOCUMENT_ROOT'].'/imap/config.php');

$imap = new SaveAttachment($env_imap['email'],$env_imap['password'],"yandex.ru");
$res = $imap->createFile('UNSEEN');

if ($res && $res != [])  {
    $ware_house = new WareHouse;
    $result = $ware_house->updateProducts($res);

    echo "Список изменённых товаров <br>";
    foreach ($result as $k => $val)
        echo "id: $k  name: $val <br>";
}






