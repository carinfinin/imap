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

}else {
   $ware_house = new WareHouse;
   $path = $_SERVER['DOCUMENT_ROOT'].'/imap/xlsx/';
   $res = $ware_house->getFiles();
   $result = $ware_house->updateProducts($res);
//    print_r($res);
}

echo "Список изменённых товаров <br>";
foreach ($result as $k => $val) {
   $name = $val['NAME'];
   $date = $val['DATE'];
   $count = $val['COUNT'];
   echo "id: $k  name: $name <br> date: $date  <br>count: $count <br><br><br>";
}






