<?
use Bitrix\Main\Type\DateTime;

class WareHouse {
    private $columnDate = 34;
    private $columnCount = 33;
    private $columnName = 0;
    private $formats = ['xlsx', 'xls', 'XLSX', 'XLS'];

    public function getProducts() {
        $Products = [];
        $el = new CIBlockElement;
        $ob = $el->GetList(['ID' => 'asc'], ['IBLOCK_ID' => 21, 'QUANTITY' => 0], false, false,
            ['ID', 'NAME', 'IBLOCK_ID', 'IBLOCK_SECTION_ID', 'DETAIL_PAGE_URL', 'PROPERTY_DATE_WAREHOUSE', 'QUANTITY']);
        while ($res = $ob->GetNext(true, false)) {
            $Products[$res['NAME']] = $res['ID'];
        }
        return $Products;
    }
    public function setProducts($id, $date, $count) {
        CIBlockElement::SetPropertyValuesEx($id, false, array('DATE_WAREHOUSE' => $date));

        $StoreID = \Bitrix\Catalog\StoreProductTable::getList(array(
            'filter' => array('=PRODUCT_ID' => $id,'=STORE.ACTIVE' => 'Y'),
            'select' => array('ID'),
        ))->fetch();

        $arFields = Array(
            "PRODUCT_ID" => $id,
            "STORE_ID" => 1,
            "AMOUNT" => $count,
        );

        $ID = CCatalogStoreProduct::Update($StoreID['ID'], $arFields);
    }

    public function getFiles($dir = '/imap/xlsx/') {
        $dir = $_SERVER['DOCUMENT_ROOT']. $dir;
        $date = file_get_contents($_SERVER['DOCUMENT_ROOT'] . '/imap/date_update.txt');

        $list = scandir($dir);
        unset($list[0],$list[1]);

        foreach ($list as $file)
        {
            if(in_array(pathinfo($file)['extension'], $this->formats)) {
                return [$date =>  $_SERVER['DOCUMENT_ROOT'].'/imap/xlsx/' . $file];
            }
        }
    }

    public function updateProducts($arr) {
        $products = $this->getProducts();
        $productsSet = [];
        foreach ($arr as $date => $path) {

            $path_info = pathinfo($path);
            if(!in_array($path_info['extension'], $this->formats)) {
                unlink($path);
                die("Не поддерживаемый формат файла");
            }


            $xls = PHPExcel_IOFactory::load($path);
            $xls->setActiveSheetIndex(0);
            $sheet = $xls->getActiveSheet();

            for ($i = 1; $i <= $sheet->getHighestRow(); $i++) {
                $nColumn = PHPExcel_Cell::columnIndexFromString($sheet->getHighestColumn());
                if($sheet->getCellByColumnAndRow($this->columnDate, $i)->getValue()) {
                    $nameProducts = $sheet->getCellByColumnAndRow($this->columnName, $i)->getValue();  //NAME
                    $dateProducts = ($sheet->getCellByColumnAndRow($this->columnDate, $i)->getValue()) * -1; //date
                    $countProducts = $sheet->getCellByColumnAndRow($this->columnCount, $i)->getValue(); //count
                    if($products[$nameProducts] && $dateProducts != '#NULL!' && $countProducts != '#NULL!') {
                        $this->setProducts($products[$nameProducts], $this->getDate($date, $dateProducts), $countProducts);
                        $productsSet[$products[$nameProducts]]['NAME'] = $nameProducts;
                        $productsSet[$products[$nameProducts]]['DATE'] = $dateProducts;
                        $productsSet[$products[$nameProducts]]['COUNT'] = $countProducts;
                    }
                }
            }
//            unlink($path);
        }
        return $productsSet;
    }
    private function getDate($date, $dateProducts) {
        $objDateTime = DateTime::createFromPhp(new \DateTime($date));
        return $objDateTime->add("$dateProducts day")->format("d.m.Y");
    }
}