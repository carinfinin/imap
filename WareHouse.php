<?
use Bitrix\Main\Type\DateTime;

class WareHouse {
    private $columnDate = 21;
    private $columnName = 0;
    private $formats = ['xlsx', 'xls', 'XLSX', 'XLS'];

    public function getProducts() {
        $Products = [];
        $el = new CIBlockElement;
        $ob = $el->GetList(['ID' => 'asc'], [' ' => 21, 'QUANTITY' => 0], false, false,
            ['ID', 'NAME', 'IBLOCK_ID', 'IBLOCK_SECTION_ID', 'DETAIL_PAGE_URL', 'PROPERTY_DATE_WAREHOUSE', 'QUANTITY']);
        while ($res = $ob->GetNext(true, false)) {
            $Products[$res['NAME']] = $res['ID'];
        }
        return $Products;
    }
    public function setProducts($id, $date) {
        CIBlockElement::SetPropertyValuesEx($id, false, array('DATE_WAREHOUSE' => $date));
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
                    $dateProducts = $sheet->getCellByColumnAndRow($this->columnDate, $i)->getValue(); //date
                    if($products[$nameProducts]) {
                        $this->setProducts($products[$nameProducts], $this->getDate($date, $dateProducts));
                        $productsSet[$products[$nameProducts]] = $nameProducts;
                    }
                }
            }
            unlink($path);
        }
        return $productsSet;
    }
    private function getDate($date, $dateProducts) {
        $objDateTime = DateTime::createFromPhp(new \DateTime($date));
        return $objDateTime->add("$dateProducts day")->format("d.m.Y");
    }
}