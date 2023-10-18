<?php
require 'vendor/autoload.php';


use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

use function PHPSTORM_META\type;

class quanlydoanhthu extends Controller
{
    public function start()
    {

        $serverName = "localhost";
        $userName = "root";
        $password = "";
        $dataBaseName = "quan_ly_khach_san";

        $connection = new mysqli($serverName, $userName, $password, $dataBaseName);
        if ($connection->connect_error) {
            die("Connection failed: " . $connection->connect_error);
        }

        $HoaDonModel = $this->model("HoaDonModel");

        $dataHoadon = $HoaDonModel->getAllHoaDon();

        if (isset($_POST['locdoanhthu'])) {
            $loc = $_POST['lochoadon'];
            $dataLoc = [];
            $dataExcel = null;

            if ($loc == "theongay") {
                $currTime = date('d');
                $dataLoc = $HoaDonModel->getTheoNgay($currTime);
                $dataExcel = $HoaDonModel->getTheoNgay($currTime);
            } elseif ($loc == "theothang") {
                $currTime = date('m');
                $dataLoc = $HoaDonModel->getTheoThang($currTime);
                $dataExcel = $HoaDonModel->getTheoThang($currTime);
            } elseif ($loc == "theonam") {
                $currTime = date('Y');
                $dataLoc = $HoaDonModel->getTheoNam($currTime);
                $dataExcel = $HoaDonModel->getTheoNam($currTime);
            }

            // print_r(gettype($dataExcel));
        }




        if (isset($_POST['xuatexcel'])) {
            $spreadsheet = new Spreadsheet();
            $sheet = $spreadsheet->getActiveSheet();
            $sqltk = "SELECT sohopdong, roomID, name, ngaydat, tongchiphi FROM `hoadon`";
            $data_export = mysqli_query($connection, $sqltk);
            //định dạng cột tiêu đề
            $sheet->getColumnDimension('A')->setAutoSize(true);
            $sheet->getColumnDimension('B')->setAutoSize(true);
            $sheet->getColumnDimension('C')->setAutoSize(true);
            $sheet->getColumnDimension('D')->setAutoSize(true);
            $sheet->getColumnDimension('E')->setAutoSize(true);
            // căn lề cácc tiêu đề trong các ô
            $sheet->getStyle('A1:E1')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
            // Tạo tiêu đề
            $sheet
                ->setCellValue('A1', 'Số hợp đồng')
                ->setCellValue('B1', 'Số phòng')
                ->setCellValue('C1', 'Tên khách hàng')
                ->setCellValue('D1', 'Ngày đặt')
                ->setCellValue('E1', 'Tổng chi phí');
            // Ghi dữ liệu
            $rowCount = 2;
            foreach ($data_export as $key => $value) {
                $sheet->setCellValue('A' . $rowCount, $value['sohopdong']);
                $sheet->setCellValue('B' . $rowCount, $value['roomID']);
                $sheet->setCellValue('C' . $rowCount, $value['name']);
                $sheet->setCellValue('D' . $rowCount, $value['ngaydat']);
                $sheet->setCellValue('E' . $rowCount, $value['tongchiphi']);
                //căn lề cho các văn bản trong các ô thuộc mỗi hàng
                $sheet->getStyle('A' . $rowCount . ':E' . $rowCount)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                $rowCount++;
            }
            $writer = new Xlsx($spreadsheet);
            $writer->setOffice2003Compatibility(true);
            $filename = "Hoadon" . time() . ".xlsx";
            $writer->save($filename);
            header("location:" . $filename);
        }

        $this->view("defaultLayout", [
            "container" => "quanlydoanhthu",
            "dataHoadon" => $dataHoadon,
            "dataLoc" => $dataLoc
        ]);
    }

    public function chitiet($sohopdong)
    {
        $ServiceModel = $this->model("ServiceModel");
        $RoomOrdered = $this->model("RoomOrdered");

        $datadichvu = $ServiceModel->layDichVuDaDangKyQuaSHD($sohopdong);
        $dataPhongDaDat = $RoomOrdered->layPhongDaDatQuaSHD($sohopdong);


        $this->view("defaultLayout", [
            "container" => "chitiethopdongkhachhang",
            "datadichvu" => $datadichvu,
            "dataPhongDaDat" => $dataPhongDaDat
        ]);
    }
}
