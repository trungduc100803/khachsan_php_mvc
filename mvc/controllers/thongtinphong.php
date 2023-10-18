<?php
class thongtinphong extends Controller
{
    public function start()
    {

        $serviceModel = $this->model("ServiceModel");
        $roomModel = $this->model("RoomModel");
        $hoaDonModel = $this->model("HoaDonModel");
        $roomOrdered = $this->model("RoomOrdered");
        $customerModel = $this->model("CustomerModel");


        $dataRooms = $roomModel->getARoomForName($_SESSION['username']);
        $dataService = $serviceModel->getAllService();

        $tenkhachhang = $_SESSION['username'];
        $ngaydat = $_SESSION['ngaydat'];
        $roomID = $_SESSION['roomID'];

        $dataR = $roomOrdered->getSHD($roomID, $tenkhachhang, $ngaydat);
        $sohopdong = mysqli_fetch_assoc($dataR);

        $dataDichvu = $serviceModel->layDichVuDaDatVoiSHD($sohopdong['sohopdong']);

        if (isset($_POST['datdichvu'])) {
            $arrTendichvu = [];
            $arrSoluong = [];
            $arrGia = [];
            $arrThanhtien = [];
            $lengthService = $_POST['lengthService'];

            $tenkhachhang = $_POST['kh'];
            $ngaydat = $_POST['ngayd'];
            $dataR = $roomOrdered->getSHD($roomID, $tenkhachhang, $ngaydat);
            $sohopdong = mysqli_fetch_assoc($dataR);



            for ($i = 0; $i < $lengthService; $i++) {
                array_push($arrTendichvu, $_POST['tendichvu' . $i + 1]);
                array_push($arrSoluong, $_POST['thanhtien' . $i + 1] / $_POST['gia' . $i + 1]);
                array_push($arrGia, $_POST['gia' . $i + 1]);
                array_push($arrThanhtien, $_POST['thanhtien' . $i + 1]);
            }

            for ($j = 0; $j < count($arrSoluong); $j++) {
                if ($arrSoluong[$j] > 0) {
                    $serviceModel->addService($tenkhachhang, $sohopdong['sohopdong'], $arrTendichvu[$j], $arrSoluong[$j], $arrGia[$j], $arrThanhtien[$j]);
                }
            }
            echo '<script>alert("Bạn đã đặt xong các dịch vụ");</script>';
            echo '<script>window.history.back();</script>';
        }

        if (isset($_POST['nhanphong'])) {
            $roomid = $_POST['roomid'];
            $currentTime = date('Y-m-d H:i:s');

            $dataRoom = $roomModel->getARoom($roomid);
            $room = mysqli_fetch_row($dataRoom);

            $hoaDonModel->updateNgayNhan($room[3], $roomid, $currentTime);
            $roomModel->updateTrangthaiphong($roomid, 'Đã nhận');
            echo '<script>window.history.back();</script>';
            // print_r($room);
        }

        if (isset($_POST['btnthanhtoan'])) {

            $tongthanhtoan = $_POST['tongthanhtoan'];
            $roomid = $_POST['roomid'];

            $ngaytraphong = date('Y-m-d H:i:s');

            $dataRoom = $roomModel->getARoom($roomid);
            $room = mysqli_fetch_row($dataRoom);

            $tenkhachhang = $_POST['kh'];
            $ngaydat = $_POST['ngayd'];
            $dataR = $roomOrdered->getSHD($roomID, $tenkhachhang, $ngaydat);
            $sohopdong = mysqli_fetch_assoc($dataR);


            $hoaDonModel->updateNgayTra($sohopdong['sohopdong'], $ngaytraphong, $tongthanhtoan);
            $roomModel->updateTrangthaiphong($roomID, 'Đang trống');
            $roomModel->setPhongtrong($roomID);
            $customerModel->suatrangthaidat($room[3]);
            echo '<script>window.history.back();</script>';
        }
        $this->view("defaultLayout", [
            "container" => "thongtinphong",
            "dataRooms" => $dataRooms,
            "dataService" => $dataService,
            "hoaDonModel" => $hoaDonModel,
            "dataDichvu" => $dataDichvu
        ]);
    }
}
