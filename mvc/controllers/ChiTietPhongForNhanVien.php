<?php
class ChiTietPhongForNhanVien extends Controller
{
    public function start($roomID)
    {
        $roomModel = $this->model("RoomModel");
        $customerModel = $this->model("CustomerModel");
        $roomOrdered = $this->model("RoomOrdered");
        $serviceModel = $this->model("ServiceModel");
        $hoaDonModel = $this->model("HoaDonModel");


        $dataService = $serviceModel->getAllService();


        $dataRoom =  $roomModel->getARoom($roomID);
        $room = mysqli_fetch_row($dataRoom);
        $tenkhachhang = $room[3];
        $dataDichvu = null;
        if ($tenkhachhang == "Đang trống") {
        } else {
            $datangaydat = $hoaDonModel->getAHoaDon($tenkhachhang, $roomID);
            $numrow = mysqli_num_rows($datangaydat);
            $datangaydat->data_seek($numrow - 1);
            $ngaydat = $datangaydat->fetch_row();

            $dataR = $roomOrdered->getSHD($roomID, $tenkhachhang, $ngaydat[4]);
            $sohopdong = mysqli_fetch_assoc($dataR);

            $dataDichvu = $serviceModel->layDichVuDaDatVoiSHD($sohopdong['sohopdong']);
        }



        ///xu ly dat phong
        if (isset($_POST['datphong'])) {
            $loaikh = $_POST['loaikhachhang'];

            if ($loaikh == "Khách hàng mới") {
                // lay du lieu tu form dat phong
                $tenkhachhang = $_POST['tenkhachhang'];
                $cccd = $_POST['cccd'];
                $sdt = $_POST['sdt'];
                $ngaysinh = $_POST['ngaysinh'];
                $quoctich = $_POST['quoctich'];
                $diachi = $_POST['diachi'];
                $gioitinh = $_POST['gioitinh'];
                $ngaydat = date('Y-m-d H:i:s');
                $phuongthucthanhtoan = $_POST['phuongthucthanhtoan'];
                $tiencoc = $_POST['tiencoc'];
                $loaiphong =  $roomModel->getLoaiPhong($roomID);
                $gia =  $roomModel->getGiaPhong($roomID);



                //kiem tra xem có tồn tại kh
                $checkCustomer = $customerModel->getACustomer($tenkhachhang);
                $dataCheckCustomer = mysqli_num_rows($checkCustomer);

                $dataC = mysqli_fetch_row($checkCustomer);

                if ($dataCheckCustomer > 0 && $dataC[9] == "Đã đặt") {
                    echo '<script>alert("Khách hàng vẫn đang đặt phòng");</script>';
                    echo '<script>window.history.back();</script>';
                } else if ($dataCheckCustomer > 0 && $dataC[9] == "Đã trả") {
                    $roomOrdered->addRoomOrdered($roomID, $ngaydat, '', $loaiphong, $gia, $tenkhachhang, $cccd, $sdt, $ngaysinh, $quoctich, $diachi, $gioitinh);
                    $roomModel->updateStatusRoom($roomID, 'Đã đặt', $tenkhachhang);
                    $customerModel->suatrangdat($tenkhachhang);

                    $dataR = $roomOrdered->getSHD($roomID, $tenkhachhang, $ngaydat);
                    $shd = mysqli_fetch_row($dataR);

                    $hoaDonModel->addHoaDon($tenkhachhang, $roomID, $ngaydat, '', '', 0, $shd[5], $tiencoc, $phuongthucthanhtoan);
                    echo '<script>window.history.back();</script>';
                    echo '<script>window.history.back();</script>';
                } else {
                    $roomOrdered->addRoomOrdered($roomID, $ngaydat, '', $loaiphong, $gia, $tenkhachhang, $cccd, $sdt, $ngaysinh, $quoctich, $diachi, $gioitinh);
                    $roomModel->updateStatusRoom($roomID, 'Đã đặt', $tenkhachhang);
                    $customerModel->addCustomer($tenkhachhang, $cccd, $sdt, $ngaysinh, $quoctich, $diachi, $gioitinh, 1, 'Đã đặt');

                    $dataR = $roomOrdered->getSHD($roomID, $tenkhachhang, $ngaydat);
                    $shd = mysqli_fetch_row($dataR);

                    $hoaDonModel->addHoaDon($tenkhachhang, $roomID, $ngaydat, '', '', 0, $shd[5], $tiencoc, $phuongthucthanhtoan);
                    echo '<script>window.history.back();</script>';
                    echo '<script>window.history.back();</script>';
                }
            } else {
                $phuongthucthanhtoan = $_POST['phuongthucthanhtoan'];
                $tiencoc = $_POST['tiencoc'];
                $username = $_POST['tenkhachhang'];
                $currentDateTime = date('Y-m-d H:i:s');

                //kiem tra xem có tồn tại kh
                $checkCustomer = $customerModel->getACustomer($username);
                $dataCheckCustomer = mysqli_num_rows($checkCustomer);

                //get room 
                $dataRoom = mysqli_fetch_assoc($roomModel->getARoom($roomID));
                $loaiphong = $dataRoom['loaiphong'];
                $gia = $dataRoom['gia'];

                $dataC = mysqli_fetch_assoc($checkCustomer);
                $cccd = $dataC['cccd'];
                $sdt = $dataC['sdt'];
                $ngaysinh = $dataC['ngaysinh'];
                $quoctich = $dataC['quoctich'];
                $diachi = $dataC['diachi'];
                $gioitinh = $dataC['gioitinh'];

                if ($dataCheckCustomer > 0 && $dataC['dangdat'] == "Đã đặt") {
                    echo '<script>alert("Khách hàng vẫn đang đặt phòng");</script>';
                    echo '<script>window.history.back();</script>';
                } else if ($dataCheckCustomer > 0 && $dataC['dangdat'] == "Đã trả") {
                    $roomOrdered->addRoomOrdered($roomID, $currentDateTime, '', $loaiphong, $gia, $username, $cccd, $sdt, $ngaysinh, $quoctich, $diachi, $gioitinh);
                    $roomModel->updateStatusRoom($roomID, 'Đã đặt', $username);
                    $customerModel->suatrangdat($username);

                    $dataR = $roomOrdered->getSHD($roomID, $username, $currentDateTime);
                    $shd = mysqli_fetch_row($dataR);

                    $hoaDonModel->addHoaDon($username, $roomID, $currentDateTime, '', '', 0, $shd[5], $tiencoc, $phuongthucthanhtoan);
                    echo '<script>window.history.back();</script>';
                    echo '<script>window.hicstory.back();</script>';
                } else {
                    $roomOrdered->addRoomOrdered($roomID, $currentDateTime, '', $loaiphong, $gia, $username, $cccd, $sdt, $ngaysinh, $quoctich, $diachi, $gioitinh);
                    $roomModel->updateStatusRoom($roomID, 'Đã đặt', $username);
                    $customerModel->addCustomer($username, $cccd, $sdt, $ngaysinh, $quoctich, $diachi, $gioitinh, 1, 'Đã đặt');

                    $dataR = $roomOrdered->getSHD($roomID, $username, $currentDateTime);
                    $shd = mysqli_fetch_row($dataR);

                    $hoaDonModel->addHoaDon($username, $roomID, $currentDateTime, '', '', 0, $shd[5], $tiencoc, $phuongthucthanhtoan);
                    echo '<script>window.history.back();</script>';
                    echo '<script>window.history.back();</script>';
                }
            }
        }


        //nhan phong
        if (isset($_POST['btnnhanphong'])) {
            $ngaynhanphong = date('Y-m-d H:i:s');

            $dataRoom = $roomModel->getARoom($roomID);
            $room = mysqli_fetch_row($dataRoom);


            $hoaDonModel->updateNgayNhan($room[3], $roomID, $ngaynhanphong);
            $roomModel->updateTrangthaiphong($roomID, 'Đã nhận');
            echo '<script>window.history.back();</script>';
        }


        //xoa phong
        if (isset($_POST['btnxoaphong'])) {

            $roomModel->deleteRoom($roomID);

            header("Location: ../../../learn_mvc/quanLyPhong");
        }





        //datdichvu
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



        //tra phong
        if (isset($_POST['btntraphong'])) {
            $tongthanhtoan = $_POST['tongthanhtoan'];

            $ngaytraphong = date('Y-m-d H:i:s');

            $dataRoom = $roomModel->getARoom($roomID);
            $room = mysqli_fetch_row($dataRoom);

            $tenkhachhang = $_POST['tenkh'];
            $ngaydat = $_POST['nd'];
            $dataR = $roomOrdered->getSHD($roomID, $tenkhachhang, $ngaydat);
            $sohopdong = mysqli_fetch_assoc($dataR);


            $hoaDonModel->updateNgayTra($sohopdong['sohopdong'], $ngaytraphong, $tongthanhtoan);
            $roomModel->updateTrangthaiphong($roomID, 'Đang trống');
            $roomModel->setPhongtrong($roomID);
            $customerModel->suatrangthaidat($room[3]);
            echo '<script>window.history.back();</script>';
        }

        if (isset($_GET['roomID'])) {
            // var_dump($_GET['roomID']);
        }
        $this->view("defaultLayout", [
            "container" => "ChiTietPhongForNhanVien",
            "room" => $room,
            "dataService" => $dataService,
            "hoaDonModel" => $hoaDonModel,
            "dataDichvu" => $dataDichvu
        ]);
    }


    public function sua($id)
    {

        $roomModel = $this->model("RoomModel");
        $hoaDonModel = $this->model("HoaDonModel");
        $dataRoom = $roomModel->getARoom($id);

        if (isset($_POST['btnsuaphong'])) {
            $room = mysqli_fetch_row($dataRoom);

            if ($room[2] == "Đang trống") {
                $tenphong = $_POST['tenphong'];
                $loaiphong = $_POST['loaiphong'];
                $trangthaiphong = $_POST['trangthaiphong'];
                $gia = $_POST['gia'];

                $roomModel->updateARoom($tenphong, $loaiphong, $trangthaiphong, '', $gia);
            } else if ($room[2] == "Đã đặt") {
                $tenphong = $_POST['tenphong'];
                $loaiphong = $_POST['loaiphong'];
                $trangthaiphong = $_POST['trangthaiphong'];
                $ngaydat = $_POST['ngaydat'];
                $khachhang = $_POST['khachhang'];
                $gia = $_POST['gia'];

                $hoaDonModel->updateHoaDon($khachhang, $tenphong, $ngaydat, '');
                $roomModel->updateARoom($tenphong, $loaiphong, 'Đã đặt', $khachhang, $gia);
            } else {
                $tenphong = $_POST['tenphong'];
                $loaiphong = $_POST['loaiphong'];
                $trangthaiphong = $_POST['trangthaiphong'];
                $ngaydat = $_POST['ngaydat'];
                $ngaynhan = $_POST['ngaynhan'];
                $khachhang = $_POST['khachhang'];
                $gia = $_POST['gia'];

                $hoaDonModel->updateHoaDon($khachhang, $tenphong, $ngaydat, $ngaynhan);
                $roomModel->updateARoom($tenphong, $loaiphong, $trangthaiphong, $ngaydat, $ngaynhan, $khachhang, $gia);
            }
        }


        $dataRoom = $roomModel->getARoom($id);



        $this->view("fullLayout", [
            "container" => "suaphong",
            "dataRoom" => $dataRoom,
            "hoaDonModel" => $hoaDonModel,
        ]);
    }
}
