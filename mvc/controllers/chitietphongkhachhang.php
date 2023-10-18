<?php
class chitietphongkhachhang extends Controller
{
    public function start($roomID)
    {
        $roomModel = $this->model("RoomModel");
        $customerModel = $this->model("CustomerModel");
        $roomOrdered = $this->model("RoomOrdered");
        $serviceModel = $this->model("ServiceModel");
        $hoaDonModel = $this->model("HoaDonModel");

        ///xu ly dat phong
        if (isset($_POST['datphong'])) {

            // lay du lieu tu form dat phong
            $phuongthucthanhtoan = $_POST['phuongthucthanhtoan'];
            $tiencoc = $_POST['tiencoc'];
            $username = $_SESSION['username'];
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

            // print_r($cccd);

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
                echo '<script>window.history.back();</script>';
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


        //nhan phong
        if (isset($_POST['btnnhanphong'])) {

            $ngaynhanphong = $_POST['ngaynhanphong'];

            $dataRoom = $roomModel->getARoom($roomID);
            $room = mysqli_fetch_row($dataRoom);


            $hoaDonModel->updateNgayNhan($room[3], $roomID, $ngaynhanphong);
            $roomModel->updateTrangthaiphong($roomID, 'Đã nhận');
        }


        //xoa phong
        if (isset($_POST['btnxoaphong'])) {

            $roomModel->deleteRoom($roomID);

            header("Location: ../../../learn_mvc/quanLyPhong");
        }


        //tra phong
        if (isset($_POST['btntraphong'])) {
            $ngaytraphong = $_POST['ngaytraphong'];

            $dataRoom = $roomModel->getARoom($roomID);
            $room = mysqli_fetch_row($dataRoom);


            $hoaDonModel->updateNgayTra($room[3], $roomID, $ngaytraphong);
            $roomModel->updateTrangthaiphong($roomID, 'Đang trống');
            $roomModel->setPhongtrong($roomID);
            $customerModel->suatrangthaidat($room[3]);
        }


        //datdichvu
        if (isset($_POST['datdichvu'])) {

            $n = $_POST['aa'];
            echo $n;
        }


        $dataService = $serviceModel->getAllService();
        $dataRoom =  $roomModel->getARoom($roomID);
        $room = mysqli_fetch_row($dataRoom);

        if (isset($_GET['roomID'])) {
            // var_dump($_GET['roomID']);
        }
        $this->view("defaultLayout", [
            "container" => "chitietphongkhachhang",
            "room" => $room,
            "dataService" => $dataService,
            "hoaDonModel" => $hoaDonModel
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
            "hoaDonModel" => $hoaDonModel
        ]);
    }
}
