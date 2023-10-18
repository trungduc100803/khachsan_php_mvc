<?php
class quanLyPhong extends Controller
{
    public function start()
    {
        $roomModel = $this->model("RoomModel");
        $hoadonModel = $this->model("HoaDonModel");

        // print_r($dataAllRoom);
        if (isset($_POST['themphong'])) {
            $tenphong = $_POST['tenphong'];
            $loaiphong = $_POST['loaiphong'];
            $trangthaiphong = $_POST['trangthaiphong'];
            $gia = $_POST['gia'];

            $checkRoom = $roomModel->getARoom($tenphong);
            if (mysqli_num_rows($checkRoom) > 0) {
                echo "Đã tồn tại";
            } else {
                $roomModel->addRoom($tenphong, $loaiphong, $trangthaiphong, $trangthaiphong, $gia);
            }

            // echo $tenphong, $loaiphong, $trangthaiphong, $gia;
        }
        // $loc = "";

        if (isset($_POST['loc'])) {
            $loc = $_POST['locphong'];
        }


        $dataRoomAlone = $roomModel->getRoomType("Phòng đơn");
        $dataRoomCouple = $roomModel->getRoomType("Phòng đôi");
        $dataRoomFamily = $roomModel->getRoomType("Phòng gia đình");

        $this->view("defaultLayout", [
            "container" => "quanLyPhong",
            "roomAlone" => $dataRoomAlone,
            "roomCouple" => $dataRoomCouple,
            "roomFamily" => $dataRoomFamily,
            "hoadonModel" => $hoadonModel,
            "locphong" => $loc
        ]);
    }
}
