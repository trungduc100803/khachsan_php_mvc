<?php
class dichvukhachhang extends Controller
{
    public function  start()
    {
        $roomOrdered = $this->model("RoomOrdered");
        $serviceModel = $this->model("ServiceModel");

        // session_start();
        $tenkhachhang = $_SESSION['username'];
        $ngaydat = $_SESSION['ngaydat'];
        $roomID = $_SESSION['roomID'];

        $dataR = $roomOrdered->getSHD($roomID, $tenkhachhang, $ngaydat);
        $sohopdong = mysqli_fetch_assoc($dataR);


        $dataDichvu = $serviceModel->layDichVuDaDatVoiSHD($sohopdong['sohopdong']);
        // print_r($dataDichvu);
        $this->view("defaultLayout", [
            "container" => "dichvukhachhang",
            "dataDichvu" => $dataDichvu
        ]);
    }

    public function delete()
    {
    }

    public function sua()
    {
    }
}
