<?php
class hoadonkhachhang extends Controller
{


    public function start()
    {

        $HoaDonModel = $this->model("HoaDonModel");

        $dataHoaDon = $HoaDonModel->layTatCaHoaDon($_SESSION['username']);

        $this->view("defaultLayout", [
            "container" => "hoadonkhachhang",
            "dataHoaDon" => $dataHoaDon

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
