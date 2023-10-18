<?php
class QuanLyPhongForNhanVien  extends Controller
{
    public function start()
    {
        $roomModel = $this->model("RoomModelForNhanVien");
        $hoadonModel = $this->model("HoaDonModel");

        // print_r($dataAllRoom);

        $dataRoomAlone = $roomModel->getRoomType("Phòng đơn");
        $dataRoomCouple = $roomModel->getRoomType("Phòng đôi");
        $dataRoomFamily = $roomModel->getRoomType("Phòng gia đình");

        $this->view("defaultLayout", [
            "container" => "QuanLyPhongForNhanVien",
            "roomAlone" => $dataRoomAlone,
            "roomCouple" => $dataRoomCouple,
            "roomFamily" => $dataRoomFamily,
            "hoadonModel" => $hoadonModel
        ]);
    }
}
