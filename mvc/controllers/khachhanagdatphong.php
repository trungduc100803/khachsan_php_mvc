<?php
class khachhanagdatphong extends Controller
{
    public function start()
    {

        $roomModel = $this->model("RoomModel");





        $dataRoomAlone = $roomModel->getRoomTypeAndNoOrder("Phòng đơn");
        $dataRoomCouple = $roomModel->getRoomTypeAndNoOrder("Phòng đôi");
        $dataRoomFamily = $roomModel->getRoomTypeAndNoOrder("Phòng gia đình");



        $this->view("defaultLayout", [
            "container" => "khachhanagdatphong",
            "roomAlone" => $dataRoomAlone,
            "roomCouple" => $dataRoomCouple,
            "roomFamily" => $dataRoomFamily
        ]);
    }
}
