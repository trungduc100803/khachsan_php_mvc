<?php
class home extends Controller
{

    public function start()
    {
        $data = $this->model("RoomModel");
        $n = $data->getAllRoom();
        $this->view("defaultLayout", [
            "n" => $n,
            "container" => "home"
        ]);
    }

    public function slider()
    {
        // print_r($arr_url);
    }
}
