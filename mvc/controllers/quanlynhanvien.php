<?php
class quanlynhanvien extends Controller
{
    public function start()
    {
        $this->view("defaultLayout", [
            "container" => "quanlynhanvien"
        ]);
    }
}
