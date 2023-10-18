<?php
class hoadon extends Controller
{
    public function start()
    {

        $this->view("defaultLayout", [
            "container" => "hoadon",

        ]);
    }

}
