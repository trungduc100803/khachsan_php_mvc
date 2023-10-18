<?php
    class traphongkhachhang extends Controller{
        public function start(){
            $this->view("defaultLayout", [
                "container"=> "traphongkhachhang"
            ]);
        }
    }
