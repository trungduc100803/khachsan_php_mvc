<?php
    class taikhoancanhan extends Controller{
        public function start(){
            $this->view("fullLayout", [
                "container"=> "taikhoancanhan"
            ]);
        }
    }
