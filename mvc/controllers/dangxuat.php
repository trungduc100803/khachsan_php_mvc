<?php
class dangxuat extends Controller
{
    public function start()
    {
        $_SESSION['role'] = "";
        header("Location: ../../../learn_mvc/home");
    }
}
