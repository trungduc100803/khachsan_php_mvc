<?php
class App
{
    protected $controller = "home";
    protected $action = "start";
    protected $params = [];


    public function __construct()
    {
        $arr_url = $this->URL_process();


        if (file_exists("../learn_mvc/mvc/controllers/" . $arr_url[0] . ".php")) {
            $this->controller = $arr_url[0];
            unset($arr_url[0]);
        }
        require_once "../learn_mvc/mvc/controllers/" . $this->controller . ".php";
        $this->controller = new $this->controller;



        if (isset($arr_url[1])) {
            if (method_exists($this->controller, $arr_url[1])) {
                $this->action = $arr_url[1];
            }
            unset($arr_url[1]);
        }

        $this->params = $arr_url ? $arr_url : [];

        call_user_func_array([$this->controller, $this->action], $this->params);
    }


    function URL_process()
    {
        if (isset($_GET['url'])) {
            return explode("/", filter_var(trim($_GET['url'], "/")));
        }
    }
}
