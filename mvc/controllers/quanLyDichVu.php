<?php
class quanLyDichVu extends Controller
{
    public function start()
    {
        $serviceModel = $this->model("ServiceModel");


        if (isset($_POST['themdichvu'])) {
            $tendichvu = $_POST['tendichvu'];
            $giadichvu = $_POST['giadichvu'];

            $serviceModel->themdichvu($tendichvu, $giadichvu);
        }







        $dataService = $serviceModel->getAllService();

        $this->view("defaultLayout", [
            "container" => "quanLyDichVu",
            "service" => $dataService
        ]);
    }

    public function action($id)
    {
        $serviceModel = $this->model("ServiceModel");

        $serviceModel->xoadichvu($id);
        header("Location: ../../../../learn_mvc/quanLyDichVu");
    }


    public function sua($id)
    {
        $serviceModel = $this->model("ServiceModel");

        if (isset($_POST['dongdv'])) {
            header("Location: ../../../../learn_mvc/quanLyDichVu");
        }

        if (isset($_POST['suadv'])) {
            $tendv = $_POST['tendv'];
            $giadv = $_POST['giadv'];

            $serviceModel->suadichvu($id, $tendv, $giadv);
            header("Location: ../../../../learn_mvc/quanLyDichVu");
        }

        $dataS = $serviceModel->getAService($id);

        $this->view("fullLayout", [
            "container" => "suadichvu",
            "service" => $dataS
        ]);
    }
}
