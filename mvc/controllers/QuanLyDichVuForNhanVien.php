<?php
class QuanLyDichVuForNhanVien extends Controller
{
    public function start()
    {
        $serviceModel = $this->model("ServiceModel");

        if (isset($_POST['themdichvu'])) {
            $tendichvu = $_POST['tendichvu'];
            $giadichvu = $_POST['giadichvu'];

            $serviceModel->themdichvu($tendichvu, $giadichvu);
        }

        $dataservice = $serviceModel->getAllService();

        $this->view("defaultLayout", [
            "container" => "QuanLyDichVuForNhanVien",
            "dataService" => $dataservice,
        ]);
    }

    public function sua($id)
    {

        $serviceModel = $this->model("ServiceModel");


        if (isset($_POST['btnsuadv'])) {
            $ten = $_POST['txtten'];
            $soluong = $_POST['txtsoluong'];
            $gia = $_POST['txtgiadv'];

            $serviceModel->suadichvu($id, $ten, $soluong, $gia);
        }

        if (isset($_POST['btnxoa'])) {
            $serviceModel->DeleteService($id);
            header('Location: ../../../learn_mvc/QuanLyDichVuForNhanVien');
        }

        $dataservice = $serviceModel->getService($id);

        $this->view("fullLayout", [
            "container" => "suadichvu",
            "service" => $dataservice,

        ]);
    }
}
