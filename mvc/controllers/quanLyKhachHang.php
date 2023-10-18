<?php
class quanLyKhachHang extends Controller
{
    public function start()
    {

        $customerModel = $this->model("CustomerModel");
        $dataCustomers = $customerModel->getAllCustomer();



        $this->view("defaultLayout", [
            "container" => "quanLyKhachHang",
            "customers" => $dataCustomers,
        ]);
    }

    public function edit($ID)
    {
        $customerModel = $this->model("CustomerModel");


        if (isset($_POST['btnsuakh'])) {
            $tenkhachhang = $_POST['tenkhachhang'];
            $cccd = $_POST['cccd'];
            $sdt = $_POST['sdt'];
            $ngaysinh = $_POST['ngaysinh'];
            $quoctich = $_POST['quoctich'];
            $diachi = $_POST['diachi'];
            $gioitinh = $_POST['gioitinh'];

            $customerModel->updateCustomer($tenkhachhang, $cccd, $sdt, $ngaysinh, $quoctich, $diachi, $gioitinh);
            header('Location: ../../../learn_mvc/quanLyKhachHang');
        }

        $dataCustomer = $customerModel->getACustomerForID($ID);

        $this->view("fullLayout", [
            "container" => "suathongtinkhachhang",
            "customer" => $dataCustomer,

        ]);
    }


    public function chitietkhachhang($id, $name)
    {
        $customerModel = $this->model("CustomerModel");
        $RoomOrdered = $this->model("RoomOrdered");

        $data_phong_da_dat = $RoomOrdered->layTatCaTheoTen($name);
        $dataCustomer = $customerModel->getACustomerForID($id);

        $this->view("defaultLayout", [
            "container" => "chitietkhachhang",
            "customer" => $dataCustomer,
            "data_phong_da_dat" => $data_phong_da_dat
        ]);
    }
}
