<?php
class login extends Controller
{
    public function start()
    {
        $this->view("fullLayout", [
            "container" => "login",
        ]);


        if (isset($_POST['login'])) {
            $tendangnhap = $_POST['tendangnhap'];
            $mk = $_POST['matkhau'];

            $userModel = $this->model("UsersModel");
            $dataUser = $userModel->getAUserToLogin($tendangnhap, $mk);

            $user = mysqli_num_rows($dataUser);

            if ($user > 0) {
                $userQ = mysqli_fetch_row($dataUser);
                $_SESSION['role'] = $userQ[4];
                $_SESSION['username'] = $userQ[1];
                header("location: ./home");
            } else {
                echo "err";
            }
            // var_dump($user);

            // echo $mk, $tendangnhap;

        }
    }
}
