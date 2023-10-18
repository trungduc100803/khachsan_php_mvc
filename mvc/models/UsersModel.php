<?php
class UsersModel extends DB
{
    public function getAllUser()
    {
        $query = "SELECT * FROM users";
        $data = mysqli_query($this->connection, $query);

        return $data;
    }

    public function getAUserToLogin($username, $password)
    {
        $query = "SELECT * FROM users WHERE name = '$username' AND password = '$password'";
        $datas = mysqli_query($this->connection, $query);

        return $datas;
    }
}
