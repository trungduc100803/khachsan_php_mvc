<?php
class CustomerModel extends DB
{

    public function getSoLanDat($name)
    {
        $query = "SELECT solandatphong FROM customer WHERE name = '$name'";
        $data = mysqli_query($this->connection, $query);

        $num =  mysqli_fetch_row($data);
        return $num[0];
    }

    public function getAllCustomer()
    {
        $query = "SELECT * FROM customer";

        return mysqli_query($this->connection, $query);
    }

    public function getACustomer($name)
    {
        $query = "SELECT * FROM customer WHERE name = '$name'";

        return mysqli_query($this->connection, $query);
    }

    public function getACustomerForCCCD($cccd)
    {
        $query = "SELECT * FROM customer WHERE cccd = '$cccd'";

        return mysqli_query($this->connection, $query);
    }

    public function getACustomerForID($ID)
    {
        $query = "SELECT * FROM customer WHERE customerID = '$ID'";

        return mysqli_query($this->connection, $query);
    }

    public function addCustomer($name, $cccd, $sdt, $ngaysinh, $quoctich, $diachi, $gioitinh, $solandat, $dangdat)
    {
        $query = "INSERT INTO customer(customerID, name, cccd, sdt, ngaysinh, quoctich, diachi, gioitinh, solandatphong, dangdat)
        VALUES('', '$name', '$cccd', '$sdt', '$ngaysinh', '$quoctich', '$diachi', '$gioitinh', $solandat, '$dangdat')";

        mysqli_query($this->connection, $query);
    }

    public function updateCustomer($name, $cccd, $sdt, $ngaysinh, $quoctich, $diachi, $gioitinh)
    {
        $query = "UPDATE customer
        SET cccd = '$cccd', sdt = '$sdt', ngaysinh ='$ngaysinh', quoctich = '$quoctich', diachi = '$diachi', gioitinh = '$gioitinh' WHERE name = '$name'";

        mysqli_query($this->connection, $query);
    }


    public function suatrangthaidat($name)
    {
        $query = "UPDATE customer SET dangdat= 'Đã trả' WHERE name = '$name'";
        mysqli_query($this->connection, $query);
    }

    public function suatrangdat($name)
    {
        $query = "UPDATE customer SET dangdat= 'Đã đặt' WHERE name = '$name'";
        mysqli_query($this->connection, $query);
    }
}
