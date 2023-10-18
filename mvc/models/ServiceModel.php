<?php
class ServiceModel extends DB
{

    public function addService($name, $sohopdong, $tendichvu, $soluong, $gia, $thanhtien)
    {
        $query = "INSERT INTO dichvudadangki VALUES('', '$name', '$sohopdong', '$tendichvu', '$soluong', '$gia', '$thanhtien')";

        mysqli_query($this->connection, $query);
    }

    public function layDichVuDaDatVoiSHD($sohopdong)
    {
        $query = "SELECT * FROM dichvudadangki WHERE sohopdong = '$sohopdong'";

        return mysqli_query($this->connection, $query);
    }

    public function getAllService()
    {
        $q = "SELECT * FROM service";
        return mysqli_query($this->connection, $q);
    }

    public function layDichVuDaDangKyQuaSHD($sohopdong)
    {
        $q = "SELECT * FROM dichvudadangki WHERE sohopdong ='$sohopdong'";
        return mysqli_query($this->connection, $q);
    }

    public function getAService($id)
    {
        $q = "SELECT * FROM service WHERE id = '$id'";
        return mysqli_query($this->connection, $q);
    }




    public function themdichvu($tendichvu, $giadichvu)
    {
        $query = "INSERT INTO service VALUES('', '$tendichvu', '$giadichvu', 1)";

        mysqli_query($this->connection, $query);
    }


    public function xoadichvu($id)
    {
        $query = "DELETE FROM service WHERE id = '$id'";

        mysqli_query($this->connection, $query);
    }


    public function suadichvu($id, $tendichvu, $gia)
    {
        $query = "UPDATE  service SET tendichvu = '$tendichvu', gia='$gia'  WHERE id = '$id'";

        mysqli_query($this->connection, $query);
    }
    public function getService($id)
    {
        $q = "SELECT * FROM service WHERE id = '$id'";
        return mysqli_query($this->connection, $q);
    }
}
