<?php
class HoaDonModel extends DB
{
    public function getAHoaDon($name, $roomID)
    {
        $query = "SELECT * FROM hoadon WHERE name = '$name' AND roomID = '$roomID'";

        return mysqli_query($this->connection, $query);
    }


    public function layTatCaHoaDon($name)
    {
        $query = "SELECT * FROM hoadon WHERE name = '$name'";

        return mysqli_query($this->connection, $query);
    }

    public function getAllHoaDon()
    {
        $query = "SELECT * FROM hoadon";

        return mysqli_query($this->connection, $query);
    }

    public function getTheoNgay($n)
    {
        $query = "SELECT * FROM hoadon";
        $arr = [];
        $dataQuery = mysqli_query($this->connection, $query);
        while ($data = mysqli_fetch_array($dataQuery)) {
            $ngayGio = $data['ngaytra'];
            list($ngay, $gio) = explode(" ", $ngayGio);

            $ngayHienTai = $ngay;
            list($nam, $thang, $ngay) = explode("-", $ngayHienTai);
            if ($ngay == $n) {
                array_push($arr, $data);
            }
        }

        return $arr;
    }

    public function getTheoThang($t)
    {
        $query = "SELECT * FROM hoadon";
        $arr = [];
        $dataQuery = mysqli_query($this->connection, $query);
        while ($data = mysqli_fetch_array($dataQuery)) {
            $ngayGio = $data['ngaytra'];
            list($ngay, $gio) = explode(" ", $ngayGio);

            $ngayHienTai = $ngay;
            list($nam, $thang, $ngay) = explode("-", $ngayHienTai);
            if ($thang == $t) {
                array_push($arr, $data);
            }
        }

        return $arr;
    }


    public function getTheoNam($na)
    {
        $query = "SELECT * FROM hoadon";
        $arr = [];
        $dataQuery = mysqli_query($this->connection, $query);
        while ($data = mysqli_fetch_array($dataQuery)) {
            $ngayGio = $data['ngaytra'];
            list($ngay, $gio) = explode(" ", $ngayGio);

            $ngayHienTai = $ngay;
            list($nam, $thang, $ngay) = explode("-", $ngayHienTai);
            if ($nam == $na) {
                array_push($arr, $data);
            }
        }

        return $arr;
    }
    public function addHoaDon($name, $roomID, $ngaydat, $ngaynhan, $ngaytra, $tongchiphi, $shd, $tiencoc, $phttTiencoc)
    {
        $query = "INSERT INTO  hoadon VALUES('', '$name', '$roomID', '$shd' , '$ngaydat', '$ngaynhan', '$ngaytra', '$tongchiphi', '$tiencoc', '$phttTiencoc')";

        mysqli_query($this->connection, $query);
    }


    public function updateHoaDon($name, $roomID, $ngaydat, $ngaynhan)
    {
        $query = "UPDATE hoadon SET ngaydat = '$ngaydat', ngaynhan = '$ngaynhan' WHERE name = '$name' AND roomID = '$roomID' ";

        mysqli_query($this->connection, $query);
    }

    public function updateNgayNhan($name, $roomID, $ngaynhan)
    {
        $query = "UPDATE hoadon SET  ngaynhan = '$ngaynhan' WHERE name = '$name' AND roomID = '$roomID' ";

        mysqli_query($this->connection, $query);
    }

    public function updateNgayTra($sohopdong, $ngaytra, $tongchiphi)
    {
        $query = "UPDATE hoadon SET  ngaytra = '$ngaytra', tongchiphi = '$tongchiphi' WHERE sohopdong='$sohopdong' ";

        mysqli_query($this->connection, $query);
    }
}
