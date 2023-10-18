<?php
class RoomOrdered extends DB
{
    public function addRoomOrdered($roomID, $ngaydat, $ngaynhan, $loaiphong, $gia, $name, $cccd, $sdt, $ngaysinh, $quoctich, $diachi, $gioitinh)
    {
        $query = "INSERT INTO rooms_ordered VALUES('$roomID', '$ngaydat', '$ngaynhan', '$loaiphong', '$gia', '', '$name', '$cccd', '$sdt', '$ngaysinh', '$quoctich', '$diachi', '$gioitinh')";

        mysqli_query($this->connection, $query);
    }

    public function layTatCaTheoTen($name)
    {
        $query = "SELECT * FROM rooms_ordered WHERE name = '$name'";

        return mysqli_query($this->connection, $query);
    }


    public function getSHD($roomID, $name, $ngaydat,)
    {
        $query = "SELECT * FROM rooms_ordered WHERE roomID = '$roomID' AND name = '$name' AND ngaydat= '$ngaydat'";

        return mysqli_query($this->connection, $query);
    }

    public function layPhongDaDatQuaSHD($sohopdong)
    {
        $query = "SELECT * FROM rooms_ordered WHERE sohopdong = '$sohopdong'";

        return mysqli_query($this->connection, $query);
    }
}
