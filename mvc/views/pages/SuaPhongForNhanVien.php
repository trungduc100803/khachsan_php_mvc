<?php
$dataRoom = $data['dataRoom'];
$room = mysqli_fetch_row($dataRoom);
// print_r($room);
?>


<form style="background-image: linear-gradient(to right, #C8E4B2, white);" method="post" class="modals" id="exampleModalsua" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
    <div class="modal-dialog" style="padding: 100px 400px;">
        <div class="modal-content" style="padding: 20px; background-color: white; border-radius: 10px; box-shadow: rgba(100, 100, 111, 0.2) 0px 7px 29px 0px;">
            <div class="modal-header">
                <h1 class="modal-title fs-5" style="margin-bottom: 20px; text-align: center; font-size: 40px;" id="exampleModalLabel">Sửa thông tin phòng</h1>
            </div>
            <div class="modal-body">
                <?php
                if ($room[2] == "Đang trống") {
                ?>

                    <label for="">Tên phòng</label>
                    <input name="tenphong" style="margin-top: 6px; margin-bottom: 20px;" readonly value="<?php echo  $room[0] ?>" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">

                    <label for="">Loại phòng</label>
                    <input name="loaiphong" style="margin-top: 6px; margin-bottom: 20px;" readonly value="<?php echo $room[1] ?>" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">

                    <label for="">Trạng thái phòng</label>
                    <select name="trangthaiphong" style="margin-top: 6px; margin-bottom: 20px;" value="<?php echo $room[2] ?>" class="form-select" aria-label="Default select example">
                        <option value="Đang trống" selected>Đang trống</option>
                        <option value="Đã đặt">Đã đặt</option>
                        <option value="Đã nhận">Đã nhận</option>
                    </select>

                    <label for="">Giá</label>
                    <select name="gia" style="margin-top: 6px; margin-bottom: 20px;" value="<?php echo $room[6] ?>" class="form-select" aria-label="Default select example">
                        <option selected>Giá</option>
                        <option value="1000000">1000000</option>
                        <option value="1500000">1500000</option>
                        <option value="2000000">2000000</option>
                    </select>
            </div>
        <?php
                } else if ($room[2] == "Đã đặt") {
                    $hoaDonModel = $data["hoaDonModel"];
                    $datangaydat = $hoaDonModel->getAHoaDon($room[3], $room[0]);
                    $ngaydat = mysqli_fetch_row($datangaydat);
        ?>
            <label for="">Tên phòng</label>
            <input name="tenphong" style="margin-top: 6px; margin-bottom: 20px;" readonly value="<?php echo $room[0] ?>" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">

            <label for="">Loại phòng</label>
            <input name="loaiphong" style="margin-top: 6px; margin-bottom: 20px;" readonly value="<?php echo $room[1] ?>" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">


            <label for="">Trạng thái phòng</label>
            <select name="trangthaiphong" style="margin-top: 6px; margin-bottom: 20px;" value="<?php echo $room[2] ?>" class="form-select" aria-label="Default select example">
                <option value="Đang trống">Đang trống</option>
                <option value="Đã đặt" selected>Đã đặt</option>
                <option value="Đã nhận">Đã nhận</option>
            </select>

            <label for="">Ngày đặt</label>
            <input name="ngaydat" style="margin-top: 6px; margin-bottom: 20px;" value="<?php echo $ngaydat[3] ?>" type="date" class="form-control" aria-label="" aria-describedby="basic-addon1">

            <label for="">Khách hàng</label>
            <input name="khachhang" style="margin-top: 6px; margin-bottom: 20px;" value="<?php echo $ngaydat[1] ?>" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">


            <label for="">Giá</label>
            <select name="gia" style="margin-top: 6px; margin-bottom: 20px;" value="<?php echo $room[6] ?>" class="form-select" aria-label="Default select example">
                <option selected>Giá</option>
                <option value="1000000">1000000</option>
                <option value="1500000">1500000</option>
                <option value="2000000">2000000</option>
            </select>
        <?php } else {
                    $hoaDonModel = $data["hoaDonModel"];
                    $datangaydat = $hoaDonModel->getAHoaDon($room[3], $room[0]);
                    $ngaydat = mysqli_fetch_row($datangaydat);
        ?>
            <label for="">Tên phòng</label>
            <input name="tenphong" style="margin-top: 6px; margin-bottom: 20px;" readonly value="<?php echo $room[0] ?>" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">

            <label for="">Loại phòng</label>
            <input name="loaiphong" style="margin-top: 6px; margin-bottom: 20px;" readonly value="<?php echo $room[1] ?>" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">


            <label for="">Trạng thái phòng</label>
            <select name="trangthaiphong" style="margin-top: 6px; margin-bottom: 20px;" value="<?php echo $room[2] ?>" class="form-select" aria-label="Default select example">
                <option value="Đang trống">Đang trống</option>
                <option value="Đã đặt">Đã đặt</option>
                <option value="Đã nhận" selected>Đã nhận</option>
            </select>

            <label for="">Ngày đặt</label>
            <input name="ngaydat" style="margin-top: 6px; margin-bottom: 20px;" value="<?php echo $ngaydat[3] ?>" type="date" class="form-control" aria-label="" aria-describedby="basic-addon1">

            <label for="">Ngày nhận</label>
            <input name="ngaynhan" style="margin-top: 6px; margin-bottom: 20px;" value="<?php echo $ngaydat[4] ?>" type="date" class="form-control" aria-label="" aria-describedby="basic-addon1">

            <label for="">Khách hàng</label>
            <input name="khachhang" style="margin-top: 6px; margin-bottom: 20px;" value="<?php echo $ngaydat[1] ?>" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">


            <label for="">Giá</label>
            <select name="gia" style="margin-top: 6px; margin-bottom: 20px;" value="<?php echo $room[6] ?>" class="form-select" aria-label="Default select example">
                <option selected>Giá</option>
                <option value="1000000">1000000</option>
                <option value="1500000">1500000</option>
                <option value="2000000">2000000</option>
            </select>
        <?php } ?>
        <div class="modal-footer">
            <a href="../../../../learn_mvc/QuanLyPhongForNhanVien" type="submit" class="btn btn-secondary" style="margin-right: 20px;">Đóng</a>
            <button type="submit" name="btnsuaphong" class="btn btn-primary">Lưu✔✔</button>
        </div>
        </div>
    </div>
</form>