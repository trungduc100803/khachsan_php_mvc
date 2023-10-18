<?php
$room = $data['room'];
$hoadon = $data['hoaDonModel'];
$services = $data['dataService'];
$sldichvu = 0;
?>

<form action="" method="post">
    <div class="chitietphong" style="display: flex; margin-top: 50px; padding-top: 30px; ">
        <div class="chitiet_left" style="width: 100%;">
            <h4>Thông tin chi tiết của phòng</h4>
            <div class="phong_content" style="margin-top: 30px;"></div>
            <div class="phong_item" style="display: flex; align-items: stretch;">
                <span style="margin-right: 10px; font-weight: 600; ">Tên phòng:</span>
                <p><?php echo "P" . $room[0]; ?></p>
            </div>
            <div class="phong_item" style="display: flex; align-items: stretch;">
                <span style="margin-right: 10px; font-weight: 600; ">Loại phòng:</span>
                <p><?php echo  $room[1]; ?></p>
            </div>
            <div class="phong_item" style="display: flex; align-items: stretch;">
                <span style="margin-right: 10px; font-weight: 600; ">Trạng thái hiện tại:</span>
                <p><?php echo $room[2]; ?></p>
            </div>
            <div class="phong_item" style="display: flex; align-items: stretch;">
                <span style="margin-right: 10px; font-weight: 600; ">Giá phòng:</span>
                <p><?php echo $room[4]; ?></p>
            </div>
            <div class="phong_item" style="display: flex; align-items: stretch;">
                <span style="margin-right: 10px; font-weight: 600; ">Tình trạng:</span>
                <p>Đã dọn dẹp</p>
            </div>
            <div class="phong_item" style="display: flex; align-items: stretch;">
                <span style="margin-right: 10px; font-weight: 600; width: 200px; ">Mô tả:</span>
                <p><?php echo $room[6]; ?></p>
            </div>
            <div class="phong_item" style="display: flex; align-items: stretch; margin-bottom: 30px;">
                <span style="margin-right: 10px; font-weight: 600; width: 200px; ">Đồ dùng trong phòng:</span>
                <p><?php echo $room[7]; ?></p>
            </div>
            <button type="button" data-bs-toggle="modal" data-bs-target="#exampleModal" class="btn btn-success">Đặt phòng</button>
            <!-- <button type="button" class="btn btn-danger" data-bs-toggle="modal" data-bs-target="#exampleModalXoaphong">Xóa phòng</button>
                    <a href="../../../../learn_mvc/chitietphong/sua/<?php echo $room[0] ?>" type="submit" name="suaphong" class="btn btn-warning">Sửa phòng</a> -->

        </div>
    </div>


    <!-- modal  dat phong -->
    <div style="margin-top: 40px;">
        <div class="img_hotel" style="width: 100%; height: 500px;">
            <img src="<?php echo $room[5]; ?>" style="width: 100%; height: 100%; object-fit: contain;" alt="">
        </div>
    </div>

</form>



<form method="POST" class="modal fade" id="exampleModal" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
                <h1 class="modal-title fs-5" id="exampleModalLabel">Đặt phòng</h1>
                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
            </div>
            <div class="modal-body">
                <!-- <label for="">Tên khách hàng</label>
                <input name="tenkhachhang" style="margin-top: 6px; margin-bottom: 20px;" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">


                <label for="">Số CCCD/CMND</label>
                <input name="cccd" style="margin-top: 6px; margin-bottom: 20px;" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">


                <label for="">Số điện thoại</label>
                <input name="sdt" style="margin-top: 6px; margin-bottom: 20px;" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">


                <label for="">Ngày sinh</label>
                <input name="ngaysinh" style="margin-top: 6px; margin-bottom: 20px;" type="date" class="form-control" aria-label="" aria-describedby="basic-addon1">

                <label for="">Quốc tịch</label>
                <select name="quoctich" style="margin-top: 6px; margin-bottom: 20px;" class="form-select" aria-label="Default select example">
                    <option selected>-- QUỐC TỊCH --</option>
                    <option value="Việt Nam">Việt Nam</option>
                    <option value="Nhật Bản">Nhật Bản</option>
                    <option value="Hàn Quốc">Hàn Quốc</option>
                </select>

                <label for="">Địa chỉ</label>
                <select name="diachi" style="margin-top: 6px; margin-bottom: 20px;" class="form-select" aria-label="Default select example">
                    <option selected>-- ĐỊA CHỈ --</option>
                    <option value="Hà Nội">Hà Nội</option>
                    <option value="Thanh Hóa">Thanh Hóa</option>
                    <option value="Bắc Ninh">Bắc Ninh</option>
                </select> -->

                <label for="">Phương thức thanh toán</label>
                <div style="margin-top: 6px; margin-bottom: 20px;" class="check">
                    <input class="form-check-input" type="radio" name="phuongthucthanhtoan" value="Tiền mặt" id="flexRadioDefault1">
                    <label class="form-check-label" for="flexRadioDefault1">
                        Tiền mặt
                    </label>

                    <input class="form-check-input" type="radio" name="phuongthucthanhtoan" value="Tài khoản" id="flexRadioDefault1">
                    <label class="form-check-label" for="flexRadioDefault1">
                        Tài khoản
                    </label>
                </div>

                <label for="">Số tiền cọc</label>
                <input name="tiencoc" value="20000" style="margin-top: 6px; margin-bottom: 20px;" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">

            </div>
            <div class="modal-footer">
                <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Đóng</button>
                <button name="datphong" type="submit" class="btn btn-primary">Lưu</button>
            </div>
        </div>
    </div>
</form>

<!-- Modal nhan phong -->
<form method="post" class="modal fade" id="exampleModalNhanphong" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
                <h1 class="modal-title fs-5" id="exampleModalLabel">Nhận phòng</h1>
                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
            </div>
            <div class="modal-body">
                <label for="">Ngày nhận</label>
                <input name="ngaynhanphong" style="margin-top: 6px; margin-bottom: 20px;" type="date" class="form-control" aria-label="" aria-describedby="basic-addon1">

            </div>
            <div class="modal-footer">
                <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
                <button type="submit" name="btnnhanphong" class="btn btn-primary">Lưu</button>
            </div>
        </div>
    </div>
</form>

<!-- Modal tra phong -->
<form method="post" class="modal fade" id="exampleModalTraphong" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
                <h1 class="modal-title fs-5" id="exampleModalLabel">Trả phòng</h1>
                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
            </div>
            <div class="modal-body">
                <label for="">Ngày trả phòng</label>
                <input name="ngaytraphong" style="margin-top: 6px; margin-bottom: 20px;" type="date" class="form-control" aria-label="" aria-describedby="basic-addon1">

            </div>
            <div class="modal-footer">
                <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
                <button type="submit" name="btntraphong" class="btn btn-primary">Trả</button>
            </div>
        </div>
    </div>
</form>

<!-- Modal xoa phong -->
<form method="post" class="modal fade" id="exampleModalXoaphong" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
                <h1 class="modal-title fs-5" id="exampleModalLabel">Nhận phòng</h1>
                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
            </div>
            <div class="modal-body">
                Bạn thực sự muốn xóa phòng này 🐱‍👤🐱‍👤🐱‍👤🐱‍👤
            </div>
            <div class="modal-footer">
                <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
                <button type="submit" name="btnxoaphong" class="btn btn-primary">Xóa</button>
            </div>
        </div>
    </div>
</form>