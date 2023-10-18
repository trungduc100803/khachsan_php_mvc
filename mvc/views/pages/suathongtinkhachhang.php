<?php
$customer = $data['customer'];
$dataC  = mysqli_fetch_row($customer);
?>

<form style="background-image: linear-gradient(to right, #C8E4B2, white);" method="post" class="modals" id="exampleModalsua" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
    <div class="modal-dialog" style="padding: 100px 400px;">
        <div class="modal-content" style="padding: 20px; background-color: white; border-radius: 10px; box-shadow: rgba(100, 100, 111, 0.2) 0px 7px 29px 0px;">
            <div class="modal-header">
                <h1 class="modal-title fs-5" style="margin-bottom: 20px; text-align: center; font-size: 40px;" id="exampleModalLabel">Sửa thông tin khách hàng</h1>
            </div>
            <div class="modal-body">

                <label for="">Tên khách hàng</label>
                <input disabled name="tenkhachhang" style="margin-top: 6px; margin-bottom: 20px;" value="<?php echo $dataC[1] ?>" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">

                <label for="">Số CCCD</label>
                <input name="cccd" style="margin-top: 6px; margin-bottom: 20px;" value="<?php echo $dataC[2] ?>" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">

                <label for="">Số điện thoại</label>
                <input name="sdt" style="margin-top: 6px; margin-bottom: 20px;" value="<?php echo $dataC[3] ?>" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">

                <label for="">Ngày sinh</label>
                <input name="ngaysinh" style="margin-top: 6px; margin-bottom: 20px;" value="<?php echo $dataC[4] ?>" type="date" class="form-control" aria-label="" aria-describedby="basic-addon1">

                <label for="">Quốc tịch</label>
                <select name="quoctich" style="margin-top: 6px; margin-bottom: 20px;" value="<?php echo $dataC[5] ?>" class="form-select" aria-label="Default select example">
                    <option selected>Quốc tịch</option>
                    <option value="Việt Nam">Việt Nam</option>~
                    <option value="Hàn Quốc">Hàn Quốc</option>
                    <option value="Singapor">Singapor</option>
                </select>

                <label for="">Địa chỉ</label>
                <select name="diachi" style="margin-top: 6px; margin-bottom: 20px;" value="<?php echo $dataC[6] ?>" class="form-select" aria-label="Default select example">
                    <option selected>Địa chỉ</option>
                    <option value="Hà Nội">Hà Nội</option>
                    <option value="Thanh Hóa">Thanh Hóa</option>
                    <option value="Hải Dương">Hải Dương</option>
                </select>

                <label for="">Giới tính</label>
                <div style="margin-top: 6px; margin-bottom: 20px;" class="check">
                    <?php
                    if ($dataC[7] == "Nam") {
                    ?>
                        <input class="form-check-input" type="radio" name="gioitinh" checked value="Nam" id="flexRadioDefault1">
                        <label class="form-check-label" for="flexRadioDefault1">
                            Nam
                        </label>

                        <input class="form-check-input" type="radio" name="gioitinh" value="Nữ" id="flexRadioDefault1">
                        <label class="form-check-label" for="flexRadioDefault1">
                            Nữ
                        </label>
                    <?php } else {
                    ?>
                        <input class="form-check-input" type="radio" name="gioitinh" value="Nam" id="flexRadioDefault1">
                        <label class="form-check-label" for="flexRadioDefault1">
                            Nam
                        </label>

                        <input class="form-check-input" type="radio" name="gioitinh" checked value="Nữ" id="flexRadioDefault1">
                        <label class="form-check-label" for="flexRadioDefault1">
                            Nữ
                        </label>
                    <?php } ?>
                </div>

            </div>
            <div class="modal-footer">
                <a href="../../../../learn_mvc/quanLyKhachHang" type="submit" class="btn btn-secondary" style="margin-right: 20px;">Đóng</a>
                <button type="submit" name="btnsuakh" class="btn btn-primary">Lưu✔✔</button>
            </div>
        </div>
    </div>
</form>