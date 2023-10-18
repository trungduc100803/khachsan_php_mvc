<?php
$room = $data['room'];
$hoadon = $data['hoaDonModel'];
$services = $data['dataService'];
$sldichvu = 0;
?>

<?php
if ($room[2] == "Đang trống") {

?>
    <form action="" id="formmain" method="post">
        <div class="chitietphong" style="display: flex; margin-top: 50px; ">
            <div class="chitiet_left" style="width: 50%;">
                <h4>Thông tin chi tiết của phòng</h4>
                <div class="phong_content" style="margin-top: 30px;">
                    <div class="phong_item" style="display: flex; align-items: stretch;">
                        <span style="margin-right: 10px; font-weight: 600;">Tên phòng:</span>
                        <p><?php echo "P" . $room[0]; ?></p>
                    </div>
                    <div class="phong_item" style="display: flex; align-items: stretch;">
                        <span style="margin-right: 10px; font-weight: 600;">Loại phòng:</span>
                        <p><?php echo  $room[1]; ?></p>
                    </div>
                    <div class="phong_item" style="display: flex; align-items: stretch;">
                        <span style="margin-right: 10px; font-weight: 600;">Trạng thái hiện tại:</span>
                        <p><?php echo $room[2]; ?></p>
                    </div>
                    <button type="button" data-bs-toggle="modal" data-bs-target="#mdatphong" class="btn btn-success">Đặt phòng</button>
                    <a href="../../../../learn_mvc/ChiTietPhongForNhanVien/sua/<?php echo $room[0] ?>" type="submit" name="suaphong" class="btn btn-warning">Sửa phòng</a>

                </div>
            </div>
            <div class="chitiet_right" style="width: 50%;">
                <h4>Thông tin dịch vụ</h4>
                <form id="formright" action="" method="POST">
                    <table style="width: 100%; border: 1px solid #333; padding: 10px; border-radius: 10px; margin-top: 30px;">
                        <thead>
                            <th style="padding: 10px;">Tên dịch vụ</th>
                            <th style="padding: 10px;">Số lượng</th>
                            <th style="padding: 10px;">Giá</th>
                            <th style="padding: 10px;">Thành tiền</th>
                        </thead>
                        <tbody>
                            <?php
                            $i = 0;
                            while ($service = mysqli_fetch_array($services)) {
                                $lengthSevice = mysqli_num_rows($services);
                                $i++;
                            ?>
                                <tr>
                                    <td style="padding: 10px; padding-left: 10px; "><?php echo $service['tendichvu'] ?></td>
                                    <td style="padding: 10px;">
                                        <input hidden type="text" name="lengthService" value="<?php echo $lengthSevice ?>" id="">
                                        <input hidden type="text" name="<?php echo "tendichvu" . $i  ?>" value="<?php echo $service['tendichvu'] ?>" id="">
                                        <input hidden type="text" name="<?php echo "gia" . $i  ?>" value="<?php echo $service['gia'] ?>" id="">
                                        <button id="btntru" type="button" style="font-size: 14px; padding:0px; width: 24px; height: 24px; border-radius: 100px; margin-left: 4px;" class="btn btn-plus btn-danger">-</button>
                                        <input class="input" style="width: 40px; border: none;" type="text" name="aa" id="<?php echo "ip" ?>" value="0">
                                        <button id="btncong" type="button" style="font-size: 14px; padding:0px; width: 24px; height: 24px; border-radius: 100px; margin-left: 4px;" class="btn btn-plus btn-success">+</button>
                                    </td>
                                    <td id="<?php echo "gia"  ?>" style="padding: 10px ; padding-right: 10px;"><?php echo $service['gia'] ?></td>
                                    <td class="<?php echo "tt" . $service[0] ?>" style="padding: 10px ; padding-right: 10px;">
                                        <input type="number" style="width: 70px; border: none;" name="<?php echo "thanhtien" . $i  ?>" value="0" id="<?php echo "ipT" ?>">
                                    </td>
                                </tr>
                            <?php
                            }
                            ?>
                        </tbody>
                    </table>
                </form>

            </div>

            <script>
                const AllBtnCong = document.querySelectorAll('#btncong')
                const AllBtnTru = document.querySelectorAll('#btntru')
                const tongtien = document.querySelector('.tongtien')
                let tong = 0

                const handleTru = (ip) => {
                    let val = ip.value

                    if (ip.value <= 0) {
                        ip.value = 0
                    } else {
                        val--
                        let g = parseInt(gia.innerHTML) * val
                        ip.value = val
                        ipT.value = g
                        tong = g
                    }
                }


                const handleCong = (i, ip, gia, ipT) => {
                    let val = ip.value
                    val++
                    let g = parseInt(gia.innerHTML) * val
                    ip.value = val
                    ipT.value = g
                }



                AllBtnCong.forEach((btnCong, i) => {
                    btnCong.addEventListener('click', () => {
                        const ip = btnCong.parentElement.querySelector('.input')
                        const gia = btnCong.parentElement.parentElement.querySelector('#gia')
                        const ipT = btnCong.parentElement.parentElement.querySelector('#ipT')
                        handleCong(i, ip, gia, ipT)
                    })
                })

                AllBtnTru.forEach((btnTru, i) => {
                    btnTru.addEventListener('click', () => {
                        const ip = btnTru.parentElement.querySelector('input')
                        const gia = btnTru.parentElement.parentElement.querySelector('#gia')
                        const ipT = btnTru.parentElement.parentElement.querySelector('#ipT')
                        handleTru(i, ip, gia, ipT)
                    })
                })
            </script>

            <!-- modal  dat phong -->
            <form method="POST" class="modal fade" id="mdatphong" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
                <div class="modal-dialog">
                    <div class="modal-content">
                        <div class="modal-header">
                            <h1 class="modal-title fs-5" id="exampleModalLabel">Đặt phòng</h1>
                            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                        </div>
                        <div class="modal-body">
                            <label for="">Loại khách hàng</label>
                            <div style="margin-top: 6px; margin-bottom: 20px;" class="check">
                                <input onclick="handleChangeForm()" class="form-check-input input_loaikh" type="radio" name="loaikhachhang" value="Khách hàng mới" id="flexRadioDefault1">
                                <label style=" margin-right: 20px;" class="form-check-label" for="flexRadioDefault1">
                                    Khách hàng mới
                                </label>

                                <input onclick="handleChangeForm()" class="form-check-input input_loaikh1" type="radio" name="loaikhachhang" value="Khách hàng cũ" id="flexRadioDefault1">
                                <label class="form-check-label" for="flexRadioDefault1">
                                    Khách hàng cũ
                                </label>
                            </div>

                            <div class="main_input">
                                <!-- import input -->
                            </div>
                        </div>
                        <div class="modal-footer">
                            <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Đóng</button>
                            <button name="datphong" type="submit" class="btn btn-primary">Lưu</button>
                        </div>
                    </div>
                </div>
            </form>

            <script>
                const input_loaikh = document.querySelector('.input_loaikh')
                const input_loaikh1 = document.querySelector('.input_loaikh1')
                const modal_body = document.querySelector('.main_input')

                const handleChangeForm = () => {
                    if (input_loaikh.checked) {
                        const html = `
                            <label for="">Tên khách hàng</label>
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
                            </select>
    
                            <label for="">Giới tính</label>
                            <div style="margin-top: 6px; margin-bottom: 20px;" class="check">
                                <input class="form-check-input" type="radio" name="gioitinh" value="Nam" id="flexRadioDefault1">
                                <label class="form-check-label" for="flexRadioDefault1">
                                    Nam
                                </label>
    
                                <input class="form-check-input" type="radio" name="gioitinh" value="Nữ" id="flexRadioDefault1">
                                <label class="form-check-label" for="flexRadioDefault1">
                                    Nữ
                                </label>
                            </div>
    
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
                            <input name="tiencoc" value="20000" style="margin-top: 6px; margin-bottom: 20px;" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">`
                        modal_body.innerHTML = html
                    } else if (input_loaikh1.checked) {
                        const html = `
                            <label for="">Tên khách hàng</label>
                            <input name="tenkhachhang" style="margin-top: 6px; margin-bottom: 20px;" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">
                            
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
                            <input name="tiencoc" value="20000" style="margin-top: 6px; margin-bottom: 20px;" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">`
                        modal_body.innerHTML = html
                    }
                }
            </script>
        </div>
    </form>
<?php } else if ($room[2] == "Đã đặt") {
    $datangaydat = $hoadon->getAHoaDon($room[3], $room[0]);
    $numrow = mysqli_num_rows($datangaydat);
    $datangaydat->data_seek($numrow - 1);
    $ngaydat = $datangaydat->fetch_row(); ?>

    <form action="" id="formmain" method="post">
        <div class="chitietphong" style="display: flex; margin-top: 50px; ">
            <div class="chitiet_left" style="width: 50%;">
                <h4>Thông tin chi tiết của phòng</h4>
                <div class="phong_content" style="margin-top: 30px;">
                    <div class="phong_item" style="display: flex; align-items: stretch;">
                        <input value="<?php echo $room[3]; ?>" hidden type="text" name="kh" id="">
                        <input value="<?php echo  $ngaydat[4]; ?>" hidden type="text" name="ngayd" id="">
                        <span style="margin-right: 10px; font-weight: 600;">Tên phòng:</span>
                        <p><?php echo "P" . $room[0]; ?></p>
                    </div>
                    <div class="phong_item" style="display: flex; align-items: stretch;">
                        <span style="margin-right: 10px; font-weight: 600;">Loại phòng:</span>
                        <p><?php echo  $room[1]; ?></p>
                    </div>
                    <div class="phong_item" style="display: flex; align-items: stretch;">
                        <span style="margin-right: 10px; font-weight: 600;">Trạng thái hiện tại:</span>
                        <p><?php echo $room[2]; ?></p>
                    </div>

                    <div class="phong_item" style="display: flex; align-items: stretch;">
                        <span style="margin-right: 10px; font-weight: 600;">Ngày đặt:</span>
                        <p><?php echo $ngaydat[4]; ?></p>
                    </div>
                    <div class="phong_item" style="display: flex; align-items: stretch;">
                        <span style="margin-right: 10px; font-weight: 600;">Khách hàng:</span>
                        <p><?php echo $room[3]; ?></p>
                    </div>

                    <div class="phong_item" style="display: flex; align-items: stretch;">
                        <span style="margin-right: 10px; font-weight: 600;">Giá:</span>
                        <p><?php echo $room[4]; ?></p>
                    </div>
                    <button type="button" class="btn btn-success" data-bs-toggle="modal" data-bs-target="#exampleModalNhan">Nhận phòng</button>
                    <button type="submit" class="btn btn-info" name="datdichvu" data-bs-toggle="modal" data-bs-target="#exampleModalDV">Đặt dịch vụ</button>
                    <a href="../../../../learn_mvc/ChiTietPhongForNhanVien/sua/<?php echo $room[0] ?>" type="submit" name="suaphong" class="btn btn-warning">Sửa phòng</a>
                </div>
            </div>
            <div class="chitiet_right" style="width: 50%;">
                <h4>Thông tin dịch vụ</h4>
                <form id="formright" action="" method="POST">
                    <table style="width: 100%; border: 1px solid #333; padding: 10px; border-radius: 10px; margin-top: 30px;">
                        <thead>
                            <th style="padding: 10px;">Tên dịch vụ</th>
                            <th style="padding: 10px;">Số lượng</th>
                            <th style="padding: 10px;">Giá</th>
                            <th style="padding: 10px;">Thành tiền</th>
                        </thead>
                        <tbody>
                            <?php
                            $i = 0;
                            while ($service = mysqli_fetch_array($services)) {
                                $lengthSevice = mysqli_num_rows($services);
                                $i++;
                            ?>
                                <tr>
                                    <td style="padding: 10px; padding-left: 10px; "><?php echo $service['tendichvu'] ?></td>
                                    <td style="padding: 10px;">
                                        <input hidden type="text" name="lengthService" value="<?php echo $lengthSevice ?>" id="">
                                        <input hidden type="text" name="<?php echo "tendichvu" . $i  ?>" value="<?php echo $service['tendichvu'] ?>" id="">
                                        <input hidden type="text" name="<?php echo "gia" . $i  ?>" value="<?php echo $service['gia'] ?>" id="">
                                        <button id="btntru" type="button" style="font-size: 14px; padding:0px; width: 24px; height: 24px; border-radius: 100px; margin-left: 4px;" class="btn btn-plus btn-danger">-</button>
                                        <input class="input" style="width: 40px; border: none;" type="text" name="aa" id="<?php echo "ip" ?>" value="0">
                                        <button id="btncong" type="button" style="font-size: 14px; padding:0px; width: 24px; height: 24px; border-radius: 100px; margin-left: 4px;" class="btn btn-plus btn-success">+</button>
                                    </td>
                                    <td id="<?php echo "gia"  ?>" style="padding: 10px ; padding-right: 10px;"><?php echo $service['gia'] ?></td>
                                    <td class="<?php echo "tt" . $service[0] ?>" style="padding: 10px ; padding-right: 10px;">
                                        <input type="number" style="width: 70px; border: none;" name="<?php echo "thanhtien" . $i  ?>" value="0" id="<?php echo "ipT" ?>">
                                    </td>
                                </tr>
                            <?php
                            }
                            ?>
                        </tbody>
                    </table>
                </form>

            </div>

            <script>
                const AllBtnCong1 = document.querySelectorAll('#btncong')
                const AllBtnTru1 = document.querySelectorAll('#btntru')
                const tongtien2 = document.querySelector('.tongtien')
                let tong1 = 0

                const handleTru1 = (ip) => {
                    let val = ip.value

                    if (ip.value <= 0) {
                        ip.value = 0
                    } else {
                        val--
                        let g = parseInt(gia.innerHTML) * val
                        ip.value = val
                        ipT.value = g
                        tong1 = g
                    }
                }


                const handleCong1 = (i, ip, gia, ipT) => {
                    let val = ip.value
                    val++
                    let g = parseInt(gia.innerHTML) * val
                    ip.value = val
                    ipT.value = g
                }



                AllBtnCong1.forEach((btnCong, i) => {
                    btnCong.addEventListener('click', () => {
                        const ip = btnCong.parentElement.querySelector('.input')
                        const gia = btnCong.parentElement.parentElement.querySelector('#gia')
                        const ipT = btnCong.parentElement.parentElement.querySelector('#ipT')
                        handleCong1(i, ip, gia, ipT)
                    })
                })

                AllBtnTru1.forEach((btnTru, i) => {
                    btnTru.addEventListener('click', () => {
                        const ip = btnTru.parentElement.querySelector('input')
                        const gia = btnTru.parentElement.parentElement.querySelector('#gia')
                        const ipT = btnTru.parentElement.parentElement.querySelector('#ipT')
                        handleTru1(i, ip, gia, ipT)
                    })
                })
            </script>

            <!-- Modal nhan phong-->
            <form method="post" class="modal fade" id="exampleModalNhan" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
                <div class="modal-dialog">
                    <div class="modal-content">
                        <div class="modal-header">
                            <h1 class="modal-title fs-5" id="exampleModalLabel">Nhận phòng</h1>
                            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                        </div>
                        <div class="modal-body">
                            <h5>❤❤Bạn muốn nhận phòng❤❤</h5>
                        </div>
                        <div class="modal-footer">
                            <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Đóng</button>
                            <button type="submit" name="btnnhanphong" class="btn btn-primary">OK</button>
                        </div>
                    </div>
                </div>
            </form>

            <!-- Modal tra phong-->
            <form method="post" class="modal fade" id="exampleModalTra" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
                <div class="modal-dialog">
                    <div class="modal-content">
                        <div class="modal-header">
                            <h1 class="modal-title fs-5" id="exampleModalLabel">Trả phòng</h1>
                            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                        </div>
                        <div class="modal-body">
                            <div class="body_item" style="display: flex; align-items: center;   margin-bottom: 12px;">
                                <span style="font-weight: 600; margin-right: 6px;">Tên phòng:</span>
                                <span><?php echo $room[0]; ?></span>
                            </div>
                            <div class="body_item" style="display: flex; align-items: center;   margin-bottom: 12px;">
                                <span style="font-weight: 600; margin-right: 6px;">Loại phòng:</span>
                                <span><?php echo $room[1]; ?></span>
                            </div>
                            <div class="body_item" style="display: flex; align-items: center;   margin-bottom: 12px;">
                                <span style="font-weight: 600; margin-right: 6px;">Ngày đặt:</span>
                                <span><?php echo $ngaydat[4]; ?></span>
                            </div>
                            <input value="<?php echo $room[3]; ?>" hidden type="text" name="kh" id="">
                            <input value="<?php echo  $ngaydat[4]; ?>" hidden type="text" name="ngayd" id="">
                            <div class="body_item" style="display: flex; align-items: center;   margin-bottom: 12px;">
                                <span style="font-weight: 600; margin-right: 6px;">Ngày nhận:</span>
                                <span><?php echo $ngaydat[5]; ?></span>
                            </div>
                            <div class="body_item" style="display: flex; align-items: center;   margin-bottom: 12px;">
                                <span style="font-weight: 600; margin-right: 6px;">Giá phòng:</span>
                                <span><?php echo $room[4]; ?></span>
                            </div>

                            <h5 style="padding-top: 24ps; margin-bottom: 20px;">Thông tin các dịch vu</h5>
                            <?php if (mysqli_num_rows($data['dataDichvu']) > 0) {

                            ?>
                                <form style="padding-top: 50px; margin-bottom: 30px;" action="" method="post">
                                    <table style="width: 100%; border: 1px solid #333; padding: 10px; border-radius: 10px; margin-top: 30px;">
                                        <thead>
                                            <th style="padding: 10px;">Tên dịch vụ</th>
                                            <th style="padding: 10px;">Số lượng</th>
                                            <th style="padding: 10px;">Giá</th>
                                            <th style="padding: 10px;">Thành tiền</th>
                                        </thead>
                                        <tbody>
                                            <?php
                                            $tong = 0;
                                            while ($service = mysqli_fetch_array($data['dataDichvu'])) {
                                            ?>
                                                <tr>
                                                    <td style="padding: 10px;"><?php echo $service['tendichvu'] ?></td>
                                                    <td style="padding: 10px;"><?php echo $service['soluong'] ?></td>
                                                    <td style="padding: 10px;"><?php echo $service['gia'] ?></td>
                                                    <td style="padding: 10px;"><?php echo $service['thanhtien'] ?></td>
                                                    <td hidden style="padding: 10px;"><?php $tong += $service['thanhtien']; ?></td>
                                                </tr>
                                            <?php } ?>
                                        </tbody>
                                    </table>
                                </form>
                                <input type="text" name="tongthanhtoan" value="<?php echo $room[4] + $tong; ?>" hidden id="">
                                <input type="text" name="tenkh" value="<?php echo $room[3]; ?>" hidden id="">
                                <input type="text" name="nd" value="<?php echo $ngaydat[4]; ?>" hidden id="">
                                <h4>Tổng thanh toán: <?php echo $room[4] + $tong; ?> </h4>
                            <?php } else {

                            ?>
                                <input type="text" name="tenkh" value="<?php echo $room[3]; ?>" hidden id="">
                                <input type="text" name="nd" value="<?php echo $ngaydat[4]; ?>" hidden id="">
                                <input type="text" name="tongthanhtoan" value="<?php echo $room[4]; ?>" hidden id="">
                                <h4>Tổng thanh toán: <?php echo $room[4]; ?> </h4>

                            <?php } ?>
                        </div>
                        <div class="modal-footer">
                            <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Đóng</button>
                            <button type="submit" name="btntraphong" class="btn btn-primary">Lưu</button>
                        </div>
                    </div>
                </div>
            </form>

            <!-- Modal dat dich vu-->
            <form method="post" class="modal fade" id="exampleModalDV" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
                <div class="modal-dialog">
                    <div class="modal-content">
                        <div class="modal-header">
                            <h1 class="modal-title fs-5" id="exampleModalLabel">Đặt dịch vụ</h1>
                            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                        </div>
                        <div class="modal-body">
                            <label for="">Chọn dịch vụ</label>
                            <table style="width: 100%; border: 1px solid #333; padding: 10px; border-radius: 10px; margin-top: 30px;">
                                <thead>
                                    <th style="padding: 10px;">Tên dịch vụ</th>
                                    <th style="padding: 10px;">Số lượng</th>
                                    <th style="padding: 10px;">Thành tiền</th>
                                </thead>
                                <tbody>
                                    <?php
                                    $i = 0;
                                    while ($service = mysqli_fetch_array($services)) {
                                        $i++;
                                    ?>
                                        <tr>
                                            <td style="padding: 10px; padding-left: 10px;"><?php echo $service['tendichvu'] ?></td>
                                            <td style="padding: 10px;">
                                                <button type="button" style="font-size: 14px; padding:0px; width: 24px; height: 24px; border-radius: 100px; margin-right: 4px;" class="btn btn-less btn-success" name="btnbo">-</button>
                                                <span class="sl"></span>
                                                <button type="button" style="font-size: 14px; padding:0px; width: 24px; height: 24px; border-radius: 100px; margin-left: 4px;" class="btn btn-plus btn-danger" name="btnthem">+</button>
                                            </td>
                                            <td style="padding: 10px ; padding-right: 10px;"><?php echo $service['gia'] ?></td>
                                        </tr>
                                    <?php
                                    }
                                    ?>
                                </tbody>
                            </table>
                        </div>
                        <div class="modal-footer">
                            <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Đóng</button>
                            <button type="submit" name="btntraphong" class="btn btn-primary">Lưu</button>
                        </div>
                    </div>
                </div>
            </form>
        </div>
        </div>
    </form>
<?php } else {
    $datangaydat = $hoadon->getAHoaDon($room[3], $room[0]);
    $numrow = mysqli_num_rows($datangaydat);
    $datangaydat->data_seek($numrow - 1);
    $ngaydat = $datangaydat->fetch_row(); ?>

    <form action="" id="formmain" method="post">
        <div class="chitietphong" style="display: flex; margin-top: 50px; ">
            <div class="chitiet_left" style="width: 50%;">
                <h4>Thông tin chi tiết của phòng</h4>
                <div class="phong_content" style="margin-top: 30px;">
                    <div class="phong_item" style="display: flex; align-items: stretch;">
                        <input value="<?php echo $room[3]; ?>" hidden type="text" name="kh" id="">
                        <input value="<?php echo  $ngaydat[4]; ?>" hidden type="text" name="ngayd" id="">
                        <span style="margin-right: 10px; font-weight: 600;">Tên phòng:</span>
                        <p><?php echo "P" . $room[0]; ?></p>
                    </div>
                    <div class="phong_item" style="display: flex; align-items: stretch;">
                        <span style="margin-right: 10px; font-weight: 600;">Loại phòng:</span>
                        <p><?php echo  $room[1]; ?></p>
                    </div>
                    <div class="phong_item" style="display: flex; align-items: stretch;">
                        <span style="margin-right: 10px; font-weight: 600;">Trạng thái hiện tại:</span>
                        <p><?php echo $room[2]; ?></p>
                    </div>

                    <div class="phong_item" style="display: flex; align-items: stretch;">
                        <span style="margin-right: 10px; font-weight: 600;">Ngày đặt:</span>
                        <p><?php echo $ngaydat[4]; ?></p>
                    </div>
                    <div class="phong_item" style="display: flex; align-items: stretch;">
                        <span style="margin-right: 10px; font-weight: 600;">Ngày nhận:</span>
                        <p><?php echo $ngaydat[5]; ?></p>
                    </div>
                    <div class="phong_item" style="display: flex; align-items: stretch;">
                        <span style="margin-right: 10px; font-weight: 600;">Khách hàng:</span>
                        <p><?php echo $room[3]; ?></p>
                    </div>

                    <div class="phong_item" style="display: flex; align-items: stretch;">
                        <span style="margin-right: 10px; font-weight: 600;">Giá:</span>
                        <p><?php echo $room[4]; ?></p>
                    </div>
                    <button type="button" class="btn btn-primary" data-bs-toggle="modal" data-bs-target="#exampleModalTra">Trả phòng</button>
                    <button type="submit" name="datdichvu" class="btn btn-info" data-bs-toggle="modal" data-bs-target="#exampleModalDV">Đặt dịch vụ</button>
                    <a href="../../../../learn_mvc/ChiTietPhongForNhanVien/sua/<?php echo $room[0] ?>" class="btn btn-warning" type="submit" name="suaphong">Sửa phòng</a>
                </div>
            </div>
            <div class="chitiet_right" style="width: 50%;">
                <h4>Thông tin dịch vụ</h4>
                <form id="formright" action="" method="POST">
                    <table style="width: 100%; border: 1px solid #333; padding: 10px; border-radius: 10px; margin-top: 30px;">
                        <thead>
                            <th style="padding: 10px;">Tên dịch vụ</th>
                            <th style="padding: 10px;">Số lượng</th>
                            <th style="padding: 10px;">Giá</th>
                            <th style="padding: 10px;">Thành tiền</th>
                        </thead>
                        <tbody>
                            <?php
                            $i = 0;
                            while ($service = mysqli_fetch_array($services)) {
                                $lengthSevice = mysqli_num_rows($services);
                                $i++;
                            ?>
                                <tr>
                                    <td style="padding: 10px; padding-left: 10px; "><?php echo $service['tendichvu'] ?></td>
                                    <td style="padding: 10px;">
                                        <input hidden type="text" name="lengthService" value="<?php echo $lengthSevice ?>" id="">
                                        <input hidden type="text" name="<?php echo "tendichvu" . $i  ?>" value="<?php echo $service['tendichvu'] ?>" id="">
                                        <input hidden type="text" name="<?php echo "gia" . $i  ?>" value="<?php echo $service['gia'] ?>" id="">
                                        <button id="btntru" type="button" style="font-size: 14px; padding:0px; width: 24px; height: 24px; border-radius: 100px; margin-left: 4px;" class="btn btn-plus btn-danger">-</button>
                                        <input class="input" style="width: 40px; border: none;" type="text" name="aa" id="<?php echo "ip" ?>" value="0">
                                        <button id="btncong" type="button" style="font-size: 14px; padding:0px; width: 24px; height: 24px; border-radius: 100px; margin-left: 4px;" class="btn btn-plus btn-success">+</button>
                                    </td>
                                    <td id="<?php echo "gia"  ?>" style="padding: 10px ; padding-right: 10px;"><?php echo $service['gia'] ?></td>
                                    <td class="<?php echo "tt" . $service[0] ?>" style="padding: 10px ; padding-right: 10px;">
                                        <input type="number" style="width: 70px; border: none;" name="<?php echo "thanhtien" . $i  ?>" value="0" id="<?php echo "ipT" ?>">
                                    </td>
                                </tr>
                            <?php
                            }
                            ?>
                        </tbody>
                    </table>
                </form>

            </div>

            <script>
                const AllBtnCong2 = document.querySelectorAll('#btncong')
                const AllBtnTru13 = document.querySelectorAll('#btntru')
                const tongtien3 = document.querySelector('.tongtien')
                let tong3 = 0

                const handleTru3 = (ip) => {
                    let val = ip.value

                    if (ip.value <= 0) {
                        ip.value = 0
                    } else {
                        val--
                        let g = parseInt(gia.innerHTML) * val
                        ip.value = val
                        ipT.value = g
                        tong3 = g
                    }
                }


                const handleCong3 = (i, ip, gia, ipT) => {
                    let val = ip.value
                    val++
                    let g = parseInt(gia.innerHTML) * val
                    ip.value = val
                    ipT.value = g
                }



                AllBtnCong2.forEach((btnCong, i) => {
                    btnCong.addEventListener('click', () => {
                        const ip = btnCong.parentElement.querySelector('.input')
                        const gia = btnCong.parentElement.parentElement.querySelector('#gia')
                        const ipT = btnCong.parentElement.parentElement.querySelector('#ipT')
                        handleCong3(i, ip, gia, ipT)
                    })
                })

                AllBtnTru13.forEach((btnTru, i) => {
                    btnTru.addEventListener('click', () => {
                        const ip = btnTru.parentElement.querySelector('input')
                        const gia = btnTru.parentElement.parentElement.querySelector('#gia')
                        const ipT = btnTru.parentElement.parentElement.querySelector('#ipT')
                        handleTru3(i, ip, gia, ipT)
                    })
                })
            </script>

            <!-- Modal nhan phong-->
            <form method="post" class="modal fade" id="exampleModalNhan" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
                <div class="modal-dialog">
                    <div class="modal-content">
                        <div class="modal-header">
                            <h1 class="modal-title fs-5" id="exampleModalLabel">Nhận phòng</h1>
                            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                        </div>
                        <div class="modal-body">
                            <h5>❤❤Bạn muốn nhận phòng❤❤</h5>
                        </div>
                        <div class="modal-footer">
                            <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Đóng</button>
                            <button type="submit" name="btnnhanphong" class="btn btn-primary">OK</button>
                        </div>
                    </div>
                </div>
            </form>

            <!-- Modal tra phong-->
            <form method="post" class="modal fade" id="exampleModalTra" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
                <div class="modal-dialog">
                    <div class="modal-content">
                        <div class="modal-header">
                            <h1 class="modal-title fs-5" id="exampleModalLabel">Trả phòng</h1>
                            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                        </div>
                        <div class="modal-body">
                            <div class="body_item" style="display: flex; align-items: center;   margin-bottom: 12px;">
                                <span style="font-weight: 600; margin-right: 6px;">Tên phòng:</span>
                                <span><?php echo $room[0]; ?></span>
                            </div>
                            <div class="body_item" style="display: flex; align-items: center;   margin-bottom: 12px;">
                                <span style="font-weight: 600; margin-right: 6px;">Loại phòng:</span>
                                <span><?php echo $room[1]; ?></span>
                            </div>
                            <div class="body_item" style="display: flex; align-items: center;   margin-bottom: 12px;">
                                <span style="font-weight: 600; margin-right: 6px;">Ngày đặt:</span>
                                <span><?php echo $ngaydat[4]; ?></span>
                            </div>
                            <input value="<?php echo $room[3]; ?>" hidden type="text" name="kh" id="">
                            <input value="<?php echo  $ngaydat[4]; ?>" hidden type="text" name="ngayd" id="">
                            <div class="body_item" style="display: flex; align-items: center;   margin-bottom: 12px;">
                                <span style="font-weight: 600; margin-right: 6px;">Ngày nhận:</span>
                                <span><?php echo $ngaydat[5]; ?></span>
                            </div>
                            <div class="body_item" style="display: flex; align-items: center;   margin-bottom: 12px;">
                                <span style="font-weight: 600; margin-right: 6px;">Giá phòng:</span>
                                <span><?php echo $room[4]; ?></span>
                            </div>

                            <h5 style="padding-top: 24ps; margin-bottom: 20px;">Thông tin các dịch vu</h5>
                            <?php if (mysqli_num_rows($data['dataDichvu']) > 0) {

                            ?>
                                <form style="padding-top: 50px; margin-bottom: 30px;" action="" method="post">
                                    <table style="width: 100%; border: 1px solid #333; padding: 10px; border-radius: 10px; margin-top: 30px;">
                                        <thead>
                                            <th style="padding: 10px;">Tên dịch vụ</th>
                                            <th style="padding: 10px;">Số lượng</th>
                                            <th style="padding: 10px;">Giá</th>
                                            <th style="padding: 10px;">Thành tiền</th>
                                        </thead>
                                        <tbody>
                                            <?php
                                            $tong = 0;
                                            while ($service = mysqli_fetch_array($data['dataDichvu'])) {
                                            ?>
                                                <tr>
                                                    <td style="padding: 10px;"><?php echo $service['tendichvu'] ?></td>
                                                    <td style="padding: 10px;"><?php echo $service['soluong'] ?></td>
                                                    <td style="padding: 10px;"><?php echo $service['gia'] ?></td>
                                                    <td style="padding: 10px;"><?php echo $service['thanhtien'] ?></td>
                                                    <td hidden style="padding: 10px;"><?php $tong += $service['thanhtien']; ?></td>
                                                </tr>
                                            <?php } ?>
                                        </tbody>
                                    </table>
                                </form>
                                <input type="text" name="tongthanhtoan" value="<?php echo $room[4] + $tong; ?>" hidden id="">
                                <input type="text" name="tenkh" value="<?php echo $room[3]; ?>" hidden id="">
                                <input type="text" name="nd" value="<?php echo $ngaydat[4]; ?>" hidden id="">
                                <h4>Tổng thanh toán: <?php echo $room[4] + $tong; ?> </h4>
                            <?php } else {

                            ?>
                                <input type="text" name="tenkh" value="<?php echo $room[3]; ?>" hidden id="">
                                <input type="text" name="nd" value="<?php echo $ngaydat[4]; ?>" hidden id="">
                                <input type="text" name="tongthanhtoan" value="<?php echo $room[4]; ?>" hidden id="">
                                <h4>Tổng thanh toán: <?php echo $room[4]; ?> </h4>

                            <?php } ?>
                        </div>
                        <div class="modal-footer">
                            <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Đóng</button>
                            <button type="submit" name="btntraphong" class="btn btn-primary">Lưu</button>
                        </div>
                    </div>
                </div>
            </form>

            <!-- Modal dat dich vu-->
            <form method="post" class="modal fade" id="exampleModalDV" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
                <div class="modal-dialog">
                    <div class="modal-content">
                        <div class="modal-header">
                            <h1 class="modal-title fs-5" id="exampleModalLabel">Đặt dịch vụ</h1>
                            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                        </div>
                        <div class="modal-body">
                            <label for="">Chọn dịch vụ</label>
                            <table style="width: 100%; border: 1px solid #333; padding: 10px; border-radius: 10px; margin-top: 30px;">
                                <thead>
                                    <th style="padding: 10px;">Tên dịch vụ</th>
                                    <th style="padding: 10px;">Số lượng</th>
                                    <th style="padding: 10px;">Thành tiền</th>
                                </thead>
                                <tbody>
                                    <?php
                                    $i = 0;
                                    while ($service = mysqli_fetch_array($services)) {
                                        $i++;
                                    ?>
                                        <tr>
                                            <td style="padding: 10px; padding-left: 10px;"><?php echo $service['tendichvu'] ?></td>
                                            <td style="padding: 10px;">
                                                <button type="button" style="font-size: 14px; padding:0px; width: 24px; height: 24px; border-radius: 100px; margin-right: 4px;" class="btn btn-less btn-success" name="btnbo">-</button>
                                                <span class="sl"></span>
                                                <button type="button" style="font-size: 14px; padding:0px; width: 24px; height: 24px; border-radius: 100px; margin-left: 4px;" class="btn btn-plus btn-danger" name="btnthem">+</button>
                                            </td>
                                            <td style="padding: 10px ; padding-right: 10px;"><?php echo $service['gia'] ?></td>
                                        </tr>
                                    <?php
                                    }
                                    ?>
                                </tbody>
                            </table>
                        </div>
                        <div class="modal-footer">
                            <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Đóng</button>
                            <button type="submit" name="btntraphong" class="btn btn-primary">Lưu</button>
                        </div>
                    </div>
                </div>
            </form>
        </div>
        </div>
    </form>
<?php } ?>