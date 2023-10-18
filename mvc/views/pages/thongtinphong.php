<?php
$rooms = $data['dataRooms'];
$hoaDonModel = $data['hoaDonModel'];
// session_start();

?>

<?php
if (mysqli_num_rows($rooms) > 0) {
    $room = mysqli_fetch_assoc($rooms);
    $datangaydat = $hoaDonModel->getAHoaDon($room['name'], $room['roomID']);
    $numrow = mysqli_num_rows($datangaydat);
    $datangaydat->data_seek($numrow - 1);
    $ngaydat = $datangaydat->fetch_row();

    $_SESSION['ngaydat'] = $ngaydat[4];
    $_SESSION['ngaynhan'] = $ngaydat[5];
    $_SESSION['roomID'] = $room['roomID'];
    if ($room['trangthaiphong'] == "Đã đặt") {

?>
        <form action="" method="post" style="display: flex;">
            <div class="thongtinphong" style="padding-top: 30px; width: 50%;">
                <input type="text" name="roomid" id="" value="<?php echo $room['roomID']; ?>" hidden>

                <h4 style="margin-bottom: 20px;">Thông tin phòng đang đặt</h4>

                <div class="phong_item" style="display: flex; align-items: stretch;">
                    <span style="margin-right: 10px; font-weight: 600;">Tên phòng:</span>
                    <p><?php echo "P" . $room['roomID']; ?></p>
                </div>
                <div class="phong_item" style="display: flex; align-items: stretch;">
                    <span style="margin-right: 10px; font-weight: 600;">Loại phòng:</span>
                    <p><?php echo  $room['loaiphong']; ?></p>
                </div>
                <div class="phong_item" style="display: flex; align-items: stretch;">
                    <input value="<?php echo $room['name']; ?>" hidden type="text" name="kh" id="">
                    <input value="<?php echo  $ngaydat[4]; ?>" hidden type="text" name="ngayd" id="">
                    <span style="margin-right: 10px; font-weight: 600;">Trạng thái hiện tại:</span>
                    <p><?php echo $room['trangthaiphong']; ?></p>
                </div>

                <div class="phong_item" style="display: flex; align-items: stretch;">
                    <span style="margin-right: 10px; font-weight: 600;">Ngày đặt:</span>
                    <p><?php echo $_SESSION['ngaydat']; ?></p>
                </div>
                <div class="phong_item" style="display: flex; align-items: stretch;">
                    <span style="margin-right: 10px; font-weight: 600;">Khách hàng:</span>
                    <p><?php echo $room['name']; ?></p>
                </div>

                <div class="phong_item" style="display: flex; align-items: stretch;">
                    <span style="margin-right: 10px; font-weight: 600;">Giá:</span>
                    <p><?php echo $room['gia']; ?></p>
                </div>
                <div class="phong_item" style="display: flex; align-items: stretch; margin-bottom: 30px;">
                    <span style="margin-right: 10px; font-weight: 600;">Tình trạng:</span>
                    <p>Đã dọn dẹp</p>
                </div>
                <button type="submit" name="nhanphong" class="btn btn-success">Nhận phòng</button>
                <button type="submit" name="datdichvu" class="btn btn-warning">Đặt dịch vụ</button>
            </div>

            <!-- dich vu -->
            <div class="chitiet_right" style="width: 50%;">
                <h4>Thông tin dịch vụ</h4>
                <form action="" method="POST">
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
                            while ($service = mysqli_fetch_array($data['dataService'])) {
                                $lengthSevice = mysqli_num_rows($data['dataService']);
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


        </form>
    <?php
    } else {
    ?>
        <form action="" method="post" style="display: flex;">

            <div class="thongtinphong" style="padding-top: 30px; width: 50%;">
                <h4 style="margin-bottom: 20px;">Thông tin phòng đang đặt</h4>
                <input type="text" name="roomid" id="" value="<?php echo  $room['roomID']; ?>" hidden>

                <div class="phong_item" style="display: flex; align-items: stretch;">
                    <span style="margin-right: 10px; font-weight: 600;">Tên phòng:</span>
                    <p><?php echo "P" . $room['roomID']; ?></p>
                </div>
                <div class="phong_item" style="display: flex; align-items: stretch;">
                    <span style="margin-right: 10px; font-weight: 600;">Loại phòng:</span>
                    <p><?php echo  $room['loaiphong']; ?></p>
                </div>
                <div class="phong_item" style="display: flex; align-items: stretch;">
                    <span style="margin-right: 10px; font-weight: 600;">Trạng thái hiện tại:</span>
                    <p><?php echo $room['trangthaiphong']; ?></p>
                </div>

                <div class="phong_item" style="display: flex; align-items: stretch;">
                    <span style="margin-right: 10px; font-weight: 600;">Ngày đặt:</span>
                    <p><?php echo $_SESSION['ngaydat']; ?></p>
                </div>
                <div class="phong_item" style="display: flex; align-items: stretch;">
                    <span style="margin-right: 10px; font-weight: 600;">Ngày nhận:</span>
                    <p><?php echo $_SESSION['ngaynhan']; ?></p>
                </div>
                <div class="phong_item" style="display: flex; align-items: stretch;">
                    <span style="margin-right: 10px; font-weight: 600;">Khách hàng:</span>
                    <p><?php echo $room['name']; ?></p>
                </div>

                <div class="phong_item" style="display: flex; align-items: stretch;">
                    <span style="margin-right: 10px; font-weight: 600;">Giá:</span>
                    <p><?php echo $room['gia']; ?></p>
                </div>
                <div class="phong_item" style="display: flex; align-items: stretch; margin-bottom: 30px;">
                    <span style="margin-right: 10px; font-weight: 600;">Tình trạng:</span>
                    <p>Đã dọn dẹp</p>
                </div>
                <input value="<?php echo $room['name']; ?>" hidden type="text" name="kh" id="">
                <input value="<?php echo  $ngaydat[4]; ?>" hidden type="text" name="ngayd" id="">
                <button type="button" name="traphong" data-bs-toggle="modal" data-bs-target="#exampleModaltraphong" class="btn btn-success">Trả phòng</button>
                <button type="submit" name="datdichvu" class="btn btn-warning">Đặt dịch vụ</button>
            </div>

            <!-- dich vu -->
            <div class="chitiet_right" style="width: 50%;">
                <h4>Thông tin dịch vụ</h4>
                <form action="" method="POST">
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
                            while ($service = mysqli_fetch_array($data['dataService'])) {
                                $lengthSevice = mysqli_num_rows($data['dataService']);
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
                const AllBtnCongs = document.querySelectorAll('#btncong')
                const AllBtnTrus = document.querySelectorAll('#btntru')
                const tongtiens = document.querySelector('.tongtien')
                let tongs = 0

                const handleTrus = (ip) => {
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


                const handleCongs = (i, ip, gia, ipT) => {
                    let val = ip.value
                    val++
                    let g = parseInt(gia.innerHTML) * val
                    ip.value = val
                    ipT.value = g
                }



                AllBtnCongs.forEach((btnCong, i) => {
                    btnCong.addEventListener('click', () => {
                        const ip = btnCong.parentElement.querySelector('.input')
                        const gia = btnCong.parentElement.parentElement.querySelector('#gia')
                        const ipT = btnCong.parentElement.parentElement.querySelector('#ipT')
                        handleCongs(i, ip, gia, ipT)
                    })
                })

                AllBtnTrus.forEach((btnTru, i) => {
                    btnTru.addEventListener('click', () => {
                        const ip = btnTru.parentElement.querySelector('input')
                        const gia = btnTru.parentElement.parentElement.querySelector('#gia')
                        const ipT = btnTru.parentElement.parentElement.querySelector('#ipT')
                        handleTrus(i, ip, gia, ipT)
                    })
                })
            </script>

        </form>
        <form method="post" class="modal fade" id="exampleModaltraphong" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
            <div class="modal-dialog">
                <div class="modal-content">
                    <div class="modal-header">
                        <h1 class="modal-title fs-5" id="exampleModalLabel">Thông tin chi tiết phòng</h1>
                        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                    </div>
                    <div class="modal-body">
                        <div class="body_item" style="display: flex; align-items: center;   margin-bottom: 12px;">
                            <span style="font-weight: 600; margin-right: 6px;">Tên phòng:</span>
                            <span><?php echo $room['roomID']; ?></span>
                        </div>
                        <div class="body_item" style="display: flex; align-items: center;   margin-bottom: 12px;">
                            <span style="font-weight: 600; margin-right: 6px;">Loại phòng:</span>
                            <span><?php echo $room['loaiphong']; ?></span>
                        </div>
                        <div class="body_item" style="display: flex; align-items: center;   margin-bottom: 12px;">
                            <span style="font-weight: 600; margin-right: 6px;">Ngày đặt:</span>
                            <span><?php echo $_SESSION['ngaydat']; ?></span>
                        </div>
                        <input value="<?php echo $room['name']; ?>" hidden type="text" name="kh" id="">
                        <input value="<?php echo  $ngaydat[4]; ?>" hidden type="text" name="ngayd" id="">
                        <div class="body_item" style="display: flex; align-items: center;   margin-bottom: 12px;">
                            <span style="font-weight: 600; margin-right: 6px;">Ngày nhận:</span>
                            <span><?php echo $_SESSION['ngaynhan']; ?></span>
                        </div>
                        <div class="body_item" style="display: flex; align-items: center;   margin-bottom: 12px;">
                            <span style="font-weight: 600; margin-right: 6px;">Giá phòng:</span>
                            <span><?php echo $room['gia']; ?></span>
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
                            <input type="text" name="tongthanhtoan" value="<?php echo $room['gia'] + $tong; ?>" hidden id="">
                            <input type="text" name="roomid" value="<?php echo $room['roomID']; ?>" hidden id="">
                            <h4>Tổng thanh toán: <?php echo $room['gia'] + $tong; ?> </h4>
                        <?php } else {

                        ?>
                            <input type="text" name="roomid" value="<?php echo $room['roomID']; ?>" hidden id="">
                            <input type="text" name="tongthanhtoan" value="<?php echo $room['gia']; ?>" hidden id="">
                            <h4>Tổng thanh toán: <?php echo $room['gia']; ?> </h4>

                        <?php } ?>

                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
                        <button type="submit" name="btnthanhtoan" class="btn btn-primary">Thanh toán</button>
                    </div>
                </div>
            </div>
        </form>
    <?php
    }
} else {
    ?>



    <div class="thongtinphongchuadat" style="width: 100%; height: 100vh; display: flex; align-items: center; justify-content: center; flex-direction: column;">
        <h3>Hiện tại bạn chưa đặt phòng nào</h3>
        <a href="./khachhanagdatphong">
            <button type="button" class="btn btn-success">Nhấn để đặt phòng</button>
        </a>
    </div>

<?php
}
?>