<?php
$customer = $data['customer'];
$dataC  = mysqli_fetch_row($customer);
$data_phong_da_dat = $data['data_phong_da_dat'];
?>

<div class="chitietkhachhang" style="display: flex;">
    <div class="chitietkhachhang_left" style="width: 35%;">
        <h4 style="margin: 30px 0;">Thông tin chi tiết khách hàng</h4>
        <div style="display: flex; align-items: stretch; margin-bottom: 6px;" class="chitietkhachhang_item">
            <span style="font-weight: 600; margin-right: 10px;">Họ tên:</span>
            <p><?php echo $dataC[1] ?></p>
        </div>
        <div style="display: flex; align-items: stretch; margin-bottom: 6px;" class="chitietkhachhang_item">
            <span style="font-weight: 600; margin-right: 10px;">Số CCCD/CMND:</span>
            <p><?php echo $dataC[2] ?></p>
        </div>
        <div style="display: flex; align-items: stretch; margin-bottom: 6px;" class="chitietkhachhang_item">
            <span style="font-weight: 600; margin-right: 10px;">Số điện thoại:</span>
            <p><?php echo $dataC[3] ?></p>
        </div>
        <div style="display: flex; align-items: stretch; margin-bottom: 6px;" class="chitietkhachhang_item">
            <span style="font-weight: 600; margin-right: 10px;">Ngày sinh:</span>
            <p><?php echo $dataC[4] ?></p>
        </div>
        <div style="display: flex; align-items: stretch; margin-bottom: 6px;" class="chitietkhachhang_item">
            <span style="font-weight: 600; margin-right: 10px;">Quốc tịch:</span>
            <p><?php echo $dataC[5] ?></p>
        </div>
        <div style="display: flex; align-items: stretch; margin-bottom: 6px;" class="chitietkhachhang_item">
            <span style="font-weight: 600; margin-right: 10px;">Địa chỉ:</span>
            <p><?php echo $dataC[6] ?></p>
        </div>

        <div style="display: flex; align-items: stretch; margin-bottom: 6px;" class="chitietkhachhang_item">
            <span style="font-weight: 600; margin-right: 10px;">Giới tính:</span>
            <p><?php echo $dataC[7] ?></p>
        </div>
    </div>


    <div class="chitietkhachhang_right" style="width: 65%;">
        <h4 style="margin: 30px 0;">Thông tin các phòng khách đã đặt</h4>
        <table style="width: 100%; border: 1px solid #333; padding: 10px; border-radius: 10px; margin-top: 30px;">
            <thead>
                <th style="padding: 10px;">Tên phòng</th>
                <th style="padding: 10px;">Loại phòng</th>
                <th style="padding: 10px;">Ngày đặt</th>
                <th style="padding: 10px;">Ngày nhận</th>
                <th style="padding: 10px;">Thành tiền</th>
            </thead>
            <?php
            while ($phong_da_dat = mysqli_fetch_assoc($data_phong_da_dat)) {
            ?>
                <tbody>
                    <tr>
                        <td style="padding: 10px; padding-left: 10px;"><?php echo $phong_da_dat['roomID']; ?></td>
                        <td style="padding: 10px;"><?php echo $phong_da_dat['loaiphong']; ?></td>
                        <td style="padding: 10px;"><?php echo $phong_da_dat['ngaydat']; ?></td>
                        <td style="padding: 10px;"><?php echo $phong_da_dat['ngaynhan']; ?></td>
                        <td style="padding: 10px; padding-right: 10px;">
                            <?php echo $phong_da_dat['gia']; ?>
                        </td>
                    </tr>
                </tbody>
            <?php } ?>
        </table>
    </div>

</div>