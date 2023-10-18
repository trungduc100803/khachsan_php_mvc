<?php
$phongDaDat = mysqli_fetch_assoc($data['dataPhongDaDat']);
?>

<form action="" method="post" style="display: flex;">
    <div class="thongtinphong" style="padding-top: 30px; width: 50%;">

        <h4 style="margin-bottom: 20px;">Thông tin chi tiết hóa đơn</h4>

        <div class="phong_item" style="display: flex; align-items: stretch;">
            <span style="margin-right: 10px; font-weight: 600;">Tên phòng:</span>
            <p><?php echo "P" . $phongDaDat['roomID']; ?></p>
        </div>
        <div class="phong_item" style="display: flex; align-items: stretch;">
            <span style="margin-right: 10px; font-weight: 600;">Loại phòng:</span>
            <p><?php echo  $phongDaDat['loaiphong']; ?></p>
        </div>
        <div class="phong_item" style="display: flex; align-items: stretch;">
            <span style="margin-right: 10px; font-weight: 600;">Ngày đặt:</span>
            <p><?php echo $phongDaDat['ngaydat']; ?></p>
        </div>
        <div class="phong_item" style="display: flex; align-items: stretch;">
            <span style="margin-right: 10px; font-weight: 600;">Khách hàng:</span>
            <p><?php echo $phongDaDat['name']; ?></p>
        </div>
        <div class="phong_item" style="display: flex; align-items: stretch;">
            <span style="margin-right: 10px; font-weight: 600;">Giá:</span>
            <p><?php echo $phongDaDat['gia']; ?></p>
        </div>
    </div>

    <div class="dichvu">
        <?php
        if (mysqli_num_rows($data['datadichvu']) > 0) {

        ?>
            <h4 style="margin-bottom: 20px;">Thông tin dịch vụ dịch vụ</h4>
            <table style="width: 100%; border: 1px solid #333; padding: 10px; border-radius: 10px; margin-top: 30px;">
                <thead>
                    <th style="padding: 10px;">Tên dịch vụ</th>
                    <th style="padding: 10px;">Số lượng</th>
                    <th style="padding: 10px;">Giá</th>
                    <th style="padding: 10px;">Thành tiền</th>
                </thead>
                <tbody>
                    <?php
                    while ($dichvu = mysqli_fetch_assoc($data['datadichvu'])) {
                    ?>
                        <tr>
                            <td style="padding: 10px;"><?php echo $dichvu['tendichvu'] ?></td>
                            <td style="padding: 10px;"><?php echo $dichvu['soluong'] ?></td>
                            <td style="padding: 10px;"><?php echo $dichvu['gia'] ?></td>
                            <td style="padding: 10px;"><?php echo $dichvu['thanhtien'] ?></td>
                        </tr>
                    <?php } ?>
                </tbody>
            </table>
        <?php
        } else {
        ?>
            <h4 style="margin-bottom: 20px;">Phòng không sử dụng dịch vụ</h4>
        <?php } ?>
    </div>


</form>