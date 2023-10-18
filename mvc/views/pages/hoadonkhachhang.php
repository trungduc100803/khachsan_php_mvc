<?php
$dataHoaDon = $data['dataHoaDon'];
?>

<?php
if (mysqli_num_rows($dataHoaDon) > 0) {
?>
    <div class="hoadonkhachhang">
        <form style="padding-top: 50px;" action="" method="post">
            <table style="width: 100%; border: 1px solid #333; padding: 10px; border-radius: 10px; margin-top: 30px;">
                <thead>
                    <th style="padding: 10px;">Số hợp đồng</th>
                    <th style="padding: 10px;">Ngày đặt</th>
                    <th style="padding: 10px;">Ngày nhận</th>
                    <th style="padding: 10px;">Ngày trả</th>
                    <th style="padding: 10px;">Tổng chi phí</th>
                    <th style="padding: 10px;">Chức năng</th>
                </thead>
                <tbody>
                    <?php
                    while ($hoadon = mysqli_fetch_array($dataHoaDon)) {
                    ?>
                        <tr>
                            <td style="padding: 10px; padding-left: 10px;"><?php echo $hoadon['sohopdong'] ?></td>
                            <td style="padding: 10px;"><?php echo $hoadon['ngaydat'] ?></td>
                            <td style="padding: 10px;"><?php echo $hoadon['ngaynhan'] ?></td>
                            <td style="padding: 10px;"><?php echo $hoadon['ngaytra'] ?></td>
                            <td style="padding: 10px;"><?php echo $hoadon['tongchiphi'] ?></td>
                            <td style="padding: 10px; padding-right: 10px;">
                                <a href="../../../../learn_mvc/hoadonkhachhang/chitiet/<?php echo $hoadon['sohopdong'] ?>">
                                    <button type="button" class="btn btn-info">
                                        Chi tiết
                                    </button>
                                </a>
                            </td>
                        </tr>
                    <?php } ?>
                </tbody>
            </table>
        </form>
    </div>
<?php
} else {

?>

    <div class="hoadonkhachhangchuadat" style="width: 100%; height: 100vh; display: flex; align-items: center; justify-content: center; flex-direction: column;">
        <h3>Chưa có bất kì hóa đơn nào</h3>
        <button type="button" class="btn btn-success">Nhấn để đặt phòng</button>
    </div>

<?php
}
?>