<?php
$customers = $data['customers'];

?>

<div class="quanLyKhachHang">
    <h4>Danh sách khách hàng</h4>
    <table style="width: 100%; border: 1px solid #333; padding: 10px; border-radius: 10px; margin-top: 30px;">
        <thead>
            <th style="padding: 10px;">Tên khách hàng</th>
            <th style="padding: 10px;">Số điện thoại</th>
            <th style="padding: 10px;">Địa chỉ</th>
            <th style="padding: 10px;">Giới tính</th>
            <th style="padding: 10px;">Chức năng</th>
        </thead>
        <tbody>
            <?php
            $i = 0;
            while ($customer = mysqli_fetch_array($customers)) {
                $i++;
            ?>
                <tr>
                    <td style="padding: 10px ; padding-left: 10px;"><?php echo $customer['name'] ?></td>
                    <td style="padding: 10px;"><?php echo $customer['sdt'] ?></td>
                    <td style="padding: 10px;"><?php echo $customer['diachi'] ?></td>
                    <td style="padding: 10px ;"><?php echo $customer['gioitinh'] ?></td>
                    <td style="padding: 10px; padding-right: 10px;">
                        <a href="./quanLyKhachHang/edit/<?php echo $customer['customerID'] ?>" type="button" class="exampleModalsua btn btn-info" data-bs-toggle="modal">
                            Sửa
                        </a>
                        <button type="button" class="btn btn-danger" data-bs-toggle="modal" data-bs-target="#exampleModalxoa">

                            <a data-bs-toggle="modal" data-bs-target="#exampleModalxoa" href="./edit/<?php echo $customer['customerID'] ?>" style="text-decoration: none;  color: white;">Xóa</a>

                        </button>
                        <button type="button" class="btn btn-success" data-bs-toggle="modal">
                            <a href="./quanLyKhachHang/chitietkhachhang/<?php echo $customer['customerID'] . "/" . $customer['name'] ?>" style="text-decoration: none;  color: white;">Chi tiết</a>
                        </button>
                    </td>
                </tr>
            <?php
            } ?>
        </tbody>
    </table>


    <!-- modal sua -->


    <!-- modal xoa -->
    <form method="post" class="modal fade" id="exampleModalxoa" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header">
                    <h1 class="modal-title fs-5" id="exampleModalLabel">Xóa khách hàng</h1>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    Bạn chắc chắn muốn xóa khách hàng này!!!!😊❤😊😊❤
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Đóng</button>
                    <button type="submit" name="btnxoakh" class="btn btn-primary">Đúng vậy 👌👌👌</button>
                </div>
            </div>
        </div>
    </form>
</div>