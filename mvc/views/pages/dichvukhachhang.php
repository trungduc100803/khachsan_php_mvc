<?php
if (mysqli_num_rows($data['dataDichvu']) > 0) {
?>
  <form style="padding-top: 50px;" action="" method="post">
    <table style="width: 100%; border: 1px solid #333; padding: 10px; border-radius: 10px; margin-top: 30px;">
      <thead>
        <th style="padding: 10px;">Số hợp đồng</th>
        <th style="padding: 10px;">Tên dịch vụ</th>
        <th style="padding: 10px;">Số lượng</th>
        <th style="padding: 10px;">Giá</th>
        <th style="padding: 10px;">Thành tiền</th>
        <th style="padding: 10px;">Chức năng</th>
      </thead>
      <tbody>
        <?php
        while ($service = mysqli_fetch_array($data['dataDichvu'])) {
        ?>
          <tr>
            <td style="padding: 10px; padding-left: 10px;"><?php echo $service['sohopdong'] ?></td>
            <td style="padding: 10px;"><?php echo $service['tendichvu'] ?></td>
            <td style="padding: 10px;"><?php echo $service['soluong'] ?></td>
            <td style="padding: 10px;"><?php echo $service['gia'] ?></td>
            <td style="padding: 10px;"><?php echo $service['thanhtien'] ?></td>
            <td style="padding: 10px; padding-right: 10px;">
              <a href="../../../../learn_mvc/dichvukhachhang/sua/<?php echo $service['sohopdong'] ?>">
                <button type="button" class="btn btn-info" data-bs-toggle="modal" data-bs-target="#exampleModalsua">
                  Sửa
                </button>
              </a>
              <a href="../../../../learn_mvc/dichvukhachhang/delete/<?php echo $service['sohopdong'] ?>/<?php echo $service['tendichvu'] ?>">
                <button type="button" id="btnxoa" class="btn btn-danger">
                  Xóa
                </button>
              </a>
            </td>
            <td>

            </td>
          </tr>
        <?php } ?>
      </tbody>
    </table>
  </form>
<?php } else {
?>
  <div class="thongtinphongchuadat" style="width: 100%; height: 100vh; display: flex; align-items: center; justify-content: center; flex-direction: column;">
    <h3>Hiện tại bạn chưa đặt bất kì dịch vụ nào!</h3>
  </div>
<?php } ?>