<form method="post" class="quanlydoanhthu" style="padding-top: 50px;">
    <h4>Doanh thu</h4>
    <div class="quanlydoanhthu_main" style="margin: 30px 0;">
        <div class="quanlydoanhthu_control" style="margin-bottom: 30px; display: flex; align-items: center;">
            <select style="width: 200px; margin-right: 12px;" name="lochoadon" class="form-select" aria-label="Default select example">
                <option selected>-- LỌC --</option>
                <option value="theongay">Theo ngày</option>
                <option value="theothang">Theo tháng</option>
                <option value="theonam">Theo năm</option>
            </select>
            <button type="submit" name="locdoanhthu" class="btn btn-success">
                Lọc
            </button>
        </div>

        <table style="width: 100%; border: 1px solid #333; padding: 10px; border-radius: 10px; margin-top: 30px; margin-bottom: 50px;">
            <thead>
                <th style="padding: 10px;">Số hợp đồng</th>
                <th style="padding: 10px;">Tên phòng</th>
                <th style="padding: 10px;">Khách hàng</th>
                <th style="padding: 10px;">Ngày đặt</th>
                <th style="padding: 10px;">Tổng thanh toán</th>
                <th style="padding: 10px;">Thao tác</th>
            </thead>
            <?php
            $tongchiphi = 0;
            if (isset($data['dataLoc'])) {
                for ($i = 0; $i < count($data['dataLoc']); $i++) {
                    $loc = $data['dataLoc'][$i];

            ?>
                    <tbody>
                        <tr>
                            <input type="number" name="chiphi" hidden value="<?php $tongchiphi += $loc['tongchiphi'] ?>" id="">
                            <td style="padding: 10px ; padding-left: 10px;"><?php echo $loc['sohopdong'] ?></td>
                            <td style="padding: 10px;"><?php echo $loc['roomID'] ?></td>
                            <td style="padding: 10px;"><?php echo $loc['name'] ?></td>
                            <td style="padding: 10px ;"><?php echo $loc['ngaydat'] ?></td>
                            <td style="padding: 10px; padding-right: 10px;"><?php echo $loc['tongchiphi'] ?></td>
                            <td style="padding: 10px; padding-right: 10px;">
                                <a href="../../../../learn_mvc/quanlydoanhthu/chitiet/<?php echo $loc['sohopdong'] ?>">
                                    <button type="button" class="btn btn-info">
                                        Chi tiết
                                    </button>
                                </a>
                            </td>
                        </tr>
                    </tbody>
                <?php }
            } else { ?>

            <?php } ?>
        </table>

        <div style="display: flex; align-items: center;">
            <h5 style="margin-right: 12px;">Tổng doanh thu:</h5>
            <h3 style="color: red;"><?php echo $tongchiphi; ?></h3>

        </div>
        <button type="submit" name="xuatexcel" class="btn btn-success">
            Xuất Excel
        </button>
    </div>


    <h4>Danh sách các phòng đã trả</h4>
    <table style="width: 100%; border: 1px solid #333; padding: 10px; border-radius: 10px; margin-top: 30px;">
        <thead>
            <th style="padding: 10px;">Số hợp đồng</th>
            <th style="padding: 10px;">Tên phòng</th>
            <th style="padding: 10px;">Khách hàng</th>
            <th style="padding: 10px;">Ngày đặt</th>
            <th style="padding: 10px;">Tổng thanh toán</th>
            <th style="padding: 10px;">Thao tác</th>
        </thead>
        <?php
        if (mysqli_num_rows($data['dataHoadon']) > 0) {
            while ($hoadon = mysqli_fetch_array($data['dataHoadon'])) {
        ?>
                <tbody>
                    <tr>
                        <td style="padding: 10px ; padding-left: 10px;"><?php echo $hoadon['sohopdong'] ?></td>
                        <td style="padding: 10px;"><?php echo $hoadon['roomID'] ?></td>
                        <td style="padding: 10px;"><?php echo $hoadon['name'] ?></td>
                        <td style="padding: 10px ;"><?php echo $hoadon['ngaydat'] ?></td>
                        <td style="padding: 10px; padding-right: 10px;"><?php echo $hoadon['tongchiphi'] ?></td>
                        <td style="padding: 10px; padding-right: 10px;">
                            <a href="../../../../learn_mvc/quanlydoanhthu/chitiet/<?php echo $hoadon['sohopdong'] ?>">
                                <button type="button" class="btn btn-info">
                                    Chi tiết
                                </button>
                            </a>
                        </td>
                    </tr>
                </tbody>

        <?php }
        } ?>
    </table>
</form>