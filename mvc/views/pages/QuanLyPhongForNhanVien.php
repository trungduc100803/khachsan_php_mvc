<style>
    .phong {
        background-color: #ccc;
    }

    .phong.dadat {
        background-color: #64CCC5;
    }

    .phong.danhan {
        background-color: #C08261;
    }
</style>

<div class="quan_ly_phong" style="margin: 0;">
    <div class="control" style="margin-bottom: 30px; padding-top: 30px; display: flex; align-items: center; justify-content: space-between;">
        <select style="width: 200px;" name="locphong" class="form-select" aria-label="Default select example">
            <option selected>-- LỌC PHÒNG --</option>
            <option value="phongdon">Phòng đơn</option>
            <option value="phongdoi">Phòng đôi</option>
            <option value="phonggiadinh">Phòng gia đình</option>
            <option value="tatca">Tất cả</option>
        </select>
    </div>


    <!-- phong don -->
    <h3 style="margin-bottom: 30px;">Phòng đơn</h3>
    <div class="danhsachphong" style="display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; color: black; text-decoration: none;">
        <?php
        $i = 0;
        while ($room = mysqli_fetch_array($data['roomAlone'])) {
            $status = "";
            if ($room['trangthaiphong'] == "Đã đặt") {
                $status = "dadat";
            } else if ($room['trangthaiphong'] == "Đã nhận") {
                $status = "danhan";
            }
            if ($room['trangthaiphong'] == "Đang trống") {
        ?>

                <a href="./ChiTietPhongForNhanVien/start/<?php echo $room['roomID'] ?>" class="phong <?php echo $status; ?>" style="color: black; text-decoration: none; cursor: pointer;padding: 14px; border-radius: 8px; box-shadow: rgba(100, 100, 111, 0.2) 0px 7px 29px 0px; width: 30%; margin-bottom: 20px; ">
                    <div class="phong_top" style="display: flex; align-items: stretch; justify-content: space-between;">
                        <p><?php echo $room['roomID'] ?></p>
                        <div class="trangthaiphong"><?php echo $room['trangthaiphong'] ?></div>
                    </div>
                    <div class="phong_body" style="display: flex; align-items: stretch; justify-content: space-between;">
                        <i class="fa-solid fa-user"></i>
                        <div class="room_right_card_name"><?php echo $room['name'] ?></div>
                    </div>
                    <div class="phong_bott" style="display: flex; align-items: stretch; justify-content: space-between;">
                        <div class="phong_time">
                            <i class="fa-solid fa-calendar-days"></i>
                            <span>Chưa đặt</span>
                        </div>
                    </div>
                </a>
            <?php
            } else {
                $hoadonModel = $data["hoadonModel"];
                $datangaydat = $hoadonModel->getAHoaDon($room['name'], $room['roomID']);
                $ngaydat = mysqli_fetch_row($datangaydat);
            ?>
                <a href="./ChiTietPhongForNhanVien/start/<?php echo $room['roomID'] ?>" class="phong <?php echo $status; ?>" style="color: black; text-decoration: none; cursor: pointer;padding: 14px; border-radius: 8px; box-shadow: rgba(100, 100, 111, 0.2) 0px 7px 29px 0px; width: 30%; margin-bottom: 20px; ">
                    <div class="phong_top" style="display: flex; align-items: stretch; justify-content: space-between;">
                        <p><?php echo $room['roomID'] ?></p>
                        <div class="trangthaiphong"><?php echo $room['trangthaiphong'] ?></div>
                    </div>
                    <div class="phong_body" style="display: flex; align-items: stretch; justify-content: space-between;">
                        <i class="fa-solid fa-user"></i>
                        <div class="room_right_card_name"><?php echo $room['name'] ?></div>
                    </div>
                    <div class="phong_bott" style="display: flex; align-items: stretch; justify-content: space-between;">
                        <div class="phong_time">
                            <i class="fa-solid fa-calendar-days"></i>
                            <span><?php echo $ngaydat[3] ?></span>
                        </div>
                    </div>
                </a>

        <?php }
        } ?>
    </div>



    <!-- phong doi -->
    <h3 style="margin-bottom: 30px;">Phòng đôi</h3>
    <div class="danhsachphong" style="display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; color: black; text-decoration: none;">
        <?php
        $i = 0;
        while ($room = mysqli_fetch_array($data['roomCouple'])) {
            $status = "";
            if ($room['trangthaiphong'] == "Đã đặt") {
                $status = "dadat";
            } else if ($room['trangthaiphong'] == "Đã nhận") {
                $status = "danhan";
            }
            if ($room['trangthaiphong'] == "Đang trống") {
        ?>

                <a href="./ChiTietPhongForNhanVien/start/<?php echo $room['roomID'] ?>" class="phong <?php echo $status; ?>" style="color: black; text-decoration: none; cursor: pointer;padding: 14px; border-radius: 8px; box-shadow: rgba(100, 100, 111, 0.2) 0px 7px 29px 0px; width: 30%; margin-bottom: 20px;">
                    <div class="phong_top" style="display: flex; align-items: stretch; justify-content: space-between;">
                        <p><?php echo $room['roomID'] ?></p>
                        <div class="trangthaiphong"><?php echo $room['trangthaiphong'] ?></div>
                    </div>
                    <div class="phong_body" style="display: flex; align-items: stretch; justify-content: space-between;">
                        <i class="fa-solid fa-user"></i>
                        <div class="room_right_card_name"><?php echo $room['name'] ?></div>
                    </div>
                    <div class="phong_bott" style="display: flex; align-items: stretch; justify-content: space-between;">
                        <div class="phong_time">
                            <i class="fa-solid fa-calendar-days"></i>
                            <span>Chưa đặt</span>
                        </div>

                    </div>
                </a>
            <?php
            } else {
                $hoadonModel = $data["hoadonModel"];
                $datangaydat = $hoadonModel->getAHoaDon($room['name'], $room['roomID']);
                $ngaydat = mysqli_fetch_row($datangaydat);
            ?>
                <a href="./ChiTietPhongForNhanVien/start/<?php echo $room['roomID'] ?>" class="phong <?php echo $status; ?>" style="color: black; text-decoration: none; cursor: pointer;padding: 14px; border-radius: 8px; box-shadow: rgba(100, 100, 111, 0.2) 0px 7px 29px 0px; width: 30%; margin-bottom: 20px;">
                    <div class="phong_top" style="display: flex; align-items: stretch; justify-content: space-between;">
                        <p><?php echo $room['roomID'] ?></p>
                        <div class="trangthaiphong"><?php echo $room['trangthaiphong'] ?></div>
                    </div>
                    <div class="phong_body" style="display: flex; align-items: stretch; justify-content: space-between;">
                        <i class="fa-solid fa-user"></i>
                        <div class="room_right_card_name"><?php echo $room['name'] ?></div>
                    </div>
                    <div class="phong_bott" style="display: flex; align-items: stretch; justify-content: space-between;">
                        <div class="phong_time">
                            <i class="fa-solid fa-calendar-days"></i>
                            <span><?php echo $ngaydat[3] ?></span>
                        </div>
                    </div>
                </a>
        <?php }
        } ?>
    </div>


    <!-- phong gia dinh -->
    <h3 style="margin-bottom: 30px;">Phòng gia đình</h3>
    <div class="danhsachphong" style="display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; color: black; text-decoration: none;">
        <?php
        $i = 0;
        while ($room = mysqli_fetch_array($data['roomFamily'])) {
            $status = "";
            if ($room['trangthaiphong'] == "Đã đặt") {
                $status = "dadat";
            } else if ($room['trangthaiphong'] == "Đã nhận") {
                $status = "danhan";
            }
            if ($room['trangthaiphong'] == "Đang trống") {
        ?>

                <a href="./ChiTietPhongForNhanVien/start/<?php echo $room['roomID'] ?>" class="phong <?php echo $status; ?>" style="color: black; text-decoration: none; cursor: pointer;padding: 14px; border-radius: 8px; box-shadow: rgba(100, 100, 111, 0.2) 0px 7px 29px 0px; width: 30%; margin-bottom: 20px;">
                    <div class="phong_top" style="display: flex; align-items: stretch; justify-content: space-between;">
                        <p><?php echo $room['roomID'] ?></p>
                        <div class="trangthaiphong"><?php echo $room['trangthaiphong'] ?></div>
                    </div>
                    <div class="phong_body" style="display: flex; align-items: stretch; justify-content: space-between;">
                        <i class="fa-solid fa-user"></i>
                        <div class="room_right_card_name"><?php echo $room['name'] ?></div>
                    </div>
                    <div class="phong_bott" style="display: flex; align-items: stretch; justify-content: space-between;">
                        <div class="phong_time">
                            <i class="fa-solid fa-calendar-days"></i>
                            <span>Chưa đặt</span>
                        </div>
                    </div>
                </a>
            <?php
            } else {
                $hoadonModel = $data["hoadonModel"];
                $datangaydat = $hoadonModel->getAHoaDon($room['name'], $room['roomID']);
                $ngaydat = mysqli_fetch_row($datangaydat);
            ?>
                <a href="./ChiTietPhongForNhanVien/start/<?php echo $room['roomID'] ?>" class="phong <?php echo $status; ?>" style="color: black; text-decoration: none; cursor: pointer;padding: 14px; border-radius: 8px; box-shadow: rgba(100, 100, 111, 0.2) 0px 7px 29px 0px; width: 30%; margin-bottom: 20px;">
                    <div class="phong_top" style="display: flex; align-items: stretch; justify-content: space-between;">
                        <p><?php echo $room['roomID'] ?></p>
                        <div class="trangthaiphong"><?php echo $room['trangthaiphong'] ?></div>
                    </div>
                    <div class="phong_body" style="display: flex; align-items: stretch; justify-content: space-between;">
                        <i class="fa-solid fa-user"></i>
                        <div class="room_right_card_name"><?php echo $room['name'] ?></div>
                    </div>
                    <div class="phong_bott" style="display: flex; align-items: stretch; justify-content: space-between;">
                        <div class="phong_time">
                            <i class="fa-solid fa-calendar-days"></i>
                            <span><?php echo $ngaydat[3] ?></span>
                        </div>
                    </div>
                </a>

        <?php  }
        } ?>
    </div>



    <!-- model them phong -->
    <form method="POST" class="modal fade" id="exampleModal" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header">
                    <h1 class="modal-title fs-5" id="exampleModalLabel">Thêm phòng</h1>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <label for="">Tên phòng</label>
                    <input name="tenphong" style="margin-top: 6px; margin-bottom: 20px;" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">

                    <label for="">Loại phòng</label>
                    <select name="loaiphong" style="margin-top: 6px; margin-bottom: 20px;" class="form-select" aria-label="Default select example">
                        <option selected>-- LOẠI PHÒNG --</option>
                        <option value="Phòng đơn">Phòng đơn</option>
                        <option value="Phòng đôi">Phòng đôi</option>
                        <option value="Phòng gia đình">Phòng gia đình</option>
                    </select>

                    <label for="">Trạng thái phòng</label>
                    <select name="trangthaiphong" style="margin-top: 6px; margin-bottom: 20px;" class="form-select" aria-label="Default select example">
                        <option selected>-- TRẠNG THÁI PHÒNG --</option>
                        <option value="Đang trống">Đang trống</option>
                        <option value="Đã đặt">Đã đặt</option>
                        <option value="Đã nhận">Đã nhận</option>
                    </select>

                    <label for="">Giá</label>
                    <input name="gia" style="margin-top: 6px; margin-bottom: 20px;" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Đóng</button>
                    <button name="themphong" type="submit" class="btn btn-primary">Lưu</button>
                </div>
            </div>
        </div>
    </form>
</div>