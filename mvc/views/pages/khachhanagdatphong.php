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
        ?>

            <a href="./chitietphongkhachhang/start/<?php echo $room['roomID'] ?>" class="phong <?php echo $status; ?>" style="color: black; text-decoration: none; cursor: pointer;padding: 14px; border-radius: 8px; box-shadow: rgba(100, 100, 111, 0.2) 0px 7px 29px 0px; width: 30%; margin-bottom: 20px; ">
                <div class="img_hotel" style="width: 330px; height: 180px; margin-bottom: 24px;">
                    <img src="<?php echo $room['img'] ?>" alt="" style="width: 100%; height: 100%; object-fit: cover;">
                </div>

                <div class="phong_body" style="display: flex; align-items: stretch; justify-content: space-between;">
                    <i class="fa-solid fa-user"></i>
                    <div class="room_right_card_name"><?php echo $room['name'] ?></div>
                </div>
                <div class="phong_bott" style="display: flex; align-items: stretch; justify-content: space-between;">
                    <div class="phong_time">
                        <i class="fa-solid fa-calendar-days"></i>
                        <span><?php echo $room['trangthaiphong'] ?></span>
                    </div>
                    <div class="phong_clean">Đã dọn dẹp</div>
                </div>

                <button type="button" style="margin-top: 30px; width: 100%;" class="btn btn-success">Xem chi tiết và đặt ngay</button>
            </a>
        <?php
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
        ?>

            <a href="./chitietphongkhachhang/start/<?php echo $room['roomID'] ?>" class="phong <?php echo $status; ?>" style="color: black; text-decoration: none; cursor: pointer;padding: 14px; border-radius: 8px; box-shadow: rgba(100, 100, 111, 0.2) 0px 7px 29px 0px; width: 30%; margin-bottom: 20px;">
                <div class="img_hotel" style="width: 330px; height: 180px; margin-bottom: 24px;">
                    <img src="<?php echo $room['img'] ?>" alt="" style="width: 100%; height: 100%; object-fit: cover;">
                </div>
                <div class="phong_body" style="display: flex; align-items: stretch; justify-content: space-between;">
                    <i class="fa-solid fa-user"></i>
                    <div class="room_right_card_name"><?php echo $room['name'] ?></div>
                </div>
                <div class="phong_bott" style="display: flex; align-items: stretch; justify-content: space-between;">
                    <div class="phong_time">
                        <i class="fa-solid fa-calendar-days"></i>
                        <span><?php echo $room['trangthaiphong'] ?></span>
                    </div>
                    <div class="phong_clean">Đã dọn dẹp</div>
                </div>
                <button type="button" style="margin-top: 30px; width: 100%;" class="btn btn-success">Xem chi tiết và đặt ngay</button>
            </a>
        <?php
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
        ?>

            <a href="./chitietphongkhachhang/start/<?php echo $room['roomID'] ?>" class="phong <?php echo $status; ?>" style="color: black; text-decoration: none; cursor: pointer;padding: 14px; border-radius: 8px; box-shadow: rgba(100, 100, 111, 0.2) 0px 7px 29px 0px; width: 30%; margin-bottom: 20px;">
                <div class="img_hotel" style="width: 330px; height: 180px; margin-bottom: 24px;">
                    <img src="<?php echo $room['img'] ?>" alt="" style="width: 100%; height: 100%; object-fit: cover;">
                </div>
                <div class="phong_body" style="display: flex; align-items: stretch; justify-content: space-between;">
                    <i class="fa-solid fa-user"></i>
                    <div class="room_right_card_name"><?php echo $room['name'] ?></div>
                </div>
                <div class="phong_bott" style="display: flex; align-items: stretch; justify-content: space-between;">
                    <div class="phong_time">
                        <i class="fa-solid fa-calendar-days"></i>
                        <span><?php echo $room['trangthaiphong'] ?></span>
                    </div>
                    <div class="phong_clean">Đã dọn dẹp</div>
                </div>
                <button type="button" style="margin-top: 30px; width: 100%;" class="btn btn-success">Xem chi tiết và đặt ngay</button>
            </a>
        <?php
        } ?>
    </div>


</div>