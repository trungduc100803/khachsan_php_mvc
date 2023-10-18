<?php
$service = mysqli_fetch_row($data['service']);

?>


<form method="post" style="display: flex; align-items: center; justify-content: center; height: 100vh;" id="exampleModalsua" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
    <div class="modal-dialog" style="width: 500px; border: 1px solid #ccc; padding: 20px; border-radius: 10px;">
        <div class="modal-content">
            <div class="modal-header">
                <h1 class="modal-title fs-5" style="margin-bottom: 40px;" id="exampleModalLabel">Sửa thông tin dịch vụ</h1>
                <!-- <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button> -->
            </div>
            <div class="modal-body">
                <label for="">Tên dịch vụ</label>
                <input name="tendv" style="margin-top: 6px; margin-bottom: 20px;" value="<?php echo $service[1] ?>" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">

                <label for="">Giá dịch vụ</label>
                <input name="giadv" style="margin-top: 6px; margin-bottom: 20px;" value="<?php echo $service[2] ?>" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">

            </div>
            <div class="modal-footer" style="margin-top: 40px;">
                <button type="submit" name="dongdv" class="btn btn-secondary" style="margin-right: 20px;" data-bs-dismiss="modal">Đóng</button>
                <button type="submit" name="suadv" class="btn btn-primary">Lưu</button>
            </div>
        </div>
    </div>
</form>