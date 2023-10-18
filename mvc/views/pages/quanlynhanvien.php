<div class="quanlynhanvien">
    <h4>Thông tin chi tiết các nhân viên</h4>
    <table style="width: 100%; border: 1px solid #333; padding: 10px; border-radius: 10px; margin-top: 30px;">
        <thead>
            <th style="padding: 10px;">Tên nhân viên</th>
            <th style="padding: 10px;">Số điện thoại</th>
            <th style="padding: 10px;">Địa chỉ</th>
            <th style="padding: 10px;">Giới tính</th>
            <th style="padding: 10px;">Vị trí công việc</th>
            <th style="padding: 10px;">Chức năng</th>
        </thead>
        <tbody>
            <tr>
                <td style="padding: 10px ; padding-left: 10px;">trung duc</td>
                <td style="padding: 10px;">132365959</td>
                <td style="padding: 10px;">Hà Nội</td>
                <td style="padding: 10px ;">Nam</td>
                <td style="padding: 10px ;">Nhân viên sale</td>
                <td style="padding: 10px; padding-right: 10px;">
                    <button type="button" class="btn btn-info" data-bs-toggle="modal" data-bs-target="#exampleModalsua">
                        Sửa
                    </button>
                    <button type="button" class="btn btn-danger" data-bs-toggle="modal" data-bs-target="#exampleModalxoa">
                        Xóa
                    </button>
                </td>
            </tr>
        </tbody>
    </table>


    <div class="modal fade" id="exampleModalsua" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header">
                    <h1 class="modal-title fs-5" id="exampleModalLabel">Sửa thông tin nhân viên</h1>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">

                    <label for="">Tên nhân viên</label>
                    <input style="margin-top: 6px; margin-bottom: 20px;" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">

                    <label for="">Số CCCD</label>
                    <input style="margin-top: 6px; margin-bottom: 20px;" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">

                    <label for="">Số điện thoại</label>
                    <input style="margin-top: 6px; margin-bottom: 20px;" type="text" class="form-control" aria-label="" aria-describedby="basic-addon1">

                    <label for="">Ngày sinh</label>
                    <input style="margin-top: 6px; margin-bottom: 20px;" type="date" class="form-control" aria-label="" aria-describedby="basic-addon1">

                    <label for="">Quốc tịch</label>
                    <select style="margin-top: 6px; margin-bottom: 20px;" class="form-select" aria-label="Default select example">
                        <option selected>Quốc tịch</option>
                        <option value="vn">Việt Nam</option>
                        <option value="hq">Hàn Quốc</option>
                        <option value="sing">Singapor</option>
                    </select>

                    <label for="">Địa chỉ</label>
                    <select style="margin-top: 6px; margin-bottom: 20px;" class="form-select" aria-label="Default select example">
                        <option selected>Địa chỉ</option>
                        <option value="hn">Hà Nội</option>
                        <option value="th">Thanh Hóa</option>
                        <option value="hd">Hải Dương</option>
                    </select>

                    <label for="">Giới tính</label>
                    <div style="margin-top: 6px; margin-bottom: 20px;" class="check">
                        <input class="form-check-input" type="radio" name="flexRadioDefault" id="flexRadioDefault1">
                        <label class="form-check-label" for="flexRadioDefault1">
                            Nam
                        </label>

                        <input class="form-check-input" type="radio" name="flexRadioDefault" id="flexRadioDefault1">
                        <label class="form-check-label" for="flexRadioDefault1">
                            Nữ
                        </label>
                    </div>
                    <label for="">Vị trí công việc</label>
                    <select style="margin-top: 6px; margin-bottom: 20px;" class="form-select" aria-label="Default select example">
                        <option selected>Vị trí</option>
                        <option value="lt">Lễ tân</option>
                        <option value="qd">Nhân viên quét dọn</option>
                    </select>

                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Đóng</button>
                    <button type="button" class="btn btn-primary">Lưu✔✔</button>
                </div>
            </div>
        </div>
    </div>

    <!-- modal xoa -->
    <div class="modal fade" id="exampleModalxoa" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header">
                    <h1 class="modal-title fs-5" id="exampleModalLabel">Xóa nhân viên</h1>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    Bạn chắc chắn muốn xóa nhân viên này!!!!😊❤😊😊❤
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Đóng</button>
                    <button type="button" class="btn btn-primary">Đúng vậy 👌👌👌</button>
                </div>
            </div>
        </div>
    </div>
</div>