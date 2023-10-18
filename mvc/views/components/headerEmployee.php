<nav class="navbar navbar-expand-lg " style="padding: 16px 50px; background-color:white; position: fixed; top: 0; left: 0; right: 0; z-index: 100;">
    <div class="container-fluid">
        <div class="collapse navbar-collapse" id="navbarSupportedContent">
            <ul class="navbar-nav me-auto mb-2 mb-lg-0" style="font-weight: 600;display: flex; align-items: center;">
                <?php
                if (isset($_SESSION['role'])) {
                    if ($_SESSION['role'] == "") {
                ?>
                        <li class="nav-item ">
                            <a class="nav-link" aria-current="page" href="../../../../learn_mvc/home">Trang chủ</a>
                        </li>

                        <li class="nav-item">
                            <a class="nav-link" href="../../../../learn_mvc/login">Đăng nhập</a>
                        </li>

                    <?php
                    } else {
                    ?>
                        <li class="nav-item actives">
                            <a class="nav-link" aria-current="page" href="./home">Trang chủ</a>
                        </li>
                        <li class="nav-item">
                            <a class="nav-link" href="../../../../learn_mvc/QuanLyPhongForNhanVien">Quản lý phòng</a>
                        </li>
                        <li class="nav-item">
                            <a class="nav-link" href="../../../../learn_mvc/QuanLyDichVuForNhanVien">Quản lý dịch vụ</a>
                        </li>
                        <li class="nav-item">
                            <a class="nav-link" href="../../../../learn_mvc/quanLyKhachHang">Quản lý khách hàng</a>
                        </li>


                        <div class="dropdown" style="margin-left: 10px; color: #595959;">
                            <li class="  dropdown-toggle" type="button" data-bs-toggle="dropdown" aria-expanded="false">
                                Xem thêm
                            </li>
                            <ul class="dropdown-menu">
                                <li class="nav-item">
                                    <a class="nav-link" href="../../../../learn_mvc/taikhoancanhan">Tài khoản</a>
                                </li>
                                <li class="nav-item">
                                    <a class="nav-link" href="../../../../learn_mvc/dangxuat">Đăng xuất</a>
                                </li>
                            </ul>
                        </div>
                <?php
                    }
                } ?>
            </ul>
            <form class="d-flex" role="search">
                <input class="form-control me-2" type="search" placeholder="Search" aria-label="Search">
                <button class="btn btn-outline-success" type="submit">Search</button>
            </form>
        </div>
    </div>
</nav>