<div class="login_container" style="display: flex; align-items: center; justify-content: center; height: 100vh; background-image: url(./public/assets/imgs/download.jpg); background-position: center; background-repeat: no-repeat; background-size: cover;">
    <div class="login" style="width: 500px; padding: 20px; background-color: #cedebd; border-radius: 10px;">
        <h2 style="text-align: center; margin-bottom: 30px;">Đăng nhập</h2>
        <form method="post" action="" autocomplete="off">
            <div class="mb-3">
                <label for="exampleInputEmail1" class="form-label">Tên đăng nhập</label>
                <input type="text" name="tendangnhap" class="form-control" id="exampleInputEmail1" aria-describedby="emailHelp">
                <!-- <div id="emailHelp" class="form-text">We'll never share your email with anyone else.</div> -->
            </div>
            <div class="mb-3">
                <label for="exampleInputPassword1" class="form-label">Mật khẩu</label>
                <input type="password" name="matkhau" class="form-control" id="exampleInputPassword1">
            </div>
            <button style="margin-top: 30px; width: 100%;" type="submit" name="login" class="btn btn-success">Login</button>
        </form>
    </div>
</div>