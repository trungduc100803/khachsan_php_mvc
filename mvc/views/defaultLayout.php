<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.1/dist/css/bootstrap.min.css">
    <link rel="stylesheet" href="./public/assets/fontawesome-free-6.4.2-web/css/all.css">
    <link rel="stylesheet" href="./mvc/css/home.css">
    <link rel="stylesheet" href="../../../learn_mvc/mvc/css/defaultLayout.css">
</head>

<body>
    <?php
    if (isset($_SESSION['role'])) {
        $role = $_SESSION['role'];
        if ($role == "") {
            require_once "components/headerCustomer.php";
        } else if ($role == "khach") {
            require_once "components/headerCustomer.php";
        } else if ($role == "nhanvien") {
            require_once "components/headerEmployee.php";
        } else {
            require_once "components/headerAdmin.php";
        }
    }
    ?>

    <div class="inner" style="background-image: linear-gradient(to right, #C8E4B2, white); min-height: 100vh; margin-top: 70px; padding-bottom: 100px;">
        <div class="" style="padding-left:70px;  padding-right: 70px; ">
            <?php require_once "./mvc/views/pages/" . $data['container'] . ".php"; ?>
        </div>
    </div>
    <?php require_once "components/footer.php"; ?>
</body>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.1/dist/js/bootstrap.bundle.min.js"></script>
<script src="../../../learn_mvc/mvc/js/home.js"></script>
<style>
    .nav-item.actives a {
        color: red;
    }
</style>
<script>
    const navItem = document.querySelectorAll('.nav-item')
    navItem.forEach(item => {
        item.addEventListener('click', () => {
            item.classList.add('actives')
            
        })
    })
</script>

</html>