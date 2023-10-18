<div class="home">

    <!-- Slideshow container -->
    <div class="slideshow-container" style="width: 100%; max-width: 2000px; height: 500px; padding-top: 30px; border-radius: 10px; overflow: hidden;">

        <!-- Full-width images with number and caption text -->
        <div class="mySlides fade" style="height: 100%;">
            <!-- <div class="numbertext">1 / 3</div> -->
            <img src="./public/assets/imgs/banner1.jpg" style="width:100%; height: 100%;">
            <!-- <div class="text">Caption Text</div> -->
        </div>

        <div class="mySlides fade" style="height: 100%;">
            <!-- <div class="numbertext">2 / 3</div> -->
            <img src="./public/assets/imgs/banner2.jpg" style="width:100%; height: 100%;">
            <!-- <div class="text">Caption Two</div> -->
        </div>

        <div class="mySlides fade" style="height: 100%;">
            <!-- <div class="numbertext">3 / 3</div> -->
            <img src="./public/assets/imgs/banner3.jpg" style="width:100%; height: 100%;">
            <!-- <div class="text">Caption Three</div> -->
        </div>

        <!-- Next and previous buttons -->
        <a class="prev" onclick="plusSlides(-1)">&#10094;</a>
        <a class="next" onclick="plusSlides(1)">&#10095;</a>
    </div>

    <div class="home_info" style="margin-top: 70px;">
        <h4 style="text-align: center; margin-bottom: 40px;">MY HOTEL</h4>
        <div class="home_info__inner" style="display: flex; justify-content: space-between;">
            <div class="home_info__card" style="width: 30%; display: flex; flex-direction: column; align-items: center; background-color: white; padding: 20px; border-radius: 10px; box-shadow: rgba(100, 100, 111, 0.2) 0px 7px 29px 0px;">
                <i style="font-size: 32px;" class="fa-solid fa-store"></i>
                <span style="margin-top: 12px; font-size: 20px; font-weight: 600;">Tiện ích</span>
                <p style="text-align: justify; margin-top: 10px;">Target elements that aren’t visible will be ignored and their corresponding nav items won’t receive an .active class. Scrollspy instances initialized in a non-visible wrapper will ignore all target elements.</p>
            </div>
            <div class="home_info__card" style="width: 30%; display: flex; flex-direction: column; align-items: center; background-color: white; padding: 20px; border-radius: 10px; box-shadow: rgba(100, 100, 111, 0.2) 0px 7px 29px 0px;">
                <i style="font-size: 32px;" class="fa-solid fa-star"></i>
                <span style="margin-top: 12px; font-size: 20px; font-weight: 600;">Sang trọng</span>
                <p style="text-align: justify; margin-top: 10px;">Target elements that aren’t visible will be ignored and their corresponding nav items won’t receive an .active class. Scrollspy instances initialized in a non-visible wrapper will ignore all target elements.</p>
            </div>
            <div class="home_info__card" style="width: 30%; display: flex; flex-direction: column; align-items: center; background-color: white; padding: 20px; border-radius: 10px; box-shadow: rgba(100, 100, 111, 0.2) 0px 7px 29px 0px;">
                <i style="font-size: 32px;" class="fa-solid fa-hotel"></i>
                <span style="margin-top: 12px; font-size: 20px; font-weight: 600;">Đẳng cấp</span>
                <p style="text-align: justify; margin-top: 10px;">Target elements that aren’t visible will be ignored and their corresponding nav items won’t receive an .active class. Scrollspy instances initialized in a non-visible wrapper will ignore all target elements.</p>
            </div>
        </div>
    </div>
</div>