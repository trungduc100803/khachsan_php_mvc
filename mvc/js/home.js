let slideIndex = 0;
showSlides();

function showSlides() {
  let i;
  let slides = document.getElementsByClassName("mySlides");
  for (i = 0; i < slides.length; i++) {
    slides[i].style.display = "none";
  }
  slideIndex++;
  if (slideIndex > slides.length) { slideIndex = 1 }
  slides[slideIndex - 1].style.display = "block";
  setTimeout(showSlides, 1500); // Change image every 2 seconds
}

let sl = 0
const spansl = document.querySelectorAll('.sl');
const btnplus = document.querySelector('.btn-plus')
btnplus.addEventListener('click', {})
spansl.forEach((e) => {
  e.innerHTML = sl;
})


const nav = document.querySelector('.nav-item')
console.log(nav)