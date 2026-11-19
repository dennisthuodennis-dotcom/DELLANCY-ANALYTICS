
document.addEventListener("DOMContentLoaded", function () {

  // Smooth navigation
  const links = document.querySelectorAll('nav a');

  links.forEach(function(link) {
    link.addEventListener("click", function() {
      const target = document.querySelector(
        link.getAttribute("href")
      );

      if (target) {
        target.scrollIntoView({
          behavior: "smooth"
        });
      }
    });
  });

  // Highlight active navigation section
  const sections = document.querySelectorAll("section");

  window.addEventListener("scroll", function() {
    let current = "";

    sections.forEach(function(section) {
      const sectionTop = section.offsetTop - 120;

      if (window.scrollY >= sectionTop) {
        current = section.getAttribute("id");
      }
    });

    links.forEach(function(link) {
      link.style.color =
        link.getAttribute("href") === "#" + current
          ? "#7257dc"
          : "";
    });
  });

});
