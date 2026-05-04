// Footer year
document.getElementById('year').textContent = new Date().getFullYear();

// Mobile nav toggle
const hamburger = document.querySelector('.hamburger');
const navLinks = document.querySelector('.nav-links');
hamburger.addEventListener('click', () => {
  const open = navLinks.classList.toggle('open');
  hamburger.setAttribute('aria-expanded', open);
});

// Close mobile nav on link click
navLinks.querySelectorAll('a').forEach(a => {
  a.addEventListener('click', () => {
    navLinks.classList.remove('open');
    hamburger.setAttribute('aria-expanded', false);
  });
});

// Fade-in on scroll
const observer = new IntersectionObserver(
  entries => entries.forEach(e => {
    if (e.isIntersecting) {
      e.target.classList.add('visible');
      observer.unobserve(e.target);
    }
  }),
  { threshold: 0.1 }
);

document.querySelectorAll('.card, .contact-card, .skills-panel').forEach(el => {
  el.classList.add('fade-in');
  observer.observe(el);
});

// Inject fade-in keyframe via JS (avoids extra stylesheet)
const style = document.createElement('style');
style.textContent = `
  .fade-in { opacity: 0; transform: translateY(18px); transition: opacity 0.45s ease, transform 0.45s ease; }
  .fade-in.visible { opacity: 1; transform: translateY(0); }
`;
document.head.appendChild(style);
