const menuToggle = document.querySelector("[data-menu-toggle]");
const nav = document.querySelector("[data-nav]");
const header = document.querySelector("[data-header]");
const backToTop = document.querySelector("[data-back-to-top]");
const form = document.querySelector("#participationForm");
const formStatus = document.querySelector("[data-form-status]");

const closeMenu = () => {
  if (!menuToggle || !nav) return;
  menuToggle.setAttribute("aria-expanded", "false");
  menuToggle.setAttribute("aria-label", "Abrir menú");
  nav.classList.remove("is-open");
  document.body.classList.remove("nav-open");
};

if (menuToggle && nav) {
  menuToggle.addEventListener("click", () => {
    const isOpen = menuToggle.getAttribute("aria-expanded") === "true";
    menuToggle.setAttribute("aria-expanded", String(!isOpen));
    menuToggle.setAttribute("aria-label", isOpen ? "Abrir menú" : "Cerrar menú");
    nav.classList.toggle("is-open", !isOpen);
    document.body.classList.toggle("nav-open", !isOpen);
  });

  nav.querySelectorAll("a").forEach((link) => {
    link.addEventListener("click", closeMenu);
  });
}

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") closeMenu();
});

const revealElements = document.querySelectorAll(".reveal");

if ("IntersectionObserver" in window) {
  const observer = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) {
          entry.target.classList.add("is-visible");
          observer.unobserve(entry.target);
        }
      });
    },
    { threshold: 0.16 }
  );

  revealElements.forEach((element) => observer.observe(element));
} else {
  revealElements.forEach((element) => element.classList.add("is-visible"));
}

const updateScrollControls = () => {
  if (!backToTop) return;
  const isVisible = window.scrollY > 540;
  backToTop.classList.toggle("is-visible", isVisible);
};

window.addEventListener("scroll", updateScrollControls, { passive: true });
updateScrollControls();

if (backToTop) {
  backToTop.addEventListener("click", () => {
    window.scrollTo({ top: 0, behavior: "smooth" });
  });
}

if (header) {
  const setHeaderHeight = () => {
    document.documentElement.style.setProperty("--header-height", `${header.offsetHeight}px`);
  };

  setHeaderHeight();
  window.addEventListener("resize", setHeaderHeight);
}

const errorMessages = {
  name: "Ingresá un nombre o alias.",
  location: "Indicá provincia o localidad.",
  area: "Seleccioná un área de interés.",
  email: "Ingresá un correo electrónico válido.",
  message: "Contanos brevemente cómo querés participar.",
};

const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

const setFieldError = (field, message = "") => {
  const wrapper = field.closest(".form-field");
  const error = document.querySelector(`[data-error-for="${field.id}"]`);

  if (wrapper) wrapper.classList.toggle("has-error", Boolean(message));
  if (error) error.textContent = message;
  field.setAttribute("aria-invalid", String(Boolean(message)));
};

const validateField = (field) => {
  const value = field.value.trim();
  let message = "";

  if (!value) {
    message = errorMessages[field.id] || "Este campo es obligatorio.";
  } else if (field.type === "email" && !emailPattern.test(value)) {
    message = errorMessages.email;
  } else if (field.id === "message" && value.length < 8) {
    message = "El mensaje debe tener al menos 8 caracteres.";
  }

  setFieldError(field, message);
  return !message;
};

if (form) {
  const fields = Array.from(form.querySelectorAll("input, select, textarea"));

  fields.forEach((field) => {
    field.addEventListener("blur", () => validateField(field));
    field.addEventListener("input", () => {
      if (field.getAttribute("aria-invalid") === "true") validateField(field);
    });
  });

  form.addEventListener("submit", (event) => {
    event.preventDefault();

    const isValid = fields.every(validateField);
    if (!isValid) {
      if (formStatus) formStatus.textContent = "Revisá los campos marcados para continuar.";
      return;
    }

    if (formStatus) {
      formStatus.textContent =
        "Gracias por tu interés. Esta primera versión registra la intención de participación de manera simbólica. Próximamente se habilitarán canales formales de contacto.";
    }

    form.reset();
    fields.forEach((field) => setFieldError(field));
  });
}
