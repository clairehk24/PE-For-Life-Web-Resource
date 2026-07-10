(function () {
  const buttons = Array.from(document.querySelectorAll("[data-accordion-target]"));

  buttons.forEach((button) => {
    const panel = document.getElementById(button.getAttribute("data-accordion-target"));
    if (!panel) return;

    button.addEventListener("click", () => {
      const shouldOpen = button.getAttribute("aria-expanded") !== "true";
      button.setAttribute("aria-expanded", String(shouldOpen));
      panel.hidden = !shouldOpen;
    });
  });
})();
