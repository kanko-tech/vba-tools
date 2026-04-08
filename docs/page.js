document.addEventListener("DOMContentLoaded", () => {
  const body = document.body;
  const toggle = document.querySelector("[data-toc-toggle]");
  const closeButtons = document.querySelectorAll("[data-toc-close]");
  const tocTargets = document.querySelectorAll("[data-toc-source]");
  const scrollTopButton = document.createElement("button");

  scrollTopButton.type = "button";
  scrollTopButton.className = "scroll-top-button";
  scrollTopButton.setAttribute("aria-label", "ページ上部へ戻る");
  scrollTopButton.textContent = "∧";

  const slugify = (text) =>
    text
      .toLowerCase()
      .trim()
      .replace(/[^\w\u3040-\u30ff\u3400-\u9fbf -]/g, "")
      .replace(/\s+/g, "-");

  const closeToc = () => {
    body.classList.remove("toc-open");
    if (toggle) toggle.setAttribute("aria-expanded", "false");
  };

  const updateScrollTopButton = () => {
    scrollTopButton.classList.toggle("is-visible", window.scrollY > 320);
  };

  if (toggle) {
    toggle.addEventListener("click", () => {
      const isOpen = body.classList.toggle("toc-open");
      toggle.setAttribute("aria-expanded", isOpen ? "true" : "false");
    });
  }

  closeButtons.forEach((button) => {
    button.addEventListener("click", closeToc);
  });

  document.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof Element)) return;
    if (!body.classList.contains("toc-open")) return;
    if (target.closest(".doc-sidebar")) return;
    if (target.closest("[data-toc-toggle]")) return;
    closeToc();
  });

  scrollTopButton.addEventListener("click", () => {
    window.scrollTo({ top: 0, behavior: "smooth" });
  });

  body.appendChild(scrollTopButton);
  updateScrollTopButton();
  window.addEventListener("scroll", updateScrollTopButton, { passive: true });

  tocTargets.forEach((tocRoot) => {
    const sourceSelector = tocRoot.getAttribute("data-toc-source");
    if (!sourceSelector) return;

    const contentRoot = document.querySelector(sourceSelector);
    if (!contentRoot) return;

    const headings = Array.from(contentRoot.querySelectorAll("h2, h3"));
    if (headings.length === 0) {
      tocRoot.innerHTML = "<p class=\"toc-empty\">No headings on this page.</p>";
      return;
    }

    const list = document.createElement("ul");
    list.className = "toc-list";

    headings.forEach((heading) => {
      if (!heading.id) {
        heading.id = slugify(heading.textContent || "section");
      }

      const item = document.createElement("li");
      item.className = heading.tagName.toLowerCase() === "h3" ? "toc-subitem" : "toc-item";

      const link = document.createElement("a");
      link.href = `#${heading.id}`;
      link.textContent = heading.textContent || "";
      link.addEventListener("click", closeToc);

      item.appendChild(link);
      list.appendChild(item);
    });

    tocRoot.innerHTML = "";
    tocRoot.appendChild(list);
  });
});
