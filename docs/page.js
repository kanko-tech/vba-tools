document.addEventListener("DOMContentLoaded", async () => {
  const targets = document.querySelectorAll("[data-markdown-source]");

  for (const target of targets) {
    const source = target.getAttribute("data-markdown-source");
    if (!source) continue;

    try {
      const response = await fetch(source);
      if (!response.ok) {
        throw new Error(`Failed to load ${source}`);
      }

      const markdown = await response.text();
      target.innerHTML = marked.parse(markdown, {
        gfm: true,
        breaks: true,
      });
    } catch (error) {
      target.innerHTML = `<p>詳細ドキュメントを読み込めませんでした: <code>${source}</code></p>`;
      console.error(error);
    }
  }
});
