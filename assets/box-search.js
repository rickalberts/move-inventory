(() => {
  const input = document.querySelector("#box-search");
  const resultCount = document.querySelector("#result-count");
  const emptyResults = document.querySelector("#empty-results");
  const categories = [...document.querySelectorAll(".inventory-category")];

  if (!input || !resultCount || !emptyResults || categories.length === 0) return;

  const normalize = (value) => value.trim().toLocaleLowerCase();

  input.addEventListener("input", () => {
    const query = normalize(input.value);
    let visibleItems = 0;

    categories.forEach((category) => {
      const categoryMatches = category.dataset.categorySearch.includes(query);
      const items = [...category.querySelectorAll(".inventory-item")];

      items.forEach((item) => {
        const matches = categoryMatches || item.dataset.itemSearch.includes(query);
        item.hidden = !matches;
        if (matches) visibleItems += 1;
      });

      category.hidden = visibleItems === 0 || !items.some((item) => !item.hidden);
    });

    resultCount.textContent = `${visibleItems} ${visibleItems === 1 ? "item" : "items"} shown`;
    emptyResults.hidden = visibleItems !== 0;
  });
})();
