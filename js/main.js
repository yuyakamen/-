(function () {
  var cfg = window.SITE_CONFIG;
  if (!cfg) return;

  function setText(id, text) {
    var el = document.getElementById(id);
    if (el) el.textContent = text;
  }

  function setHref(id, href) {
    var el = document.getElementById(id);
    if (el) el.setAttribute("href", href);
  }

  document.title = cfg.guidebookTitle + "｜" + cfg.sellerName;

  setText("brand-name", cfg.sellerName);
  setHref("instagram-link", cfg.instagramUrl);
  setText("instagram-link", cfg.instagramHandle);
  setHref("instagram-link-2", cfg.instagramUrl);

  setText("guidebook-title", cfg.guidebookTitle);
  setText("guidebook-subtitle", cfg.guidebookSubtitle);
  setText("guidebook-description", cfg.guidebookDescription);

  var statusBadge = document.getElementById("status-badge");
  if (cfg.comingSoon) {
    statusBadge.textContent = "近日公開";
    statusBadge.classList.add("coming-soon");
  } else {
    statusBadge.textContent = "PDFガイドブック 販売中";
  }

  var contentsList = document.getElementById("contents-list");
  (cfg.contents || []).forEach(function (item) {
    var li = document.createElement("li");
    li.textContent = item;
    contentsList.appendChild(li);
  });

  setText("format-note", cfg.format + (cfg.pageCountNote ? "・" + cfg.pageCountNote : ""));
  setText(
    "price",
    "¥" + cfg.priceYen.toLocaleString("ja-JP") + " " + "（税込）"
  );

  var buyButton = document.getElementById("buy-button");
  var linkIsPlaceholder =
    !cfg.stripePaymentLink || cfg.stripePaymentLink.indexOf("xxxx") !== -1;

  if (cfg.comingSoon || linkIsPlaceholder) {
    buyButton.textContent = cfg.comingSoon ? "近日公開" : "準備中";
    buyButton.classList.add("disabled");
    buyButton.removeAttribute("href");
    buyButton.setAttribute("aria-disabled", "true");
  } else {
    buyButton.textContent = "PDFを購入する";
    buyButton.setAttribute("href", cfg.stripePaymentLink);
  }

  var faqList = document.getElementById("faq-list");
  (cfg.faq || []).forEach(function (item) {
    var details = document.createElement("details");
    details.className = "faq-item";
    var summary = document.createElement("summary");
    summary.textContent = item.q;
    var p = document.createElement("p");
    p.textContent = item.a;
    details.appendChild(summary);
    details.appendChild(p);
    faqList.appendChild(details);
  });

  setHref("contact-link", "mailto:" + cfg.contactEmail);

  setText(
    "copyright",
    "© " + new Date().getFullYear() + " " + cfg.sellerName
  );
})();
