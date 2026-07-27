(function () {
  var cfg = window.SITE_CONFIG;
  if (!cfg || !cfg.legal) return;

  function setText(id, text) {
    var el = document.getElementById(id);
    if (el) el.textContent = text;
  }

  var legal = cfg.legal;

  setText("legal-name", legal.name);
  setText("legal-rep", legal.representative);
  setText("legal-address", legal.address);
  setText("legal-phone", legal.phone);
  setText("legal-email", cfg.contactEmail);
  setText("legal-price", "¥" + cfg.priceYen.toLocaleString("ja-JP") + "（税込）");
  setText("legal-extra-fee", legal.extraFee);
  setText("legal-payment", legal.paymentMethod);
  setText("legal-payment-timing", legal.paymentTiming);
  setText("legal-delivery", legal.delivery);
  setText("legal-refund", legal.refund);

  setText("copyright", "© " + new Date().getFullYear() + " " + cfg.sellerName);
})();
