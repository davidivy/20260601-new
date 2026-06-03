document.querySelectorAll("[data-scroll]").forEach((link) => {
  link.addEventListener("click", (event) => {
    const target = document.querySelector(link.getAttribute("href"));
    if (!target) return;
    event.preventDefault();
    target.scrollIntoView({ behavior: "smooth", block: "start" });
  });
});

document.querySelectorAll("[data-upload-panel]").forEach((panel) => {
  const fileInput = panel.querySelector("[data-upload-file]");
  const titleInput = panel.querySelector("[data-upload-title]");
  const tagInput = panel.querySelector("[data-upload-tag]");
  const bodyInput = panel.querySelector("[data-upload-body]");
  const preview = panel.querySelector("[data-upload-preview]");
  const submit = panel.querySelector("[data-upload-submit]");
  const status = panel.querySelector("[data-upload-status]");
  const list = panel.querySelector("[data-upload-list]");
  let selectedImage = "";

  fileInput.addEventListener("change", () => {
    const file = fileInput.files && fileInput.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      selectedImage = reader.result;
      preview.innerHTML = `<img src="${selectedImage}" alt="上傳照片預覽">`;
      status.textContent = "照片已選擇";
    };
    reader.readAsDataURL(file);
  });

  submit.addEventListener("click", () => {
    const title = titleInput.value.trim() || panel.dataset.uploadPanel;
    const tag = tagInput.value;
    const body = bodyInput.value.trim();

    if (!selectedImage && !body) {
      status.textContent = "請選擇照片或輸入文章內容";
      return;
    }

    const imageMarkup = selectedImage
      ? `<div class="entry-image"><img src="${selectedImage}" alt="${title} 上傳照片"></div>`
      : `<div class="entry-image text-entry-icon"><span class="upload-icon" aria-hidden="true"><svg viewBox="0 0 24 24"><path d="M4 5h16"></path><path d="M4 12h16"></path><path d="M4 19h10"></path></svg></span></div>`;

    const entry = document.createElement("article");
    entry.className = "entry-card";
    entry.innerHTML = `
      ${imageMarkup}
      <div>
        <span>${tag}</span>
        <h3>${title}</h3>
        <p>${body || "照片作品已加入頁面。"}</p>
      </div>
    `;
    list.prepend(entry);
    status.textContent = "已加入頁面";
    preview.textContent = "尚未選擇照片";
    fileInput.value = "";
    titleInput.value = "";
    bodyInput.value = "";
    selectedImage = "";
  });
});

document.querySelectorAll("[data-consult-form]").forEach((form) => {
  form.addEventListener("submit", () => {
    const status = form.querySelector("[data-consult-status]");
    const button = form.querySelector("button[type='submit']");
    if (status) status.textContent = "正在轉寄顧問委託訊息...";
    if (button) button.textContent = "轉寄中";
  });
});
