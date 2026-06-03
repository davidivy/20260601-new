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
  const storageKey = `yeh-upload-${location.pathname}-${panel.dataset.uploadPanel}`;

  const iconMarkup = '<div class="entry-image text-entry-icon"><span class="upload-icon" aria-hidden="true"><svg viewBox="0 0 24 24"><path d="M4 5h16"></path><path d="M4 12h16"></path><path d="M4 19h10"></path></svg></span></div>';
  const escapeHtml = (value) => value.replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;"
  }[char]));

  const createEntry = ({ title, tag, body, image }) => {
    const safeTitle = escapeHtml(title);
    const safeTag = escapeHtml(tag);
    const safeBody = escapeHtml(body || "照片作品已加入頁面。");
    const imageMarkup = image
      ? `<div class="entry-image"><img src="${image}" alt="${safeTitle} 上傳照片"></div>`
      : iconMarkup;
    const entry = document.createElement("article");
    entry.className = "entry-card uploaded-entry";
    entry.innerHTML = `
      ${imageMarkup}
      <div>
        <span>${safeTag}</span>
        <h3>${safeTitle}</h3>
        <p>${safeBody}</p>
      </div>
    `;
    return entry;
  };

  const loadEntries = () => {
    try {
      const entries = JSON.parse(localStorage.getItem(storageKey) || "[]");
      entries.reverse().forEach((entry) => list.prepend(createEntry(entry)));
    } catch {
      localStorage.removeItem(storageKey);
    }
  };

  const saveEntry = (entry) => {
    try {
      const entries = JSON.parse(localStorage.getItem(storageKey) || "[]");
      entries.unshift(entry);
      localStorage.setItem(storageKey, JSON.stringify(entries.slice(0, 12)));
    } catch {
      status.textContent = "已加入頁面，但照片太大，瀏覽器無法保存。";
    }
  };

  loadEntries();

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
      ? selectedImage
      : "";
    const data = { title, tag, body, image: imageMarkup };
    const entry = createEntry(data);
    list.prepend(entry);
    saveEntry(data);
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
