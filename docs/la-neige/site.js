const campaigns = {
  2: {
    label: "玫瑰情人節",
    title: "草莓玫瑰雙人冰",
    body: "IG 主打粉色桌拍，TikTok 拍淋醬與雙湯匙分食，FB 推雙人套餐 299 元。",
    image: "assets/month-rose.svg",
    alt: "玫瑰情人節雪花冰插畫"
  },
  4: {
    label: "伯爵茶香季",
    title: "焦糖伯爵雪花冰",
    body: "短影音拍茶湯、奶霜與雪花冰落下，FB 推平日午後茶冰組合。",
    image: "assets/month-earlgrey.svg",
    alt: "伯爵茶香雪花冰插畫"
  },
  7: {
    label: "芒果左岸",
    title: "芒果香草雪花冰",
    body: "主視覺用芒果金，TikTok 拍切芒果與限量倒數，IG 做夏日九宮格。",
    image: "assets/month-mango.svg",
    alt: "芒果雪花冰插畫"
  },
  12: {
    label: "聖誕雪花季",
    title: "聖誕草莓雪花冰",
    body: "搭配杯套、小物與交換禮物，FB 推四人套餐，IG 收集顧客聖誕合照。",
    image: "assets/month-christmas.svg",
    alt: "聖誕雪花季插畫"
  }
};

document.querySelectorAll(".month").forEach((button) => {
  button.addEventListener("click", () => {
    document.querySelectorAll(".month").forEach((item) => item.classList.remove("active"));
    button.classList.add("active");
    const data = campaigns[button.dataset.month];
    document.querySelector("#campaign-card").innerHTML = `
      <img class="campaign-illustration" src="${data.image}" alt="${data.alt}">
      <div>
        <p class="campaign-title">${data.label}</p>
        <h3>${data.title}</h3>
        <p>${data.body}</p>
      </div>
    `;
  });
});

const photoInput = document.querySelector("#photo-input");
const photoPreview = document.querySelector("#photo-preview");
let selectedPhoto = "";

photoInput.addEventListener("change", () => {
  const file = photoInput.files && photoInput.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    selectedPhoto = reader.result;
    photoPreview.innerHTML = `<img src="${selectedPhoto}" alt="上傳照片預覽">`;
  };
  reader.readAsDataURL(file);
});

document.querySelector("#add-photo").addEventListener("click", () => {
  if (!selectedPhoto) {
    photoPreview.textContent = "請先選擇一張照片";
    return;
  }
  const name = document.querySelector("#photo-name").value.trim() || "今日客人";
  const tag = document.querySelector("#photo-tag").value;
  const figure = document.createElement("figure");
  figure.innerHTML = `
    <img src="${selectedPhoto}" alt="${name} 上傳的 ${tag} 照片">
    <figcaption>${name}｜${tag}</figcaption>
  `;
  document.querySelector("#gallery").prepend(figure);
  photoPreview.textContent = "照片已加入拍照牆";
  photoInput.value = "";
  selectedPhoto = "";
});

document.querySelectorAll(".vote").forEach((button) => {
  button.addEventListener("click", () => {
    document.querySelectorAll(".vote").forEach((item) => item.classList.remove("active"));
    button.classList.add("active");
    document.querySelector("#vote-result").textContent = `已收到你的選擇：${button.dataset.flavor}。`;
  });
});

document.querySelector("#add-message").addEventListener("click", () => {
  const textarea = document.querySelector("#guest-message");
  const text = textarea.value.trim();
  if (!text) return;
  const li = document.createElement("li");
  li.textContent = `「${text}」`;
  document.querySelector("#message-list").prepend(li);
  textarea.value = "";
});

let selectedPlan = "雪室集點卡";

document.querySelectorAll(".join-plan").forEach((button) => {
  button.addEventListener("click", () => {
    selectedPlan = button.dataset.plan;
    document.querySelector("#member-status").textContent = `已選擇：${selectedPlan}`;
    document.querySelector("#member-name").focus();
  });
});

document.querySelector("#member-form").addEventListener("submit", (event) => {
  event.preventDefault();
  const name = document.querySelector("#member-name").value.trim();
  const contact = document.querySelector("#member-contact").value.trim();
  if (!name || !contact) return;
  document.querySelector("#member-status").textContent = `${name}，已為你保留「${selectedPlan}」會員方案。正式上線後可接 LINE OA、Email 或 POS 會員系統。`;
  event.currentTarget.reset();
});

const hikePhotoInput = document.querySelector("#hike-photo-input");
const hikePreview = document.querySelector("#hike-preview");
const ashanTask = document.querySelector("#ashan-task");
let selectedHikePhoto = "";

ashanTask.addEventListener("change", () => {
  const isBike = ashanTask.value === "bike";
  document.querySelector("#hike-name").placeholder = isBike ? "重機品牌車款，例如：Suzuki V-Strom" : "山名或步道，例如：陽明山、象山";
  document.querySelector("#bike-intro").placeholder = isBike ? "請介紹品牌、車款、騎乘路線或你喜歡它的原因。" : "登山任務可選填：路線心得、天氣、下山後想吃哪一碗冰。";
  document.querySelector("#hike-status").textContent = isBike ? "重機任務需 PO 騎士照，並填寫品牌車款介紹。" : "登山任務需 PO 爬山照，並填寫山名或步道。";
});

hikePhotoInput.addEventListener("change", () => {
  const file = hikePhotoInput.files && hikePhotoInput.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    selectedHikePhoto = reader.result;
    hikePreview.innerHTML = `<img src="${selectedHikePhoto}" alt="爬山照預覽">`;
  };
  reader.readAsDataURL(file);
});

document.querySelector("#hike-checkin").addEventListener("click", () => {
  const task = ashanTask.value;
  const target = document.querySelector("#hike-name").value.trim();
  const intro = document.querySelector("#bike-intro").value.trim();
  const status = document.querySelector("#hike-status");
  if (!selectedHikePhoto) {
    status.textContent = task === "bike" ? "請先選擇本月騎士重機照，才能領取阿山哥優惠。" : "請先選擇本月爬山照，才能領取阿山哥優惠。";
    return;
  }
  if (!target) {
    status.textContent = task === "bike" ? "請填寫重機品牌或車款，讓這次登錄有記憶點。" : "請填寫山名或步道名，讓這次登錄有故事。";
    return;
  }
  if (task === "bike" && intro.length < 8) {
    status.textContent = "重機任務請補一段品牌介紹或騎乘心得，至少 8 個字。";
    return;
  }
  const month = String(new Date().getMonth() + 1).padStart(2, "0");
  const code = task === "bike" ? `RIDER${month}` : `ASHAN${month}`;
  const label = task === "bike" ? `重機登錄完成：${target}` : `登山登錄完成：${target}`;
  status.textContent = `${label}。優惠碼 ${code}，可折抵招牌冰 20 元或冬季熱飲 9 折。`;
  document.querySelector("#hike-name").value = "";
  document.querySelector("#bike-intro").value = "";
  hikePhotoInput.value = "";
  selectedHikePhoto = "";
});
