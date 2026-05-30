const statuses = ["受付", "確認中", "処理中", "完了"];

const sampleTasks = [
  {
    id: createId(),
    title: "採用費と広告費の部署別集計",
    summary: "取締役会向けに4月実績と前年差分を集計。金曜午前が期限。",
    department: "経理",
    requester: "田中",
    priority: "高",
    status: "処理中",
    assignee: "経理 Ops",
    due: "2営業日以内",
    category: "レポート作成",
    nextAction: "会計CSVと広告管理画面の出力条件を確認",
    createdAt: Date.now() - 1000 * 60 * 60 * 48,
    completedAt: null
  },
  {
    id: createId(),
    title: "入社手続きチェックリスト更新",
    summary: "6月入社メンバー向けに必要書類とアカウント発行状況を整理。",
    department: "人事",
    requester: "佐藤",
    priority: "中",
    status: "確認中",
    assignee: "HR Ops",
    due: "今週中",
    category: "入退社",
    nextAction: "未提出書類とSlack招待状況を突合",
    createdAt: Date.now() - 1000 * 60 * 60 * 20,
    completedAt: null
  },
  {
    id: createId(),
    title: "契約書レビュー依頼",
    summary: "新規SaaS契約の支払条件と自動更新条項を確認。",
    department: "法務",
    requester: "山本",
    priority: "高",
    status: "受付",
    assignee: "Legal Ops",
    due: "明日まで",
    category: "契約確認",
    nextAction: "契約期間、解約通知期限、個人情報条項を抽出",
    createdAt: Date.now() - 1000 * 60 * 60 * 5,
    completedAt: null
  },
  {
    id: createId(),
    title: "備品購入申請の月次集計",
    summary: "部署別の申請件数と承認待ち金額を月次で集計。",
    department: "総務",
    requester: "鈴木",
    priority: "低",
    status: "完了",
    assignee: "Admin Ops",
    due: "完了済み",
    category: "購買管理",
    nextAction: "翌月からフォーム入力を必須化",
    createdAt: Date.now() - 1000 * 60 * 60 * 96,
    completedAt: Date.now() - 1000 * 60 * 60 * 20
  }
];

const sampleOutlookMessages = [
  {
    from: "finance-request@example.com",
    subject: "依頼: 月次経費の差分集計",
    body: "来週火曜の部門長会議までに、4月と5月の経費を部署別に比較して前年差分も出してください。金額が大きい項目はコメントが必要です。",
    department: "経理",
    requester: "小林"
  },
  {
    from: "hr-onboarding@example.com",
    subject: "申請: 入社準備チェック",
    body: "6月入社の業務委託メンバー3名について、アカウント発行、NDA、備品発送の状況を確認したいです。今週中に一覧化をお願いします。",
    department: "人事",
    requester: "佐藤"
  }
];

let tasks = loadTasks();
let microsoft365View = "overview";
let lastPlannerPayload = [];

const elements = {
  openCount: document.querySelector("#openCount"),
  urgentCount: document.querySelector("#urgentCount"),
  leadTime: document.querySelector("#leadTime"),
  automationCount: document.querySelector("#automationCount"),
  requestText: document.querySelector("#requestText"),
  requestDept: document.querySelector("#requestDept"),
  requester: document.querySelector("#requester"),
  analyzeRequest: document.querySelector("#analyzeRequest"),
  loadExample: document.querySelector("#loadExample"),
  aiResult: document.querySelector("#aiResult"),
  aiStatus: document.querySelector("#aiStatus"),
  searchTasks: document.querySelector("#searchTasks"),
  deptFilter: document.querySelector("#deptFilter"),
  resetData: document.querySelector("#resetData"),
  exportReport: document.querySelector("#exportReport"),
  insightList: document.querySelector("#insightList"),
  automationGrid: document.querySelector("#automationGrid"),
  m365Status: document.querySelector("#m365Status"),
  m365ClientId: document.querySelector("#m365ClientId"),
  m365PlanId: document.querySelector("#m365PlanId"),
  m365MailQuery: document.querySelector("#m365MailQuery"),
  m365ImportMail: document.querySelector("#m365ImportMail"),
  m365BuildPlanner: document.querySelector("#m365BuildPlanner"),
  m365Result: document.querySelector("#m365Result"),
  taskTemplate: document.querySelector("#taskTemplate")
};

function loadTasks() {
  const saved = localStorage.getItem("opspilot-tasks");
  return saved ? JSON.parse(saved) : sampleTasks;
}

function createId() {
  if (window.crypto && typeof window.crypto.randomUUID === "function") {
    return window.crypto.randomUUID();
  }
  return `task-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function saveTasks() {
  localStorage.setItem("opspilot-tasks", JSON.stringify(tasks));
}

function analyzeText(text, department, requester) {
  const urgentWords = ["至急", "明日", "今日", "午前", "締切", "期限", "取締役会", "監査", "契約", "金額"];
  const mediumWords = ["今週", "確認", "集計", "更新", "申請", "承認"];
  const lowerText = text.toLowerCase();
  const urgentScore = urgentWords.filter((word) => text.includes(word)).length;
  const mediumScore = mediumWords.filter((word) => text.includes(word)).length;
  const priority = urgentScore >= 2 ? "高" : mediumScore >= 2 ? "中" : "低";
  const category = inferCategory(text, department);
  const title = createTitle(text, category);
  const assignee = {
    経理: "経理 Ops",
    人事: "HR Ops",
    総務: "Admin Ops",
    法務: "Legal Ops",
    営業管理: "Sales Ops"
  }[department];

  return {
    id: createId(),
    title,
    summary: text.slice(0, 86) + (text.length > 86 ? "..." : ""),
    department,
    requester,
    priority,
    status: "受付",
    assignee,
    due: priority === "高" ? "1-2営業日以内" : priority === "中" ? "今週中" : "来週以降",
    category,
    nextAction: suggestNextAction(category, lowerText),
    createdAt: Date.now(),
    completedAt: null
  };
}

function inferCategory(text, department) {
  if (text.includes("契約") || text.includes("規約")) return "契約確認";
  if (text.includes("採用") || text.includes("入社") || text.includes("退職")) return "入退社";
  if (text.includes("請求") || text.includes("費") || text.includes("経費") || text.includes("集計")) return "レポート作成";
  if (text.includes("備品") || text.includes("購買") || text.includes("申請")) return "購買管理";
  return `${department}問い合わせ`;
}

function createTitle(text, category) {
  const compact = text.replace(/\s+/g, "").slice(0, 22);
  return compact ? `${category}: ${compact}` : `${category}の新規依頼`;
}

function suggestNextAction(category, lowerText) {
  const actions = {
    契約確認: "契約期間、支払条件、自動更新、個人情報条項を抽出",
    入退社: "対象者、入社日、必要アカウント、未回収書類を一覧化",
    レポート作成: "元データ、集計軸、提出形式、締切を確認",
    購買管理: "申請金額、承認者、予算科目、納期を確認"
  };
  if (lowerText.includes("slack") || lowerText.includes("スラック")) {
    return "チャンネル、対象者、権限範囲を確認";
  }
  return actions[category] || "依頼目的、期限、成果物、承認者を確認";
}

function render() {
  renderMetrics();
  renderBoard();
  renderInsights();
  renderMicrosoft365();
  renderAutomations();
}

function renderMetrics() {
  const open = tasks.filter((task) => task.status !== "完了").length;
  const urgent = tasks.filter((task) => task.priority === "高" && task.status !== "完了").length;
  const completed = tasks.filter((task) => task.completedAt);
  const averageLeadTime = completed.length
    ? completed.reduce((sum, task) => sum + (task.completedAt - task.createdAt), 0) / completed.length / 86400000
    : 0;
  const automationCandidates = new Set(tasks.map((task) => task.category)).size;

  elements.openCount.textContent = open;
  elements.urgentCount.textContent = urgent;
  elements.leadTime.textContent = `${averageLeadTime.toFixed(1)}日`;
  elements.automationCount.textContent = automationCandidates;
}

function renderBoard() {
  const search = elements.searchTasks.value.trim().toLowerCase();
  const department = elements.deptFilter.value;

  statuses.forEach((status) => {
    const list = document.querySelector(`#list-${status}`);
    list.innerHTML = "";
    const visibleTasks = tasks.filter((task) => {
      const matchesStatus = task.status === status;
      const matchesDept = department === "all" || task.department === department;
      const haystack = `${task.title} ${task.summary} ${task.requester} ${task.category}`.toLowerCase();
      return matchesStatus && matchesDept && haystack.includes(search);
    });

    if (!visibleTasks.length) {
      const empty = document.createElement("p");
      empty.className = "result-empty";
      empty.textContent = "該当タスクなし";
      list.append(empty);
      return;
    }

    visibleTasks.forEach((task) => list.append(createTaskCard(task)));
  });
}

function createTaskCard(task) {
  const card = elements.taskTemplate.content.firstElementChild.cloneNode(true);
  const priorityClass = task.priority === "高" ? "high" : task.priority === "中" ? "medium" : "low";
  card.classList.add(`is-${priorityClass}`);
  card.querySelector(".priority").className = `priority ${priorityClass}`;
  card.querySelector(".priority").textContent = `${task.priority}優先度`;
  card.querySelector("h4").textContent = task.title;
  card.querySelector("p").textContent = task.summary;
  card.querySelector(".task-meta").innerHTML = [
    task.department,
    task.assignee,
    task.due,
    task.category
  ].map((item) => `<span class="tag">${item}</span>`).join("");
  card.querySelector(".ghost-button").addEventListener("click", () => showTaskDetail(task));
  card.querySelector('[data-move="back"]').disabled = task.status === statuses[0];
  card.querySelector('[data-move="next"]').disabled = task.status === statuses[statuses.length - 1];
  card.querySelector('[data-move="back"]').addEventListener("click", () => moveTask(task.id, -1));
  card.querySelector('[data-move="next"]').addEventListener("click", () => moveTask(task.id, 1));
  return card;
}

function showTaskDetail(task) {
  elements.aiResult.innerHTML = `
    <dl>
      <dt>件名</dt><dd>${escapeHtml(task.title)}</dd>
      <dt>担当提案</dt><dd>${escapeHtml(task.assignee)}</dd>
      <dt>次アクション</dt><dd>${escapeHtml(task.nextAction)}</dd>
      <dt>業務分類</dt><dd>${escapeHtml(task.category)}</dd>
    </dl>
  `;
  elements.aiStatus.textContent = "Selected";
}

function moveTask(taskId, direction) {
  tasks = tasks.map((task) => {
    if (task.id !== taskId) return task;
    const nextIndex = Math.max(0, Math.min(statuses.length - 1, statuses.indexOf(task.status) + direction));
    const nextStatus = statuses[nextIndex];
    return {
      ...task,
      status: nextStatus,
      completedAt: nextStatus === "完了" ? Date.now() : null
    };
  });
  saveTasks();
  render();
}

function renderInsights() {
  const urgentDepartments = countBy(tasks.filter((task) => task.priority === "高"), "department");
  const topDept = Object.entries(urgentDepartments).sort((a, b) => b[1] - a[1])[0]?.[0] || "経理";
  const repeated = Object.entries(countBy(tasks, "category")).sort((a, b) => b[1] - a[1])[0]?.[0] || "レポート作成";
  const blocked = tasks.filter((task) => task.status === "確認中").length;
  const insights = [
    `${topDept}の高優先度依頼が目立ちます。締切と承認者の入力を受付フォームで必須化すると初動が速くなります。`,
    `${repeated}は標準テンプレート化の効果が高い業務です。必要項目チェックリストと出力形式を固定できます。`,
    `確認中タスクは${blocked}件です。AIが不足情報を先に質問文へ変換すると差し戻しを減らせます。`
  ];
  elements.insightList.innerHTML = insights.map((item) => `<li>${escapeHtml(item)}</li>`).join("");
}

function renderMicrosoft365() {
  const connected = Boolean(elements.m365ClientId.value.trim());
  elements.m365Status.textContent = connected ? "Ready for Graph" : "Demo mode";

  if (microsoft365View === "graph") {
    elements.m365Result.textContent = buildGraphGuide();
    return;
  }

  if (microsoft365View === "payload") {
    elements.m365Result.textContent = lastPlannerPayload.length
      ? JSON.stringify(lastPlannerPayload, null, 2)
      : "Planner登録案はまだ作成されていません。未完了タスクをもとに作成できます。";
    return;
  }

  elements.m365Result.textContent = [
    "Microsoft 365連携デモ",
    "",
    "1. Outlookの依頼メールを読み込み、AI Intakeと同じロジックでタスク化します。",
    "2. Plannerへ登録するためのGraph APIペイロードを生成します。",
    "3. 実接続時はMSAL.jsでサインインし、Graph APIへアクセストークン付きでリクエストします。",
    "",
    `現在の検索条件: ${elements.m365MailQuery.value || "未設定"}`,
    `Client ID: ${connected ? elements.m365ClientId.value.trim() : "未設定"}`,
    `Plan ID: ${elements.m365PlanId.value.trim() || "未設定"}`
  ].join("\n");
}

function buildGraphGuide() {
  const clientId = elements.m365ClientId.value.trim() || "<YOUR_ENTRA_CLIENT_ID>";
  const planId = elements.m365PlanId.value.trim() || "<PLANNER_PLAN_ID>";
  return [
    "Graph API連携メモ",
    "",
    "必要な代表スコープ:",
    "- Mail.Read: Outlookメールの読み込み",
    "- Tasks.ReadWrite: To Do / Planner系タスク操作",
    "- Group.ReadWrite.All: Plannerプランやグループ配下の操作",
    "- User.Read: サインインユーザー情報の取得",
    "",
    "MSAL設定例:",
    `clientId: \"${clientId}\"`,
    "authority: \"https://login.microsoftonline.com/common\"",
    "redirectUri: window.location.origin + window.location.pathname",
    "",
    "Outlookメール取得:",
    `GET https://graph.microsoft.com/v1.0/me/messages?$search=\"${elements.m365MailQuery.value || "依頼"}\"`,
    "",
    "Plannerタスク作成:",
    "POST https://graph.microsoft.com/v1.0/planner/tasks",
    `planId: \"${planId}\"`,
    "",
    "このデモでは認証情報を保存せず、実接続前の業務フロー確認に使えるようにしています。"
  ].join("\n");
}

function importOutlookMessages() {
  const importedTasks = sampleOutlookMessages.map((message) => {
    const task = analyzeText(message.body, message.department, message.requester);
    return {
      ...task,
      title: `Outlook: ${message.subject.replace(/^依頼: |^申請: /, "")}`,
      summary: `${message.from} からのメール: ${task.summary}`,
      category: task.category,
      nextAction: `${task.nextAction}。元メールの件名と送信者を確認`
    };
  });
  tasks = [...importedTasks, ...tasks];
  saveTasks();
  microsoft365View = "overview";
  setActiveM365Tab("overview");
  elements.m365Result.textContent = `${importedTasks.length}件のOutlookメールをタスク化しました。\n\n${importedTasks.map((task) => `- ${task.title} / ${task.priority}優先度 / ${task.assignee}`).join("\n")}`;
  renderMetrics();
  renderBoard();
  renderInsights();
}

function buildPlannerPayload() {
  const planId = elements.m365PlanId.value.trim() || "<PLANNER_PLAN_ID>";
  lastPlannerPayload = tasks
    .filter((task) => task.status !== "完了")
    .slice(0, 8)
    .map((task) => ({
      planId,
      title: task.title,
      assignments: {},
      priority: task.priority === "高" ? 1 : task.priority === "中" ? 5 : 9,
      details: {
        description: [
          task.summary,
          "",
          `部門: ${task.department}`,
          `担当提案: ${task.assignee}`,
          `期限: ${task.due}`,
          `次アクション: ${task.nextAction}`
        ].join("\n")
      }
    }));
  microsoft365View = "payload";
  setActiveM365Tab("payload");
  renderMicrosoft365();
}

function setActiveM365Tab(tabName) {
  document.querySelectorAll("[data-m365-tab]").forEach((button) => {
    button.classList.toggle("active", button.dataset.m365Tab === tabName);
  });
}

function renderAutomations() {
  const cards = [
    {
      title: "依頼分類ルール",
      body: "本文から部門、業務分類、緊急度を推定し、担当チームへ自動振り分け。"
    },
    {
      title: "不足情報チェック",
      body: "期限、成果物、承認者、対象データが抜けている場合に確認質問を自動生成。"
    },
    {
      title: "月次レポート生成",
      body: "CSVを取り込み、部署別集計と前年差分を定型フォーマットで作成。"
    }
  ];
  elements.automationGrid.innerHTML = cards.map((card) => `
    <article class="automation-card">
      <strong>${card.title}</strong>
      <p>${card.body}</p>
    </article>
  `).join("");
}

function countBy(items, key) {
  return items.reduce((acc, item) => {
    acc[item[key]] = (acc[item[key]] || 0) + 1;
    return acc;
  }, {});
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;"
  })[char]);
}

elements.analyzeRequest.addEventListener("click", () => {
  const text = elements.requestText.value.trim();
  if (!text) {
    elements.aiStatus.textContent = "Need text";
    elements.aiResult.innerHTML = '<div class="result-empty">依頼本文を入力してください。</div>';
    return;
  }
  elements.aiStatus.textContent = "Analyzed";
  const task = analyzeText(text, elements.requestDept.value, elements.requester.value.trim() || "未設定");
  tasks = [task, ...tasks];
  saveTasks();
  showTaskDetail(task);
  render();
  elements.requestText.value = "";
});

elements.loadExample.addEventListener("click", () => {
  elements.requestText.value = "来週の経営会議向けに、4月の採用費と広告費を部署別に集計して、前年差分も出してほしいです。金曜午前までに必要です。元データは会計CSVと広告管理画面にあります。";
  elements.requestDept.value = "経理";
  elements.requester.value = "田中";
});

elements.searchTasks.addEventListener("input", renderBoard);
elements.deptFilter.addEventListener("change", renderBoard);
elements.m365ClientId.addEventListener("input", renderMicrosoft365);
elements.m365PlanId.addEventListener("input", renderMicrosoft365);
elements.m365MailQuery.addEventListener("input", renderMicrosoft365);
elements.m365ImportMail.addEventListener("click", importOutlookMessages);
elements.m365BuildPlanner.addEventListener("click", buildPlannerPayload);
document.querySelectorAll("[data-m365-tab]").forEach((button) => {
  button.addEventListener("click", () => {
    microsoft365View = button.dataset.m365Tab;
    setActiveM365Tab(microsoft365View);
    renderMicrosoft365();
  });
});

elements.resetData.addEventListener("click", () => {
  localStorage.removeItem("opspilot-tasks");
  tasks = sampleTasks.map((task) => ({ ...task, id: createId() }));
  saveTasks();
  lastPlannerPayload = [];
  microsoft365View = "overview";
  setActiveM365Tab("overview");
  elements.aiResult.innerHTML = '<div class="result-empty">サンプルデータに戻しました。</div>';
  elements.aiStatus.textContent = "Reset";
  render();
});

elements.exportReport.addEventListener("click", () => {
  const rows = [
    ["件名", "部門", "担当", "優先度", "ステータス", "期限", "次アクション"],
    ...tasks.map((task) => [task.title, task.department, task.assignee, task.priority, task.status, task.due, task.nextAction])
  ];
  const csv = rows.map((row) => row.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(",")).join("\n");
  const blob = new Blob([`\ufeff${csv}`], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "opspilot-report.csv";
  link.click();
  URL.revokeObjectURL(url);
});

render();
