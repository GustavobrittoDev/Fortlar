const leadStages = [
  { key: "novo", label: "Novo contato" },
  { key: "orcamento", label: "Orcamento enviado" },
  { key: "negociacao", label: "Negociacao" },
  { key: "fechado", label: "Fechado" },
  { key: "arquivado", label: "Arquivado" },
];

const orderStatuses = ["Agendado", "Em andamento", "Concluido", "Cancelado"];
const paymentStatuses = ["Pendente", "Parcial", "Pago"];

let supabaseClient = null;
let currentUser = null;
let searchQuery = "";
let selectedOrderId = null;

let state = {
  meta: {
    fortlarPhone: "5512996619062",
    storageBucket: "crm-anexos",
  },
  stats: {},
  leads: [],
  appointments: [],
  orders: [],
  activities: [],
};

const elements = {
  authShell: document.getElementById("auth-shell"),
  appShell: document.getElementById("app-shell"),
  loginForm: document.getElementById("login-form"),
  loginError: document.getElementById("login-error"),
  appFeedback: document.getElementById("app-feedback"),
  pageTitle: document.getElementById("page-title"),
  currentDate: document.getElementById("current-date"),
  loggedUserName: document.getElementById("logged-user-name"),
  statsGrid: document.getElementById("stats-grid"),
  pipelineOverview: document.getElementById("pipeline-overview"),
  todaySchedule: document.getElementById("today-schedule"),
  activityList: document.getElementById("activity-list"),
  pipelineBoard: document.getElementById("pipeline-board"),
  customersTable: document.getElementById("customers-table"),
  agendaGrid: document.getElementById("agenda-grid"),
  ordersTable: document.getElementById("orders-table"),
  orderDetailTitle: document.getElementById("order-detail-title"),
  orderDetailContent: document.getElementById("order-detail-content"),
  financeCards: document.getElementById("finance-cards"),
  receivablesTable: document.getElementById("receivables-table"),
  search: document.getElementById("global-search"),
  sidebar: document.getElementById("sidebar"),
  sidebarToggle: document.getElementById("sidebar-toggle"),
  pageLayer: document.getElementById("page-layer"),
  sidebarOpenLeads: document.getElementById("sidebar-open-leads"),
  sidebarTodayVisits: document.getElementById("sidebar-today-visits"),
  sidebarRevenue: document.getElementById("sidebar-revenue"),
  leadModal: document.getElementById("lead-modal"),
  appointmentModal: document.getElementById("appointment-modal"),
  orderModal: document.getElementById("order-modal"),
  leadForm: document.getElementById("lead-form"),
  appointmentForm: document.getElementById("appointment-form"),
  orderForm: document.getElementById("order-form"),
  logoutButton: document.getElementById("logout-button"),
};

initialize();

async function initialize() {
  elements.currentDate.textContent = new Intl.DateTimeFormat("pt-BR", {
    weekday: "long",
    day: "2-digit",
    month: "long",
  }).format(new Date());

  bindNavigation();
  bindModals();
  bindForms();
  bindSearch();
  bindSidebar();
  bindActionDelegation();

  try {
    supabaseClient = createSupabaseClient();
  } catch (error) {
    elements.loginError.textContent = error.message;
    elements.loginError.hidden = false;
    showLogin();
    return;
  }

  supabaseClient.auth.onAuthStateChange((_event, session) => {
    currentUser = session?.user || null;
    if (!session) {
      showLogin();
    }
  });

  bindAuthentication();

  const {
    data: { session },
  } = await supabaseClient.auth.getSession();

  if (!session) {
    showLogin();
    return;
  }

  currentUser = session.user;
  try {
    await bootstrap();
    showApp();
  } catch (error) {
    showLogin();
    elements.loginError.textContent = translateError(error.message);
    elements.loginError.hidden = false;
  }
}

function createSupabaseClient() {
  const config = window.FORTLAR_SUPABASE_CONFIG || {};

  if (!config.supabaseUrl || !config.supabaseAnonKey) {
    throw new Error("Configure crm/config.js com a URL e a anon key do Supabase.");
  }

  state.meta.storageBucket = config.storageBucket || "crm-anexos";
  return window.supabase.createClient(config.supabaseUrl, config.supabaseAnonKey);
}

function bindAuthentication() {
  elements.loginForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    elements.loginError.hidden = true;

    const form = new FormData(elements.loginForm);
    const email = String(form.get("email") || "").trim();
    const password = String(form.get("password") || "");

    const { data, error } = await supabaseClient.auth.signInWithPassword({ email, password });

    if (error) {
      elements.loginError.textContent = translateError(error.message);
      elements.loginError.hidden = false;
      return;
    }

    currentUser = data.user;
    try {
      await bootstrap();
      showApp();
    } catch (bootstrapError) {
      elements.loginError.textContent = translateError(bootstrapError.message);
      elements.loginError.hidden = false;
    }
  });

  elements.logoutButton.addEventListener("click", async () => {
    await supabaseClient.auth.signOut();
    currentUser = null;
    showLogin();
  });
}

async function bootstrap() {
  const [profileResult, leadsResult, appointmentsResult, ordersResult, activitiesResult] = await Promise.all([
    supabaseClient.from("profiles").select("name, role").eq("id", currentUser.id).maybeSingle(),
    supabaseClient.from("leads").select("*").order("created_at", { ascending: false }),
    supabaseClient.from("appointments").select("*").order("date", { ascending: true }).order("time", { ascending: true }),
    supabaseClient.from("orders").select("*").order("date", { ascending: false }).order("time", { ascending: false }),
    supabaseClient.from("activities").select("*").order("created_at", { ascending: false }).limit(12),
  ]);

  throwIfError(profileResult.error);
  throwIfError(leadsResult.error);
  throwIfError(appointmentsResult.error);
  throwIfError(ordersResult.error);
  throwIfError(activitiesResult.error);

  currentUser.profile = profileResult.data || {};

  const attachmentsByOrder = await fetchAttachmentCounts();

  state.leads = (leadsResult.data || []).map(mapLead);
  state.appointments = (appointmentsResult.data || []).map(mapAppointment);
  state.orders = (ordersResult.data || []).map((row) => mapOrder(row, attachmentsByOrder));
  state.activities = (activitiesResult.data || []).map(mapActivity);
  state.stats = buildStats();

  if (!selectedOrderId && state.orders.length) {
    selectedOrderId = state.orders[0].id;
  }

  renderAll();
  await renderSelectedOrder();
}

async function fetchAttachmentCounts() {
  const { data, error } = await supabaseClient.from("attachments").select("order_id");
  throwIfError(error);

  return (data || []).reduce((accumulator, item) => {
    accumulator[item.order_id] = (accumulator[item.order_id] || 0) + 1;
    return accumulator;
  }, {});
}

function buildStats() {
  return {
    openLeads: state.leads.filter((lead) => !["fechado", "arquivado"].includes(lead.status)).length,
    wonLeads: state.leads.filter((lead) => lead.status === "fechado").length,
    activeOrders: state.orders.filter((order) => ["Agendado", "Em andamento"].includes(order.status)).length,
    pendingReceivables: state.orders
      .filter((order) => order.paymentStatus !== "Pago")
      .reduce((sum, order) => sum + Number(order.amount || 0), 0),
    scheduledToday: state.appointments.filter((item) => item.date === todayIso()).length,
  };
}

function bindNavigation() {
  const navButtons = document.querySelectorAll(".nav-item");

  navButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const section = button.dataset.section;

      navButtons.forEach((item) => item.classList.toggle("is-active", item === button));
      document.querySelectorAll(".page-section").forEach((panel) => {
        panel.classList.toggle("is-active", panel.id === section);
      });

      elements.pageTitle.textContent = getSectionTitle(section);
      closeSidebar();
    });
  });
}

function bindModals() {
  document.querySelectorAll("[data-open-modal]").forEach((button) => {
    button.addEventListener("click", () => {
      const modal = document.getElementById(button.dataset.openModal);
      modal?.showModal();
    });
  });

  document.querySelectorAll("[data-close-modal]").forEach((button) => {
    button.addEventListener("click", () => {
      button.closest("dialog")?.close();
    });
  });
}

function bindForms() {
  elements.leadForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    clearFeedback();

    try {
      const form = new FormData(elements.leadForm);

      const { error } = await supabaseClient.from("leads").insert({
        name: form.get("name"),
        phone: form.get("phone"),
        service: form.get("service"),
        address: form.get("address"),
        source: form.get("source"),
        priority: form.get("priority"),
        status: "novo",
        notes: form.get("notes"),
        last_contact: "Agora",
      });

      throwIfError(error);
      await insertActivity("Novo lead cadastrado", `${form.get("name")} entrou para ${form.get("service")}.`);
      elements.leadForm.reset();
      elements.leadModal.close();
      await bootstrap();
      showFeedback("Lead salvo com sucesso.", "success");
    } catch (error) {
      showFeedback(translateError(error.message), "error");
    }
  });

  elements.appointmentForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    clearFeedback();

    try {
      const form = new FormData(elements.appointmentForm);

      const { error } = await supabaseClient.from("appointments").insert({
        customer: form.get("customer"),
        phone: form.get("phone"),
        service: form.get("service"),
        address: form.get("address"),
        date: form.get("date"),
        time: form.get("time"),
        notes: form.get("notes"),
      });

      throwIfError(error);
      await insertActivity("Visita agendada", `${form.get("customer")} foi agendado para ${form.get("date")} as ${form.get("time")}.`);
      elements.appointmentForm.reset();
      elements.appointmentModal.close();
      await bootstrap();
      showFeedback("Visita agendada com sucesso.", "success");
    } catch (error) {
      showFeedback(translateError(error.message), "error");
    }
  });

  elements.orderForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    clearFeedback();

    try {
      const form = new FormData(elements.orderForm);

      const { data, error } = await supabaseClient
        .from("orders")
        .insert({
          customer: form.get("customer"),
          phone: form.get("phone"),
          service: form.get("service"),
          address: form.get("address"),
          date: form.get("date"),
          time: form.get("time"),
          amount: Number(form.get("amount") || 0),
          status: form.get("status"),
          payment_status: form.get("paymentStatus"),
          notes: form.get("notes"),
        })
        .select("id")
        .single();

      throwIfError(error);
      await insertActivity("Nova OS criada", `${form.get("customer")} entrou em execucao para ${form.get("service")}.`);
      selectedOrderId = data.id;
      elements.orderForm.reset();
      elements.orderModal.close();
      await bootstrap();
      showFeedback("Ordem de servico criada com sucesso.", "success");
    } catch (error) {
      showFeedback(translateError(error.message), "error");
    }
  });
}

function bindSearch() {
  elements.search.addEventListener("input", (event) => {
    searchQuery = event.target.value.trim().toLowerCase();
    renderAll();
    renderSelectedOrder();
  });
}

function bindSidebar() {
  elements.sidebarToggle?.addEventListener("click", () => {
    elements.sidebar.classList.add("is-open");
    elements.pageLayer.classList.add("is-visible");
  });

  elements.pageLayer?.addEventListener("click", closeSidebar);
}

function bindActionDelegation() {
  document.addEventListener("click", async (event) => {
    const deleteLeadButton = event.target.closest("[data-delete-lead]");

    if (deleteLeadButton) {
      event.preventDefault();
      await deleteLead(deleteLeadButton.dataset.deleteLead);
      return;
    }

    const deleteAppointmentButton = event.target.closest("[data-delete-appointment]");

    if (deleteAppointmentButton) {
      event.preventDefault();
      await deleteAppointment(deleteAppointmentButton.dataset.deleteAppointment);
      return;
    }

    const deleteOrderButton = event.target.closest("[data-delete-order]");

    if (deleteOrderButton) {
      event.preventDefault();
      await deleteOrder(deleteOrderButton.dataset.deleteOrder);
      return;
    }

    const deleteAttachmentButton = event.target.closest("[data-delete-attachment]");

    if (deleteAttachmentButton) {
      event.preventDefault();
      await deleteAttachment(
        deleteAttachmentButton.dataset.deleteAttachment,
        deleteAttachmentButton.dataset.attachmentPath,
        deleteAttachmentButton.dataset.attachmentName
      );
    }
  });
}

function closeSidebar() {
  elements.sidebar.classList.remove("is-open");
  elements.pageLayer.classList.remove("is-visible");
}

function showApp() {
  elements.authShell.classList.add("is-hidden");
  elements.appShell.classList.remove("is-hidden");
  elements.loggedUserName.textContent = currentUser?.profile?.name || currentUser?.email || "Equipe Fort Lar";
}

function showLogin() {
  elements.authShell.classList.remove("is-hidden");
  elements.appShell.classList.add("is-hidden");
}

function showFeedback(message, type = "success") {
  if (!elements.appFeedback) {
    return;
  }

  elements.appFeedback.textContent = message;
  elements.appFeedback.classList.remove("is-hidden", "is-success", "is-error");
  elements.appFeedback.classList.add(type === "error" ? "is-error" : "is-success");
}

function clearFeedback() {
  if (!elements.appFeedback) {
    return;
  }

  elements.appFeedback.textContent = "";
  elements.appFeedback.classList.add("is-hidden");
  elements.appFeedback.classList.remove("is-success", "is-error");
}

function renderAll() {
  renderSidebarSummary();
  renderStats();
  renderDashboard();
  renderPipeline();
  renderCustomers();
  renderAgenda();
  renderOrders();
  renderFinance();
}

function renderSidebarSummary() {
  elements.sidebarOpenLeads.textContent = String(state.stats.openLeads || 0);
  elements.sidebarTodayVisits.textContent = String(state.stats.scheduledToday || 0);
  elements.sidebarRevenue.textContent = formatCurrency(
    state.orders.reduce((sum, order) => sum + Number(order.amount || 0), 0)
  );
}

function renderStats() {
  const stats = [
    { label: "Leads em aberto", value: state.stats.openLeads || 0, trend: "Oportunidades atuais", icon: "Funil" },
    { label: "Negocios fechados", value: state.stats.wonLeads || 0, trend: "Conversoes registradas", icon: "Meta" },
    { label: "OS ativas", value: state.stats.activeOrders || 0, trend: "Execucao em andamento", icon: "Campo" },
    {
      label: "Recebivel pendente",
      value: formatCurrency(state.stats.pendingReceivables || 0),
      trend: "Acompanhar cobranca",
      icon: "Caixa",
    },
  ];

  elements.statsGrid.innerHTML = stats
    .map(
      (stat) => `
        <article class="stat-card">
          <div class="stat-head">
            <span class="stat-label">${stat.label}</span>
            <span class="badge">${stat.icon}</span>
          </div>
          <strong class="stat-value">${stat.value}</strong>
          <span class="stat-trend">${stat.trend}</span>
        </article>
      `
    )
    .join("");
}

function renderDashboard() {
  elements.pipelineOverview.innerHTML = leadStages
    .map((stage) => {
      const leads = getFilteredLeads().filter((lead) => lead.status === stage.key);
      return `
        <div class="info-row">
          <div>
            <strong>${stage.label}</strong>
            <span class="muted">${leads.length} lead(s)</span>
          </div>
          <span class="column-count">${leads.length}</span>
        </div>
      `;
    })
    .join("");

  const today = getAppointmentsForDate(todayIso());
  const scheduleSource = today.length ? today : getUpcomingAppointments().slice(0, 4);

  elements.todaySchedule.innerHTML = scheduleSource.length
    ? scheduleSource
        .map(
          (item) => `
            <div class="schedule-item">
              <div>
                <strong>${item.customer}</strong>
                <span class="muted">${item.service} · ${item.address}</span>
              </div>
              <span class="badge">${formatDateTime(item.date, item.time)}</span>
            </div>
          `
        )
        .join("")
    : `<div class="empty-state">Nenhum compromisso agendado.</div>`;

  const activities = state.activities.filter((activity) => matchesSearch([activity.title, activity.description]));

  elements.activityList.innerHTML = activities.length
    ? activities
        .slice(0, 5)
        .map(
          (activity) => `
            <div class="finance-row">
              <div>
                <strong>${activity.title}</strong>
                <span class="muted">${activity.description}</span>
              </div>
              <span class="badge">${activity.time}</span>
            </div>
          `
        )
        .join("")
    : `<div class="empty-state">Nenhuma atividade encontrada.</div>`;
}

function renderPipeline() {
  elements.pipelineBoard.innerHTML = leadStages
    .map((stage) => {
      const leads = getFilteredLeads().filter((lead) => lead.status === stage.key);
      return `
        <article class="pipeline-column" data-stage="${stage.key}">
          <div class="column-head">
            <div class="column-head-copy">
              <strong>${stage.label}</strong>
              <span>${leads.length} lead(s) neste estagio</span>
            </div>
            <span class="column-count">${leads.length}</span>
          </div>
          <div class="lead-stack" data-dropzone="${stage.key}">
            ${
              leads.length
                ? leads
                    .map(
                      (lead) => `
                        <div class="lead-card" draggable="true" data-lead-id="${lead.id}">
                          <div class="lead-card-header">
                            <div class="lead-title-group">
                              <strong>${lead.name}</strong>
                              <span class="lead-id">${lead.id}</span>
                            </div>
                            <div class="lead-header-actions">
                              <span class="priority ${normalizeKey(lead.priority)}">${lead.priority}</span>
                              <button
                                class="action-icon-button danger-button"
                                type="button"
                                data-delete-lead="${lead.id}"
                                aria-label="Excluir lead ${escapeHtml(lead.name)}"
                                title="Excluir lead"
                              >
                                <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
                                  <path d="M4 7H20" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
                                  <path d="M9 7V5.5C9 4.67 9.67 4 10.5 4H13.5C14.33 4 15 4.67 15 5.5V7" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
                                  <path d="M7.5 9.5V17.5C7.5 18.33 8.17 19 9 19H15C15.83 19 16.5 18.33 16.5 17.5V9.5" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
                                  <path d="M10 11V16" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
                                  <path d="M14 11V16" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
                                </svg>
                              </button>
                            </div>
                          </div>
                          <div class="lead-card-body">
                            <div class="lead-meta-row">
                              <span class="badge">${lead.service}</span>
                              <span class="lead-source">${lead.source}</span>
                            </div>
                            <p class="lead-note">${lead.notes}</p>
                            <div class="lead-contact-list">
                              <span class="lead-contact-item">${lead.address}</span>
                              <span class="lead-contact-item">${lead.phone}</span>
                            </div>
                          </div>
                          <div class="lead-card-actions">
                            <select class="status-select" data-lead-status="${lead.id}">
                              ${leadStages
                                .map(
                                  (item) => `
                                    <option value="${item.key}" ${lead.status === item.key ? "selected" : ""}>${item.label}</option>
                                  `
                                )
                                .join("")}
                            </select>
                          </div>
                        </div>
                      `
                    )
                    .join("")
                : `<div class="empty-state">Sem leads neste estagio.</div>`
            }
          </div>
        </article>
      `;
    })
    .join("");

  bindPipelineInteractions();
}

function bindPipelineInteractions() {
  let draggedLeadId = null;

  document.querySelectorAll(".lead-card").forEach((card) => {
    card.addEventListener("dragstart", () => {
      draggedLeadId = card.dataset.leadId;
      card.classList.add("dragging");
    });

    card.addEventListener("dragend", () => {
      draggedLeadId = null;
      card.classList.remove("dragging");
    });
  });

  document.querySelectorAll("[data-dropzone]").forEach((zone) => {
    zone.addEventListener("dragover", (event) => event.preventDefault());
    zone.addEventListener("drop", async () => {
      if (!draggedLeadId) {
        return;
      }

      const { error } = await supabaseClient
        .from("leads")
        .update({ status: zone.dataset.dropzone, last_contact: "Atualizado agora" })
        .eq("id", draggedLeadId);

      throwIfError(error);
      await insertActivity("Lead atualizado", `Lead movido para ${zone.dataset.dropzone}.`);
      await bootstrap();
      showFeedback("Pipeline atualizado com sucesso.", "success");
    });
  });

  document.querySelectorAll("[data-lead-status]").forEach((select) => {
    select.addEventListener("change", async (event) => {
      const { error } = await supabaseClient
        .from("leads")
        .update({ status: event.target.value, last_contact: "Atualizado agora" })
        .eq("id", event.target.dataset.leadStatus);

      throwIfError(error);
      await insertActivity("Lead atualizado", `Lead movido para ${event.target.value}.`);
      await bootstrap();
      showFeedback("Lead atualizado com sucesso.", "success");
    });
  });
}

function renderCustomers() {
  const customers = buildCustomerRows();

  elements.customersTable.innerHTML = customers.length
    ? customers
        .map(
          (customer) => `
            <tr>
              <td class="table-cell-primary">
                <strong>${customer.name}</strong><br>
                <span class="table-subline">${customer.phone || "Sem telefone"}</span>
              </td>
              <td><span class="badge">${customer.service}</span></td>
              <td class="table-cell-address">${customer.address}</td>
              <td><span class="status-pill ${leadStatusClass(customer.status)}">${getLeadStageLabel(customer.status)}</span></td>
              <td>${customer.lastContact}</td>
              <td>
                <div class="table-actions">
                  <button class="ghost-button table-button danger-button" type="button" data-delete-lead="${customer.id}">
                    Excluir
                  </button>
                </div>
              </td>
            </tr>
          `
        )
        .join("")
    : `<tr><td colspan="6"><div class="empty-state">Nenhum cliente encontrado.</div></td></tr>`;
}

function renderAgenda() {
  const groupedDays = getAgendaDays();

  elements.agendaGrid.innerHTML = groupedDays
    .map(
      (day) => `
        <article class="agenda-day">
          <h3>${day.label}</h3>
          <div class="agenda-list">
            ${
              day.items.length
                ? day.items
                    .map(
                      (item) => `
                        <div class="schedule-item">
                          <div>
                            <strong>${item.customer}</strong>
                            <span class="muted">${item.service} · ${item.address}</span>
                          </div>
                          <div class="stack-actions schedule-actions">
                            <span class="badge">${item.time}</span>
                            <a class="text-link" href="${buildCalendarLink(item)}" target="_blank" rel="noreferrer">Google Calendar</a>
                            <button class="ghost-button table-button danger-button" type="button" data-delete-appointment="${item.id}">
                              Excluir
                            </button>
                          </div>
                        </div>
                      `
                    )
                    .join("")
                : `<div class="empty-state">Sem visitas agendadas.</div>`
            }
          </div>
        </article>
      `
    )
    .join("");
}

function renderOrders() {
  const orders = getFilteredOrders();

  if (!orders.find((item) => item.id === selectedOrderId)) {
    selectedOrderId = orders[0]?.id || null;
  }

  elements.ordersTable.innerHTML = orders.length
    ? orders
        .map(
          (order) => `
            <tr class="table-row-button ${order.id === selectedOrderId ? "is-selected" : ""}" data-order-row="${order.id}">
              <td class="table-cell-primary"><strong>${order.code}</strong></td>
              <td>${order.customer}</td>
              <td>${order.service}</td>
              <td>${formatDateTime(order.date, order.time)}</td>
              <td>${formatCurrency(order.amount)}</td>
              <td>
                <select class="status-select" data-order-id="${order.id}">
                  ${orderStatuses
                    .map(
                      (status) => `
                        <option value="${status}" ${order.status === status ? "selected" : ""}>${status}</option>
                      `
                    )
                    .join("")}
                </select>
              </td>
              <td>${order.attachmentCount || 0}</td>
              <td>
                <div class="table-actions">
                  <button class="ghost-button table-button danger-button" type="button" data-delete-order="${order.id}">
                    Excluir
                  </button>
                </div>
              </td>
            </tr>
          `
        )
        .join("")
    : `<tr><td colspan="8"><div class="empty-state">Nenhuma OS encontrada.</div></td></tr>`;

  document.querySelectorAll("[data-order-id]").forEach((select) => {
    select.addEventListener("change", async (event) => {
      const { error } = await supabaseClient
        .from("orders")
        .update({ status: event.target.value })
        .eq("id", event.target.dataset.orderId);

      throwIfError(error);
      await insertActivity("OS atualizada", `A ordem foi movida para ${event.target.value}.`);
      await bootstrap();
      showFeedback("Status da OS atualizado.", "success");
    });
  });

  document.querySelectorAll("[data-order-row]").forEach((row) => {
    row.addEventListener("click", async (event) => {
      if (event.target.closest("button, select, a")) {
        return;
      }

      selectedOrderId = row.dataset.orderRow;
      renderOrders();
      await renderSelectedOrder();
    });
  });
}

async function renderSelectedOrder() {
  if (!selectedOrderId) {
    elements.orderDetailTitle.textContent = "Selecione uma ordem";
    elements.orderDetailContent.innerHTML = `<div class="empty-state">Selecione uma ordem para ver detalhes, anexos e acoes rapidas.</div>`;
    return;
  }

  const { data: orderRow, error } = await supabaseClient.from("orders").select("*").eq("id", selectedOrderId).single();
  throwIfError(error);

  const { data: attachmentsRows, error: attachmentsError } = await supabaseClient
    .from("attachments")
    .select("*")
    .eq("order_id", selectedOrderId)
    .order("created_at", { ascending: false });
  throwIfError(attachmentsError);

  const attachments = await Promise.all((attachmentsRows || []).map(createAttachmentViewModel));
  const order = mapOrder(orderRow, { [selectedOrderId]: attachments.length });

  elements.orderDetailTitle.textContent = `${order.code} · ${order.customer}`;
  elements.orderDetailContent.innerHTML = `
    <div class="detail-block">
      <div class="detail-grid">
        <div>
          <span class="detail-label">Servico</span>
          <strong>${order.service}</strong>
        </div>
        <div>
          <span class="detail-label">Agendamento</span>
          <strong>${formatDateTime(order.date, order.time)}</strong>
        </div>
        <div>
          <span class="detail-label">Pagamento</span>
          <strong>${order.paymentStatus}</strong>
        </div>
        <div>
          <span class="detail-label">Valor</span>
          <strong>${formatCurrency(order.amount)}</strong>
        </div>
      </div>
    </div>

    <div class="detail-block">
      <span class="detail-label">Endereco</span>
      <strong>${order.address}</strong>
      <p class="detail-note">${order.notes || "Sem observacoes cadastradas."}</p>
    </div>

    <div class="detail-actions">
      <a class="action-button" href="${buildWhatsAppLink(order)}" target="_blank" rel="noreferrer">Falar no WhatsApp</a>
      <button class="ghost-button" type="button" id="print-order-button">Imprimir OS</button>
      <a class="ghost-button" href="${buildCalendarLink(order)}" target="_blank" rel="noreferrer">Enviar ao Google Calendar</a>
      <button class="ghost-button danger-button" type="button" data-delete-order="${order.id}">Excluir OS</button>
    </div>

    <div class="detail-block">
      <div class="panel-head compact">
        <div>
          <span class="panel-kicker">Anexos</span>
          <h3>Evidencias da ordem</h3>
        </div>
      </div>
      ${
        attachments.length
          ? `
            <div class="attachments-list">
              ${attachments
                .map(
                  (attachment) => `
                    <div class="attachment-item">
                      <div class="attachment-copy">
                        <strong>${attachment.originalName}</strong>
                        <span class="muted">${formatFileSize(attachment.size)} · ${formatDateTimeFromIso(attachment.createdAt)}</span>
                      </div>
                      <div class="table-actions attachment-actions">
                        <a class="ghost-button table-button" href="${attachment.url}" target="_blank" rel="noreferrer">Abrir</a>
                        <button
                          class="ghost-button table-button danger-button"
                          type="button"
                          data-delete-attachment="${attachment.id}"
                          data-attachment-path="${attachment.storagePath}"
                          data-attachment-name="${escapeHtml(attachment.originalName)}"
                        >
                          Excluir
                        </button>
                      </div>
                    </div>
                  `
                )
                .join("")}
            </div>
          `
          : `<div class="empty-state">Nenhum anexo enviado ainda.</div>`
      }
    </div>

    <form class="upload-form" id="attachment-form">
      <label class="field field-full">
        <span>Adicionar foto ou video</span>
        <input type="file" name="file" accept="image/*,video/*,.pdf" required>
      </label>
      <button class="action-button" type="submit">Enviar anexo</button>
    </form>
    <p class="form-error" id="attachment-error" hidden></p>
  `;

  document.getElementById("attachment-form")?.addEventListener("submit", async (event) => {
    event.preventDefault();
    const errorElement = document.getElementById("attachment-error");
    errorElement.hidden = true;

    try {
      await uploadAttachment(selectedOrderId, event.currentTarget);
      await bootstrap();
      showFeedback("Anexo enviado com sucesso.", "success");
    } catch (uploadError) {
      errorElement.textContent = translateError(uploadError.message);
      errorElement.hidden = false;
      showFeedback(translateError(uploadError.message), "error");
    }
  });

  document.getElementById("print-order-button")?.addEventListener("click", () => {
    const printed = printOrderSafe(order);

    if (printed) {
      showFeedback("Ordem de servico enviada para impressao.", "success");
    } else {
      showFeedback("Nao foi possivel abrir a impressao da OS.", "error");
    }
  });
}

function renderFinance() {
  const total = state.orders.reduce((sum, order) => sum + Number(order.amount || 0), 0);
  const paid = state.orders
    .filter((order) => order.paymentStatus === "Pago")
    .reduce((sum, order) => sum + Number(order.amount || 0), 0);
  const pending = state.orders
    .filter((order) => order.paymentStatus !== "Pago")
    .reduce((sum, order) => sum + Number(order.amount || 0), 0);

  elements.financeCards.innerHTML = [
    { label: "Faturamento previsto", hint: "Somando todas as OS", value: formatCurrency(total) },
    { label: "Recebido", hint: "Pagamentos confirmados", value: formatCurrency(paid) },
    { label: "Em aberto", hint: "Cobrancas pendentes", value: formatCurrency(pending) },
  ]
    .map(
      (row) => `
        <div class="finance-row">
          <div>
            <strong>${row.label}</strong>
            <span class="muted">${row.hint}</span>
          </div>
          <strong>${row.value}</strong>
        </div>
      `
    )
    .join("");

  const receivables = state.orders.filter(
    (order) => order.paymentStatus !== "Pago" && matchesSearch([order.customer, order.service, order.address])
  );

  elements.receivablesTable.innerHTML = receivables.length
    ? receivables
        .map(
          (item) => `
            <tr>
              <td>${item.customer}</td>
              <td>${item.service}</td>
              <td>${formatDate(item.date)}</td>
              <td>${formatCurrency(item.amount)}</td>
              <td>
                <select class="status-select" data-payment-id="${item.id}">
                  ${paymentStatuses
                    .map(
                      (status) => `
                        <option value="${status}" ${item.paymentStatus === status ? "selected" : ""}>${status}</option>
                      `
                    )
                    .join("")}
                </select>
              </td>
            </tr>
          `
        )
        .join("")
    : `<tr><td colspan="5"><div class="empty-state">Nenhum recebivel encontrado.</div></td></tr>`;

  document.querySelectorAll("[data-payment-id]").forEach((select) => {
    select.addEventListener("change", async (event) => {
      const { error } = await supabaseClient
        .from("orders")
        .update({ payment_status: event.target.value })
        .eq("id", event.target.dataset.paymentId);

      throwIfError(error);
      await insertActivity("Financeiro atualizado", `Pagamento alterado para ${event.target.value}.`);
      await bootstrap();
      showFeedback("Pagamento atualizado com sucesso.", "success");
    });
  });
}

async function uploadAttachment(orderId, formElement) {
  const form = new FormData(formElement);
  const file = form.get("file");

  if (!(file instanceof File) || !file.name) {
    throw new Error("Selecione um arquivo antes de enviar.");
  }

  const storagePath = `orders/${orderId}/${Date.now()}-${sanitizeFileName(file.name)}`;
  const bucket = state.meta.storageBucket;

  const { error: uploadError } = await supabaseClient.storage.from(bucket).upload(storagePath, file, {
    cacheControl: "3600",
    upsert: false,
  });
  throwIfError(uploadError);

  const { error: insertError } = await supabaseClient.from("attachments").insert({
    order_id: orderId,
    original_name: file.name,
    storage_path: storagePath,
    mime_type: file.type || "application/octet-stream",
    file_size: file.size || 0,
  });
  throwIfError(insertError);

  await insertActivity("Anexo enviado", `${file.name} foi adicionado a ordem selecionada.`);
}

async function createAttachmentViewModel(row) {
  const bucket = state.meta.storageBucket;
  const { data, error } = await supabaseClient.storage.from(bucket).createSignedUrl(row.storage_path, 60 * 60);
  throwIfError(error);

  return {
    id: row.id,
    originalName: row.original_name,
    createdAt: row.created_at,
    size: row.file_size,
    storagePath: row.storage_path,
    url: data.signedUrl,
  };
}

async function insertActivity(title, description) {
  const { error } = await supabaseClient.from("activities").insert({
    title,
    description,
    time_label: "Agora",
  });
  throwIfError(error);
}

function buildCustomerRows() {
  const map = new Map();

  getFilteredLeads().forEach((lead) => {
    if (!map.has(lead.name)) {
      map.set(lead.name, {
        id: lead.id,
        name: lead.name,
        phone: lead.phone,
        service: lead.service,
        address: lead.address,
        status: lead.status,
        lastContact: lead.lastContact,
      });
    }
  });

  return Array.from(map.values());
}

function getFilteredLeads() {
  return state.leads.filter((lead) => matchesSearch([lead.name, lead.phone, lead.service, lead.address, lead.notes]));
}

function getFilteredOrders() {
  return state.orders.filter((order) =>
    matchesSearch([order.code, order.customer, order.service, order.address, order.notes])
  );
}

function getAppointmentsForDate(date) {
  return state.appointments.filter((item) => item.date === date);
}

function getUpcomingAppointments() {
  return [...state.appointments]
    .sort(byDateTime)
    .filter((item) => new Date(`${item.date}T${item.time}:00`).getTime() >= Date.now());
}

function getAgendaDays() {
  const sorted = [...state.appointments]
    .filter((item) => matchesSearch([item.customer, item.service, item.address]))
    .sort(byDateTime);

  const dates = [...new Set(sorted.map((item) => item.date))].slice(0, 3);

  return dates.length
    ? dates.map((date) => ({
        label: new Intl.DateTimeFormat("pt-BR", {
          weekday: "long",
          day: "2-digit",
          month: "2-digit",
        }).format(new Date(`${date}T12:00:00`)),
        items: sorted.filter((item) => item.date === date),
      }))
    : [{ label: "Proximos dias", items: [] }];
}

function matchesSearch(values) {
  if (!searchQuery) {
    return true;
  }

  return values.some((value) => String(value || "").toLowerCase().includes(searchQuery));
}

function byDateTime(a, b) {
  return `${a.date}T${a.time}`.localeCompare(`${b.date}T${b.time}`);
}

function mapLead(row) {
  return {
    id: row.id,
    name: row.name,
    phone: row.phone,
    service: row.service,
    address: row.address,
    source: row.source,
    priority: row.priority,
    status: row.status,
    notes: row.notes,
    createdAt: row.created_at,
    lastContact: row.last_contact,
  };
}

function mapAppointment(row) {
  return {
    id: row.id,
    customer: row.customer,
    phone: row.phone,
    service: row.service,
    address: row.address,
    date: String(row.date),
    time: row.time,
    notes: row.notes,
    createdAt: row.created_at,
  };
}

function mapOrder(row, attachmentsByOrder = {}) {
  return {
    id: row.id,
    code: row.code,
    customer: row.customer,
    phone: row.phone,
    service: row.service,
    address: row.address,
    date: String(row.date),
    time: row.time,
    amount: Number(row.amount || 0),
    status: row.status,
    paymentStatus: row.payment_status,
    notes: row.notes,
    createdAt: row.created_at,
    attachmentCount: attachmentsByOrder[row.id] || 0,
  };
}

function mapActivity(row) {
  return {
    id: row.id,
    title: row.title,
    description: row.description,
    time: row.time_label,
  };
}

function throwIfError(error) {
  if (error) {
    throw new Error(error.message);
  }
}

function translateError(message) {
  if (!message) {
    return "Nao foi possivel concluir a acao.";
  }

  if (message.includes("Invalid login credentials")) {
    return "E-mail ou senha invalidos.";
  }

  return message;
}

function formatCurrency(value) {
  return new Intl.NumberFormat("pt-BR", {
    style: "currency",
    currency: "BRL",
  }).format(Number(value || 0));
}

function formatDate(date) {
  return new Intl.DateTimeFormat("pt-BR", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  }).format(new Date(`${date}T12:00:00`));
}

function formatDateTime(date, time) {
  return `${formatDate(date)} · ${time}`;
}

function formatDateTimeFromIso(value) {
  const date = new Date(value);
  return new Intl.DateTimeFormat("pt-BR", {
    day: "2-digit",
    month: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  }).format(date);
}

function formatFileSize(size) {
  if (size >= 1024 * 1024) {
    return `${(size / (1024 * 1024)).toFixed(1)} MB`;
  }

  return `${Math.max(1, Math.round(size / 1024))} KB`;
}

function todayIso() {
  const now = new Date();
  const adjusted = new Date(now.getTime() - now.getTimezoneOffset() * 60000);
  return adjusted.toISOString().slice(0, 10);
}

function normalizeKey(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, "-");
}

function sanitizeFileName(name) {
  return String(name)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9.\-_]/g, "-");
}

function leadStatusClass(status) {
  return `status-${normalizeKey(getLeadStageLabel(status))}`;
}

function getLeadStageLabel(key) {
  const match = leadStages.find((item) => item.key === key);
  return match ? match.label : key;
}

function getSectionTitle(section) {
  const titles = {
    dashboard: "Dashboard operacional",
    pipeline: "Pipeline de vendas",
    clientes: "Base de clientes",
    agenda: "Agenda de atendimentos",
    ordens: "Ordens de servico",
    financeiro: "Controle financeiro",
  };

  return titles[section] || "Fort Lar CRM";
}

function buildWhatsAppLink(order) {
  const targetPhone = String(order.phone || state.meta.fortlarPhone).replace(/\D/g, "");
  const message = [
    `Ola ${order.customer}, aqui e a equipe Fort Lar.`,
    `Sobre a ordem ${order.code}: ${order.service}.`,
    `Atendimento previsto para ${formatDate(order.date)} as ${order.time}.`,
    `Qualquer duvida, seguimos por aqui.`,
  ].join("\n");

  return `https://wa.me/${targetPhone}?text=${encodeURIComponent(message)}`;
}

function buildCalendarLink(item) {
  const startDate = `${item.date}T${item.time}:00`;
  const endDate = new Date(new Date(startDate).getTime() + 60 * 60 * 1000);
  const details = item.notes
    ? `${item.service}\n\n${item.notes}\n\nEndereco: ${item.address}`
    : `${item.service}\n\nEndereco: ${item.address}`;

  return `https://calendar.google.com/calendar/render?action=TEMPLATE&text=${encodeURIComponent(
    `Fort Lar - ${item.customer}`
  )}&dates=${formatGoogleDate(startDate)}/${formatGoogleDate(endDate.toISOString())}&details=${encodeURIComponent(
    details
  )}&location=${encodeURIComponent(item.address)}`;
}

function formatGoogleDate(value) {
  return new Date(value).toISOString().replace(/[-:]/g, "").replace(/\.\d{3}Z$/, "Z");
}

function confirmAction(message) {
  return window.confirm(message);
}

async function deleteLead(leadId) {
  if (!leadId) {
    return;
  }

  const lead = state.leads.find((item) => item.id === leadId);

  if (!confirmAction(`Excluir o contato ${lead?.name || "selecionado"}? Esta acao nao pode ser desfeita.`)) {
    return;
  }

  try {
    const { error } = await supabaseClient.from("leads").delete().eq("id", leadId);
    throwIfError(error);
    await insertActivity("Contato excluido", `${lead?.name || "Lead"} foi removido do CRM.`);
    await bootstrap();
    showFeedback("Contato excluido com sucesso.", "success");
  } catch (error) {
    showFeedback(translateError(error.message), "error");
  }
}

async function deleteAppointment(appointmentId) {
  if (!appointmentId) {
    return;
  }

  const appointment = state.appointments.find((item) => item.id === appointmentId);

  if (!confirmAction(`Excluir a visita de ${appointment?.customer || "cliente"}?`)) {
    return;
  }

  try {
    const { error } = await supabaseClient.from("appointments").delete().eq("id", appointmentId);
    throwIfError(error);
    await insertActivity("Visita excluida", `${appointment?.customer || "Compromisso"} saiu da agenda.`);
    await bootstrap();
    showFeedback("Visita excluida com sucesso.", "success");
  } catch (error) {
    showFeedback(translateError(error.message), "error");
  }
}

async function deleteOrder(orderId) {
  if (!orderId) {
    return;
  }

  const order = state.orders.find((item) => item.id === orderId);

  if (!confirmAction(`Excluir a ordem ${order?.code || ""} de ${order?.customer || "cliente"}?`)) {
    return;
  }

  try {
    const { data: attachmentsRows, error: attachmentsError } = await supabaseClient
      .from("attachments")
      .select("storage_path")
      .eq("order_id", orderId);
    throwIfError(attachmentsError);

    const storagePaths = (attachmentsRows || []).map((item) => item.storage_path).filter(Boolean);

    if (storagePaths.length) {
      const { error: storageError } = await supabaseClient.storage.from(state.meta.storageBucket).remove(storagePaths);
      throwIfError(storageError);
    }

    const { error } = await supabaseClient.from("orders").delete().eq("id", orderId);
    throwIfError(error);

    if (selectedOrderId === orderId) {
      selectedOrderId = null;
    }

    await insertActivity("OS excluida", `${order?.code || "Ordem"} foi removida do CRM.`);
    await bootstrap();
    showFeedback("Ordem de servico excluida com sucesso.", "success");
  } catch (error) {
    showFeedback(translateError(error.message), "error");
  }
}

async function deleteAttachment(attachmentId, storagePath, attachmentName) {
  if (!attachmentId || !storagePath) {
    return;
  }

  if (!confirmAction(`Excluir o anexo ${attachmentName || "selecionado"}?`)) {
    return;
  }

  try {
    const { error: storageError } = await supabaseClient.storage.from(state.meta.storageBucket).remove([storagePath]);
    throwIfError(storageError);

    const { error } = await supabaseClient.from("attachments").delete().eq("id", attachmentId);
    throwIfError(error);

    await insertActivity("Anexo excluido", `${attachmentName || "Arquivo"} foi removido da OS.`);
    await bootstrap();
    showFeedback("Anexo excluido com sucesso.", "success");
  } catch (error) {
    showFeedback(translateError(error.message), "error");
  }
}

function printOrder(order) {
  const printWindow = window.open("", "_blank", "noopener,noreferrer,width=960,height=720");

  if (!printWindow) {
    return false;
  }

  const printedAt = new Intl.DateTimeFormat("pt-BR", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  }).format(new Date());

  const logoUrl = new URL("./assets/logo-fortlar-transparent.png", window.location.href).href;

  printWindow.document.write(`
    <!DOCTYPE html>
    <html lang="pt-BR">
      <head>
        <meta charset="UTF-8">
        <title>Ordem de Servico ${escapeHtml(order.code)}</title>
        <style>
          :root {
            color-scheme: light;
            --ink: #13304c;
            --soft: #5f748b;
            --line: rgba(19, 48, 76, 0.16);
            --brand: #0f2844;
            --accent: #0d6efd;
            --surface: #f6f9fc;
          }

          * {
            box-sizing: border-box;
          }

          body {
            margin: 0;
            padding: 32px;
            font-family: Arial, Helvetica, sans-serif;
            color: var(--ink);
            background: white;
          }

          .sheet {
            max-width: 860px;
            margin: 0 auto;
            border: 1px solid var(--line);
            border-radius: 18px;
            overflow: hidden;
          }

          .sheet-head {
            padding: 24px 28px;
            background: linear-gradient(135deg, var(--brand), #143f68);
            color: white;
            display: flex;
            justify-content: space-between;
            gap: 24px;
            align-items: center;
          }

          .sheet-head img {
            width: 120px;
            height: auto;
          }

          .sheet-head-copy {
            flex: 1;
          }

          .sheet-head-copy h1 {
            margin: 0 0 8px;
            font-size: 28px;
          }

          .sheet-head-copy p {
            margin: 0;
            color: rgba(255, 255, 255, 0.82);
            line-height: 1.5;
          }

          .sheet-code {
            min-width: 170px;
            padding: 14px 16px;
            border-radius: 16px;
            background: rgba(255, 255, 255, 0.12);
            border: 1px solid rgba(255, 255, 255, 0.14);
            text-align: right;
          }

          .sheet-code strong,
          .sheet-code span {
            display: block;
          }

          .sheet-code span {
            font-size: 12px;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            color: rgba(255, 255, 255, 0.72);
          }

          .sheet-body {
            padding: 28px;
            display: grid;
            gap: 18px;
            background: white;
          }

          .grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 14px;
          }

          .card {
            padding: 16px 18px;
            border-radius: 16px;
            background: var(--surface);
            border: 1px solid var(--line);
          }

          .card strong,
          .card span {
            display: block;
          }

          .label {
            margin-bottom: 6px;
            color: var(--soft);
            font-size: 12px;
            font-weight: 700;
            letter-spacing: 0.08em;
            text-transform: uppercase;
          }

          .value {
            font-size: 16px;
            line-height: 1.45;
            font-weight: 700;
          }

          .notes {
            min-height: 120px;
            white-space: pre-wrap;
          }

          .sheet-footer {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 24px;
            padding-top: 24px;
          }

          .signature {
            padding-top: 42px;
            border-top: 1px solid var(--line);
            text-align: center;
            color: var(--soft);
            font-size: 13px;
          }

          @page {
            size: A4;
            margin: 16mm;
          }

          @media print {
            body {
              padding: 0;
            }

            .sheet {
              border: none;
              border-radius: 0;
            }
          }
        </style>
      </head>
      <body>
        <main class="sheet">
          <header class="sheet-head">
            <div class="sheet-head-copy">
              <img src="${logoUrl}" alt="Fort Lar">
              <h1>Ordem de Servico</h1>
              <p>Documento gerado pelo CRM da Fort Lar para atendimento, aprovacao e acompanhamento do servico.</p>
            </div>
            <div class="sheet-code">
              <span>Ordem</span>
              <strong>${escapeHtml(order.code)}</strong>
              <span>Gerado em ${escapeHtml(printedAt)}</span>
            </div>
          </header>

          <section class="sheet-body">
            <div class="grid">
              <div class="card">
                <span class="label">Cliente</span>
                <strong class="value">${escapeHtml(order.customer)}</strong>
              </div>
              <div class="card">
                <span class="label">Telefone</span>
                <strong class="value">${escapeHtml(order.phone || "Nao informado")}</strong>
              </div>
              <div class="card">
                <span class="label">Servico</span>
                <strong class="value">${escapeHtml(order.service)}</strong>
              </div>
              <div class="card">
                <span class="label">Agendamento</span>
                <strong class="value">${escapeHtml(formatDateTime(order.date, order.time))}</strong>
              </div>
              <div class="card">
                <span class="label">Status da OS</span>
                <strong class="value">${escapeHtml(order.status)}</strong>
              </div>
              <div class="card">
                <span class="label">Pagamento</span>
                <strong class="value">${escapeHtml(order.paymentStatus)} · ${escapeHtml(formatCurrency(order.amount))}</strong>
              </div>
            </div>

            <div class="card">
              <span class="label">Endereco</span>
              <strong class="value">${escapeHtml(order.address)}</strong>
            </div>

            <div class="card notes">
              <span class="label">Descricao / observacoes</span>
              <strong class="value">${escapeHtml(order.notes || "Sem observacoes cadastradas.")}</strong>
            </div>

            <div class="sheet-footer">
              <div class="signature">Assinatura do cliente</div>
              <div class="signature">Assinatura Fort Lar</div>
            </div>
          </section>
        </main>
      </body>
    </html>
  `);

  printWindow.document.close();
  printWindow.focus();

  window.setTimeout(() => {
    printWindow.print();
  }, 300);

  return true;
}

function printOrderSafe(order) {
  const printWindow = window.open("", "_blank", "width=960,height=720");

  if (!printWindow) {
    return false;
  }

  try {
    const printedAt = new Intl.DateTimeFormat("pt-BR", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    }).format(new Date());

    const logoUrl = new URL("./assets/logo-fortlar-transparent.png", window.location.href).href;
    const html = `
      <!DOCTYPE html>
      <html lang="pt-BR">
        <head>
          <meta charset="UTF-8">
          <title>Ordem de Servico ${escapeHtml(order.code)}</title>
          <style>
            :root {
              color-scheme: light;
              --ink: #13304c;
              --soft: #5f748b;
              --line: rgba(19, 48, 76, 0.16);
              --brand: #0f2844;
              --surface: #f6f9fc;
            }

            * {
              box-sizing: border-box;
            }

            body {
              margin: 0;
              padding: 32px;
              font-family: Arial, Helvetica, sans-serif;
              color: var(--ink);
              background: white;
            }

            .sheet {
              max-width: 860px;
              margin: 0 auto;
              border: 1px solid var(--line);
              border-radius: 18px;
              overflow: hidden;
            }

            .sheet-head {
              padding: 24px 28px;
              background: linear-gradient(135deg, var(--brand), #143f68);
              color: white;
              display: flex;
              justify-content: space-between;
              gap: 24px;
              align-items: center;
            }

            .sheet-head img {
              width: 120px;
              height: auto;
              display: block;
              margin-bottom: 12px;
            }

            .sheet-head-copy {
              flex: 1;
            }

            .sheet-head-copy h1 {
              margin: 0 0 8px;
              font-size: 28px;
            }

            .sheet-head-copy p {
              margin: 0;
              color: rgba(255, 255, 255, 0.82);
              line-height: 1.5;
            }

            .sheet-code {
              min-width: 170px;
              padding: 14px 16px;
              border-radius: 16px;
              background: rgba(255, 255, 255, 0.12);
              border: 1px solid rgba(255, 255, 255, 0.14);
              text-align: right;
            }

            .sheet-code strong,
            .sheet-code span {
              display: block;
            }

            .sheet-code span {
              font-size: 12px;
              letter-spacing: 0.08em;
              text-transform: uppercase;
              color: rgba(255, 255, 255, 0.72);
            }

            .sheet-body {
              padding: 28px;
              display: grid;
              gap: 18px;
              background: white;
            }

            .grid {
              display: grid;
              grid-template-columns: repeat(2, minmax(0, 1fr));
              gap: 14px;
            }

            .card {
              padding: 16px 18px;
              border-radius: 16px;
              background: var(--surface);
              border: 1px solid var(--line);
            }

            .label {
              margin-bottom: 6px;
              color: var(--soft);
              font-size: 12px;
              font-weight: 700;
              letter-spacing: 0.08em;
              text-transform: uppercase;
            }

            .value {
              display: block;
              font-size: 16px;
              line-height: 1.45;
              font-weight: 700;
            }

            .notes {
              min-height: 120px;
              white-space: pre-wrap;
            }

            .sheet-footer {
              display: grid;
              grid-template-columns: repeat(2, minmax(0, 1fr));
              gap: 24px;
              padding-top: 24px;
            }

            .signature {
              padding-top: 42px;
              border-top: 1px solid var(--line);
              text-align: center;
              color: var(--soft);
              font-size: 13px;
            }

            @page {
              size: A4;
              margin: 16mm;
            }

            @media print {
              body {
                padding: 0;
              }

              .sheet {
                border: none;
                border-radius: 0;
              }
            }
          </style>
        </head>
        <body>
          <main class="sheet">
            <header class="sheet-head">
              <div class="sheet-head-copy">
                <img src="${logoUrl}" alt="Fort Lar">
                <h1>Ordem de Servico</h1>
                <p>Documento gerado pelo CRM da Fort Lar para atendimento, aprovacao e acompanhamento do servico.</p>
              </div>
              <div class="sheet-code">
                <span>Ordem</span>
                <strong>${escapeHtml(order.code)}</strong>
                <span>Gerado em ${escapeHtml(printedAt)}</span>
              </div>
            </header>

            <section class="sheet-body">
              <div class="grid">
                <div class="card">
                  <span class="label">Cliente</span>
                  <strong class="value">${escapeHtml(order.customer)}</strong>
                </div>
                <div class="card">
                  <span class="label">Telefone</span>
                  <strong class="value">${escapeHtml(order.phone || "Nao informado")}</strong>
                </div>
                <div class="card">
                  <span class="label">Servico</span>
                  <strong class="value">${escapeHtml(order.service)}</strong>
                </div>
                <div class="card">
                  <span class="label">Agendamento</span>
                  <strong class="value">${escapeHtml(formatDateTime(order.date, order.time))}</strong>
                </div>
                <div class="card">
                  <span class="label">Status da OS</span>
                  <strong class="value">${escapeHtml(order.status)}</strong>
                </div>
                <div class="card">
                  <span class="label">Pagamento</span>
                  <strong class="value">${escapeHtml(order.paymentStatus)} · ${escapeHtml(formatCurrency(order.amount))}</strong>
                </div>
              </div>

              <div class="card">
                <span class="label">Endereco</span>
                <strong class="value">${escapeHtml(order.address)}</strong>
              </div>

              <div class="card notes">
                <span class="label">Descricao / observacoes</span>
                <strong class="value">${escapeHtml(order.notes || "Sem observacoes cadastradas.")}</strong>
              </div>

              <div class="sheet-footer">
                <div class="signature">Assinatura do cliente</div>
                <div class="signature">Assinatura Fort Lar</div>
              </div>
            </section>
          </main>
        </body>
      </html>
    `;

    printWindow.document.open();
    printWindow.document.write(html);
    printWindow.document.close();

    let printTriggered = false;

    const triggerPrint = () => {
      if (printTriggered || printWindow.closed) {
        return;
      }

      printTriggered = true;
      printWindow.focus();
      printWindow.print();
    };

    printWindow.onload = () => {
      window.setTimeout(triggerPrint, 250);
    };

    window.setTimeout(triggerPrint, 900);
    return true;
  } catch (_error) {
    try {
      printWindow.close();
    } catch (_closeError) {
    }

    return false;
  }
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
