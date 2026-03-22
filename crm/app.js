const leadStages = [
  { key: "novo", label: "Novo contato" },
  { key: "orcamento", label: "Orcamento enviado" },
  { key: "negociacao", label: "Negociacao" },
  { key: "fechado", label: "Fechado" },
  { key: "arquivado", label: "Arquivado" },
];

const orderStatuses = ["Agendado", "Em andamento", "Concluido", "Cancelado"];
const paymentStatuses = ["Pendente", "Parcial", "Pago"];
const financialEntryStatuses = ["Pendente", "Pago"];

let supabaseClient = null;
let currentUser = null;
let searchQuery = "";
let selectedOrderId = null;

let state = {
  meta: {
    fortlarPhone: "5512996619062",
    storageBucket: "crm-anexos",
    financeModuleReady: true,
  },
  stats: {},
  leads: [],
  appointments: [],
  orders: [],
  activities: [],
  financialEntries: [],
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
  financeAlert: document.getElementById("finance-alert"),
  financeHealth: document.getElementById("finance-health"),
  financeBreakdown: document.getElementById("finance-breakdown"),
  receivablesTable: document.getElementById("receivables-table"),
  financeEntriesTable: document.getElementById("finance-entries-table"),
  financeTimeline: document.getElementById("finance-timeline"),
  financePlanning: document.getElementById("finance-planning"),
  financialEntryButton: document.getElementById("financial-entry-button"),
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
  financialEntryModal: document.getElementById("financial-entry-modal"),
  leadForm: document.getElementById("lead-form"),
  appointmentForm: document.getElementById("appointment-form"),
  orderForm: document.getElementById("order-form"),
  financialEntryForm: document.getElementById("financial-entry-form"),
  leadModalTitle: document.getElementById("lead-modal-title"),
  appointmentModalTitle: document.getElementById("appointment-modal-title"),
  orderModalTitle: document.getElementById("order-modal-title"),
  financialEntryModalTitle: document.getElementById("financial-entry-modal-title"),
  leadSubmitButton: document.getElementById("lead-submit-button"),
  appointmentSubmitButton: document.getElementById("appointment-submit-button"),
  orderSubmitButton: document.getElementById("order-submit-button"),
  financialEntrySubmitButton: document.getElementById("financial-entry-submit-button"),
  logoutButton: document.getElementById("logout-button"),
};

initialize();

async function initialize() {
  elements.currentDate.textContent = new Intl.DateTimeFormat("pt-BR", {
    weekday: "long",
    day: "2-digit",
    month: "long",
  }).format(new Date());

  setDefaultFormDates();

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

function setDefaultFormDates() {
  const today = todayIso();

  elements.appointmentForm?.querySelector('[name="date"]')?.setAttribute("value", today);
  elements.orderForm?.querySelector('[name="date"]')?.setAttribute("value", today);
  elements.financialEntryForm?.querySelector('[name="entryDate"]')?.setAttribute("value", today);
}

function prepareModalForCreate(modalId) {
  if (modalId === "lead-modal") {
    resetLeadFormState();
  }

  if (modalId === "appointment-modal") {
    resetAppointmentFormState();
  }

  if (modalId === "order-modal") {
    resetOrderFormState();
  }

  if (modalId === "financial-entry-modal") {
    resetFinancialEntryFormState();
  }
}

function resetLeadFormState() {
  elements.leadForm?.reset();
  setFormValue(elements.leadForm, "leadId", "");
  if (elements.leadModalTitle) {
    elements.leadModalTitle.textContent = "Cadastrar novo contato";
  }
  if (elements.leadSubmitButton) {
    elements.leadSubmitButton.textContent = "Salvar lead";
  }
}

function resetAppointmentFormState() {
  elements.appointmentForm?.reset();
  setFormValue(elements.appointmentForm, "appointmentId", "");
  setDefaultFormDates();
  if (elements.appointmentModalTitle) {
    elements.appointmentModalTitle.textContent = "Agendar visita";
  }
  if (elements.appointmentSubmitButton) {
    elements.appointmentSubmitButton.textContent = "Salvar visita";
  }
}

function resetOrderFormState() {
  elements.orderForm?.reset();
  setFormValue(elements.orderForm, "orderId", "");
  setDefaultFormDates();
  if (elements.orderModalTitle) {
    elements.orderModalTitle.textContent = "Nova ordem de servico";
  }
  if (elements.orderSubmitButton) {
    elements.orderSubmitButton.textContent = "Salvar OS";
  }
}

function resetFinancialEntryFormState() {
  elements.financialEntryForm?.reset();
  setFormValue(elements.financialEntryForm, "financialEntryId", "");
  setDefaultFormDates();
  if (elements.financialEntryModalTitle) {
    elements.financialEntryModalTitle.textContent = "Novo lancamento contabil";
  }
  if (elements.financialEntrySubmitButton) {
    elements.financialEntrySubmitButton.textContent = "Salvar lancamento";
  }
}

function setFormValue(form, name, value) {
  const field = form?.querySelector(`[name="${name}"]`);

  if (field) {
    field.value = value ?? "";
  }
}

function openLeadEditor(leadId) {
  const lead = state.leads.find((item) => item.id === leadId);

  if (!lead) {
    return;
  }

  setFormValue(elements.leadForm, "leadId", lead.id);
  setFormValue(elements.leadForm, "name", lead.name);
  setFormValue(elements.leadForm, "phone", lead.phone);
  setFormValue(elements.leadForm, "service", lead.service);
  setFormValue(elements.leadForm, "address", lead.address);
  setFormValue(elements.leadForm, "source", lead.source);
  setFormValue(elements.leadForm, "priority", lead.priority);
  setFormValue(elements.leadForm, "notes", lead.notes);

  if (elements.leadModalTitle) {
    elements.leadModalTitle.textContent = "Editar contato";
  }
  if (elements.leadSubmitButton) {
    elements.leadSubmitButton.textContent = "Salvar alteracoes";
  }

  elements.leadModal?.showModal();
}

function openAppointmentEditor(appointmentId) {
  const appointment = state.appointments.find((item) => item.id === appointmentId);

  if (!appointment) {
    return;
  }

  setFormValue(elements.appointmentForm, "appointmentId", appointment.id);
  setFormValue(elements.appointmentForm, "customer", appointment.customer);
  setFormValue(elements.appointmentForm, "phone", appointment.phone);
  setFormValue(elements.appointmentForm, "service", appointment.service);
  setFormValue(elements.appointmentForm, "date", appointment.date);
  setFormValue(elements.appointmentForm, "time", appointment.time);
  setFormValue(elements.appointmentForm, "address", appointment.address);
  setFormValue(elements.appointmentForm, "notes", appointment.notes);

  if (elements.appointmentModalTitle) {
    elements.appointmentModalTitle.textContent = "Editar visita";
  }
  if (elements.appointmentSubmitButton) {
    elements.appointmentSubmitButton.textContent = "Salvar alteracoes";
  }

  elements.appointmentModal?.showModal();
}

function openOrderEditor(orderId) {
  const order = state.orders.find((item) => item.id === orderId);

  if (!order) {
    return;
  }

  setFormValue(elements.orderForm, "orderId", order.id);
  setFormValue(elements.orderForm, "customer", order.customer);
  setFormValue(elements.orderForm, "phone", order.phone);
  setFormValue(elements.orderForm, "service", order.service);
  setFormValue(elements.orderForm, "date", order.date);
  setFormValue(elements.orderForm, "time", order.time);
  setFormValue(elements.orderForm, "amount", order.amount);
  setFormValue(elements.orderForm, "status", order.status);
  setFormValue(elements.orderForm, "paymentStatus", order.paymentStatus);
  setFormValue(elements.orderForm, "address", order.address);
  setFormValue(elements.orderForm, "notes", order.notes);

  if (elements.orderModalTitle) {
    elements.orderModalTitle.textContent = `Editar ${order.code}`;
  }
  if (elements.orderSubmitButton) {
    elements.orderSubmitButton.textContent = "Salvar alteracoes";
  }

  elements.orderModal?.showModal();
}

function openFinancialEntryEditor(entryId) {
  const entry = state.financialEntries.find((item) => item.id === entryId);

  if (!entry) {
    return;
  }

  setFormValue(elements.financialEntryForm, "financialEntryId", entry.id);
  setFormValue(elements.financialEntryForm, "entryType", entry.entryType);
  setFormValue(elements.financialEntryForm, "category", entry.category);
  setFormValue(elements.financialEntryForm, "description", entry.description);
  setFormValue(elements.financialEntryForm, "entryDate", entry.entryDate);
  setFormValue(elements.financialEntryForm, "amount", entry.amount);
  setFormValue(elements.financialEntryForm, "status", entry.status);
  setFormValue(elements.financialEntryForm, "paymentMethod", entry.paymentMethod);
  setFormValue(elements.financialEntryForm, "reference", entry.reference);

  if (elements.financialEntryModalTitle) {
    elements.financialEntryModalTitle.textContent = "Editar lancamento contabil";
  }
  if (elements.financialEntrySubmitButton) {
    elements.financialEntrySubmitButton.textContent = "Salvar alteracoes";
  }

  elements.financialEntryModal?.showModal();
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
  const financialEntries = await loadFinancialEntries();

  state.leads = (leadsResult.data || []).map(mapLead);
  state.appointments = (appointmentsResult.data || []).map(mapAppointment);
  state.orders = (ordersResult.data || []).map((row) => mapOrder(row, attachmentsByOrder));
  state.activities = (activitiesResult.data || []).map(mapActivity);
  state.financialEntries = financialEntries;
  state.stats = buildStats();

  if (!selectedOrderId && state.orders.length) {
    selectedOrderId = state.orders[0].id;
  }

  renderAll();
  await renderSelectedOrder();
}

async function loadFinancialEntries() {
  const { data, error } = await supabaseClient
    .from("financial_entries")
    .select("*")
    .order("entry_date", { ascending: false })
    .order("created_at", { ascending: false });

  if (error) {
    if (isMissingFinancialTableError(error)) {
      state.meta.financeModuleReady = false;
      return [];
    }

    throw new Error(error.message);
  }

  state.meta.financeModuleReady = true;
  return (data || []).map(mapFinancialEntry);
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
  const paidOrders = state.orders
    .filter((order) => order.paymentStatus === "Pago")
    .reduce((sum, order) => sum + Number(order.amount || 0), 0);
  const paidExtraRevenue = state.financialEntries
    .filter((entry) => entry.entryType === "Receita" && entry.status === "Pago")
    .reduce((sum, entry) => sum + Number(entry.amount || 0), 0);

  return {
    openLeads: state.leads.filter((lead) => !["fechado", "arquivado"].includes(lead.status)).length,
    wonLeads: state.leads.filter((lead) => lead.status === "fechado").length,
    activeOrders: state.orders.filter((order) => ["Agendado", "Em andamento"].includes(order.status)).length,
    pendingReceivables:
      state.orders
        .filter((order) => order.paymentStatus !== "Pago")
        .reduce((sum, order) => sum + Number(order.amount || 0), 0) +
      state.financialEntries
        .filter((entry) => entry.entryType === "Receita" && entry.status !== "Pago")
        .reduce((sum, entry) => sum + Number(entry.amount || 0), 0),
    scheduledToday: state.appointments.filter((item) => item.date === todayIso()).length,
    confirmedRevenue: paidOrders + paidExtraRevenue,
  };
}

function bindNavigation() {
  const navButtons = document.querySelectorAll(".nav-item");

  navButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const section = button.dataset.section;
      setActiveSection(section);
      closeSidebar();
    });
  });
}

function bindModals() {
  document.querySelectorAll("[data-open-modal]").forEach((button) => {
    button.addEventListener("click", () => {
      const modalId = button.dataset.openModal;
      prepareModalForCreate(modalId);
      const modal = document.getElementById(modalId);
      modal?.showModal();
    });
  });

  document.querySelectorAll("[data-close-modal]").forEach((button) => {
    button.addEventListener("click", () => {
      button.closest("dialog")?.close();
    });
  });

  elements.leadModal?.addEventListener("close", resetLeadFormState);
  elements.appointmentModal?.addEventListener("close", resetAppointmentFormState);
  elements.orderModal?.addEventListener("close", resetOrderFormState);
  elements.financialEntryModal?.addEventListener("close", resetFinancialEntryFormState);
}

function bindForms() {
  elements.leadForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    clearFeedback();

    try {
      const form = new FormData(elements.leadForm);
      const leadId = String(form.get("leadId") || "").trim();
      const existingLead = state.leads.find((item) => item.id === leadId);
      const payload = {
        name: form.get("name"),
        phone: form.get("phone"),
        service: form.get("service"),
        address: form.get("address"),
        source: form.get("source"),
        priority: form.get("priority"),
        status: existingLead?.status || "novo",
        notes: form.get("notes"),
        last_contact: leadId ? "Atualizado agora" : "Agora",
      };

      const operation = leadId
        ? supabaseClient.from("leads").update(payload).eq("id", leadId)
        : supabaseClient.from("leads").insert(payload);
      const { error } = await operation;

      throwIfError(error);
      await insertActivity(
        leadId ? "Contato atualizado" : "Novo lead cadastrado",
        leadId
          ? `${form.get("name")} teve os dados comerciais atualizados.`
          : `${form.get("name")} entrou para ${form.get("service")}.`
      );
      elements.leadModal.close();
      await bootstrap();
      showFeedback(leadId ? "Contato atualizado com sucesso." : "Lead salvo com sucesso.", "success");
    } catch (error) {
      showFeedback(translateError(error.message), "error");
    }
  });

  elements.appointmentForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    clearFeedback();

    try {
      const form = new FormData(elements.appointmentForm);
      const appointmentId = String(form.get("appointmentId") || "").trim();
      const payload = {
        customer: form.get("customer"),
        phone: form.get("phone"),
        service: form.get("service"),
        address: form.get("address"),
        date: form.get("date"),
        time: form.get("time"),
        notes: form.get("notes"),
      };

      const operation = appointmentId
        ? supabaseClient.from("appointments").update(payload).eq("id", appointmentId)
        : supabaseClient.from("appointments").insert(payload);
      const { error } = await operation;

      throwIfError(error);
      await insertActivity(
        appointmentId ? "Visita atualizada" : "Visita agendada",
        appointmentId
          ? `${form.get("customer")} teve a agenda atualizada para ${form.get("date")} as ${form.get("time")}.`
          : `${form.get("customer")} foi agendado para ${form.get("date")} as ${form.get("time")}.`
      );
      elements.appointmentModal.close();
      await bootstrap();
      showFeedback(appointmentId ? "Visita atualizada com sucesso." : "Visita agendada com sucesso.", "success");
    } catch (error) {
      showFeedback(translateError(error.message), "error");
    }
  });

  elements.orderForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    clearFeedback();

    try {
      const form = new FormData(elements.orderForm);
      const orderId = String(form.get("orderId") || "").trim();
      const payload = {
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
        };

      const query = orderId
        ? supabaseClient.from("orders").update(payload).eq("id", orderId).select("id").single()
        : supabaseClient.from("orders").insert(payload).select("id").single();
      const { data, error } = await query;

      throwIfError(error);
      await insertActivity(
        orderId ? "OS atualizada" : "Nova OS criada",
        orderId
          ? `${form.get("customer")} teve a ordem de servico atualizada.`
          : `${form.get("customer")} entrou em execucao para ${form.get("service")}.`
      );
      selectedOrderId = data.id;
      elements.orderModal.close();
      await bootstrap();
      showFeedback(orderId ? "Ordem de servico atualizada com sucesso." : "Ordem de servico criada com sucesso.", "success");
    } catch (error) {
      showFeedback(translateError(error.message), "error");
    }
  });

  elements.financialEntryForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    clearFeedback();

    if (!state.meta.financeModuleReady) {
      showFeedback("Rode o arquivo crm/supabase-finance-upgrade.sql no Supabase para liberar lancamentos.", "error");
      return;
    }

    try {
      const form = new FormData(elements.financialEntryForm);
      const financialEntryId = String(form.get("financialEntryId") || "").trim();
      const payload = {
        entry_type: form.get("entryType"),
        category: form.get("category"),
        description: form.get("description"),
        entry_date: form.get("entryDate"),
        amount: Number(form.get("amount") || 0),
        status: form.get("status"),
        payment_method: form.get("paymentMethod"),
        reference: form.get("reference"),
      };

      const operation = financialEntryId
        ? supabaseClient.from("financial_entries").update(payload).eq("id", financialEntryId)
        : supabaseClient.from("financial_entries").insert(payload);
      const { error } = await operation;

      throwIfError(error);
      await insertActivity(
        financialEntryId ? "Lancamento financeiro atualizado" : "Lancamento financeiro criado",
        financialEntryId
          ? `${form.get("description")} foi atualizado no financeiro.`
          : `${form.get("entryType")} registrada em ${form.get("category")}.`
      );
      elements.financialEntryModal.close();
      await bootstrap();
      setActiveSection("financeiro");
      showFeedback(financialEntryId ? "Lancamento atualizado com sucesso." : "Lancamento contabil salvo com sucesso.", "success");
    } catch (error) {
      showFeedback(translateError(error.message), "error");
    }
  });
}

function bindSearch() {
  elements.search.addEventListener("input", async (event) => {
    searchQuery = normalizeSearchTerm(event.target.value.trim());
    renderAll();

    const targetSection = findBestSectionForSearch();

    if (targetSection) {
      setActiveSection(targetSection);
    }

    await renderSelectedOrder();
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
    const editLeadButton = event.target.closest("[data-edit-lead]");

    if (editLeadButton) {
      event.preventDefault();
      openLeadEditor(editLeadButton.dataset.editLead);
      return;
    }

    const deleteLeadButton = event.target.closest("[data-delete-lead]");

    if (deleteLeadButton) {
      event.preventDefault();
      await deleteLead(deleteLeadButton.dataset.deleteLead);
      return;
    }

    const editAppointmentButton = event.target.closest("[data-edit-appointment]");

    if (editAppointmentButton) {
      event.preventDefault();
      openAppointmentEditor(editAppointmentButton.dataset.editAppointment);
      return;
    }

    const deleteAppointmentButton = event.target.closest("[data-delete-appointment]");

    if (deleteAppointmentButton) {
      event.preventDefault();
      await deleteAppointment(deleteAppointmentButton.dataset.deleteAppointment);
      return;
    }

    const editOrderButton = event.target.closest("[data-edit-order]");

    if (editOrderButton) {
      event.preventDefault();
      openOrderEditor(editOrderButton.dataset.editOrder);
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
      return;
    }

    const deleteFinancialEntryButton = event.target.closest("[data-delete-financial-entry]");

    const editFinancialEntryButton = event.target.closest("[data-edit-financial-entry]");

    if (editFinancialEntryButton) {
      event.preventDefault();
      openFinancialEntryEditor(editFinancialEntryButton.dataset.editFinancialEntry);
      return;
    }

    if (deleteFinancialEntryButton) {
      event.preventDefault();
      await deleteFinancialEntry(deleteFinancialEntryButton.dataset.deleteFinancialEntry);
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

function setActiveSection(section) {
  document.querySelectorAll(".nav-item").forEach((item) => {
    item.classList.toggle("is-active", item.dataset.section === section);
  });

  document.querySelectorAll(".page-section").forEach((panel) => {
    panel.classList.toggle("is-active", panel.id === section);
  });

  elements.pageTitle.textContent = getSectionTitle(section);
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

  const appointmentSource = getFilteredAppointments();
  const today = appointmentSource.filter((item) => item.date === todayIso());
  const scheduleSource = today.length ? today : getUpcomingAppointments(appointmentSource).slice(0, 4);

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

  const activities = getFilteredActivities();

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
                                class="action-icon-button"
                                type="button"
                                data-edit-lead="${lead.id}"
                                aria-label="Editar lead ${escapeHtml(lead.name)}"
                                title="Editar lead"
                              >
                                <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
                                  <path d="M4 20H8L18.5 9.5C19.33 8.67 19.33 7.33 18.5 6.5V6.5C17.67 5.67 16.33 5.67 15.5 6.5L5 17V20" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
                                </svg>
                              </button>
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
      showFeedback("Funil atualizado com sucesso.", "success");
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
                  <button class="ghost-button table-button" type="button" data-edit-lead="${customer.id}">
                    Editar
                  </button>
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
                            <button class="ghost-button table-button" type="button" data-edit-appointment="${item.id}">
                              Editar
                            </button>
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
                  <button class="ghost-button table-button" type="button" data-edit-order="${order.id}">
                    Editar
                  </button>
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
      <button class="ghost-button" type="button" data-edit-order="${order.id}">Editar OS</button>
      <button class="ghost-button" type="button" id="print-order-button">Imprimir OS</button>
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
  const financeData = buildFinanceData();

  if (elements.financialEntryButton) {
    elements.financialEntryButton.disabled = !state.meta.financeModuleReady;
    elements.financialEntryButton.textContent = state.meta.financeModuleReady ? "Novo lancamento" : "Ativar lancamentos";
  }

  if (state.meta.financeModuleReady) {
    elements.financeAlert.textContent = "";
    elements.financeAlert.classList.add("is-hidden");
  } else {
    elements.financeAlert.innerHTML =
      'Lance despesas e receitas avulsas rodando primeiro o arquivo <strong>crm/supabase-finance-upgrade.sql</strong> no Supabase.';
    elements.financeAlert.classList.remove("is-hidden");
  }

  elements.financeCards.innerHTML = financeData.summaryCards
    .map(
      (card) => `
        <article class="finance-stat-card ${card.tone}">
          <span class="finance-stat-label">${card.label}</span>
          <strong class="finance-stat-value">${card.value}</strong>
          <span class="finance-stat-hint">${card.hint}</span>
          <span class="finance-stat-trend">${card.trend}</span>
        </article>
      `
    )
    .join("");

  elements.financeHealth.innerHTML = financeData.healthItems.length
    ? financeData.healthItems
        .map(
          (item) => `
            <div class="finance-health-row">
              <div>
                <strong>${item.label}</strong>
                <span class="muted">${item.description}</span>
              </div>
              <div class="finance-health-metric">
                <strong>${item.value}</strong>
                <span class="status-pill ${item.tone}">${item.tag}</span>
              </div>
            </div>
          `
        )
        .join("")
    : `<div class="empty-state">Sem indicadores financeiros no momento.</div>`;

  elements.financeBreakdown.innerHTML = `
    <div class="finance-breakdown-group">
      <span class="finance-group-label">Receita por servico</span>
      ${renderBreakdownRows(financeData.serviceBreakdown)}
    </div>
    <div class="finance-breakdown-group">
      <span class="finance-group-label">Despesas por categoria</span>
      ${renderBreakdownRows(financeData.expenseBreakdown)}
    </div>
  `;

  elements.receivablesTable.innerHTML = financeData.receivables.length
    ? financeData.receivables
        .map(
          (item) => `
            <tr>
              <td class="table-cell-primary">
                <strong>${item.code}</strong><br>
                <span class="table-subline">${item.status}</span>
              </td>
              <td>
                <strong>${item.customer}</strong><br>
                <span class="table-subline">${item.phone || "Sem telefone"}</span>
              </td>
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
    : `<tr><td colspan="6"><div class="empty-state">Nenhum recebivel encontrado.</div></td></tr>`;

  elements.financeEntriesTable.innerHTML = state.meta.financeModuleReady
    ? financeData.entries.length
      ? financeData.entries
          .map(
            (entry) => `
              <tr>
                <td><span class="badge ${entry.entryType === "Despesa" ? "finance-badge-danger" : "finance-badge-success"}">${entry.entryType}</span></td>
                <td>${entry.category}</td>
                <td class="table-cell-primary">
                  <strong>${entry.description}</strong><br>
                  <span class="table-subline">${entry.reference || entry.paymentMethod}</span>
                </td>
                <td>${formatDate(entry.entryDate)}</td>
                <td>${formatCurrency(entry.amount)}</td>
                <td>
                  <select class="status-select" data-entry-status-id="${entry.id}">
                    ${financialEntryStatuses
                      .map(
                        (status) => `
                          <option value="${status}" ${entry.status === status ? "selected" : ""}>${status}</option>
                        `
                      )
                      .join("")}
                  </select>
                </td>
                <td>
                  <div class="table-actions">
                    <button class="ghost-button table-button" type="button" data-edit-financial-entry="${entry.id}">
                      Editar
                    </button>
                    <button class="ghost-button table-button danger-button" type="button" data-delete-financial-entry="${entry.id}">
                      Excluir
                    </button>
                  </div>
                </td>
              </tr>
            `
          )
          .join("")
      : `<tr><td colspan="7"><div class="empty-state">Nenhum lancamento manual encontrado.</div></td></tr>`
    : `<tr><td colspan="7"><div class="empty-state">Ative o upgrade financeiro no Supabase para salvar despesas e receitas avulsas.</div></td></tr>`;

  elements.financeTimeline.innerHTML = financeData.timeline.length
    ? financeData.timeline
        .map(
          (item) => `
            <div class="finance-timeline-item">
              <div>
                <strong>${item.title}</strong>
                <span class="muted">${item.description}</span>
              </div>
              <div class="finance-health-metric">
                <strong class="${item.amountTone}">${formatCurrency(item.amount)}</strong>
                <span class="badge">${item.when}</span>
              </div>
            </div>
          `
        )
        .join("")
    : `<div class="empty-state">Nenhuma movimentacao relevante encontrada.</div>`;

  elements.financePlanning.innerHTML = financeData.planning.length
    ? financeData.planning
        .map(
          (item) => `
            <div class="finance-planning-item">
              <div>
                <strong>${item.title}</strong>
                <span class="muted">${item.description}</span>
              </div>
              <span class="status-pill ${item.tone}">${item.tag}</span>
            </div>
          `
        )
        .join("")
    : `<div class="empty-state">Sem recomendacoes financeiras no momento.</div>`;

  document.querySelectorAll("[data-payment-id]").forEach((select) => {
    select.addEventListener("change", async (event) => {
      try {
        const { error } = await supabaseClient
          .from("orders")
          .update({ payment_status: event.target.value })
          .eq("id", event.target.dataset.paymentId);

        throwIfError(error);
        await insertActivity("Financeiro atualizado", `Pagamento de OS alterado para ${event.target.value}.`);
        await bootstrap();
        showFeedback("Pagamento atualizado com sucesso.", "success");
      } catch (error) {
        showFeedback(translateError(error.message), "error");
      }
    });
  });

  document.querySelectorAll("[data-entry-status-id]").forEach((select) => {
    select.addEventListener("change", async (event) => {
      try {
        const { error } = await supabaseClient
          .from("financial_entries")
          .update({ status: event.target.value })
          .eq("id", event.target.dataset.entryStatusId);

        throwIfError(error);
        await insertActivity("Lancamento atualizado", `Situacao financeira alterada para ${event.target.value}.`);
        await bootstrap();
        showFeedback("Lancamento atualizado com sucesso.", "success");
      } catch (error) {
        showFeedback(translateError(error.message), "error");
      }
    });
  });
}

function buildFinanceData() {
  const orders = getFilteredOrders().filter((order) => order.status !== "Cancelado");
  const receivables = getFilteredReceivables();
  const entries = getFilteredFinancialEntries();
  const today = todayIso();
  const currentMonthKey = today.slice(0, 7);

  const orderProjected = sumAmounts(orders, "amount");
  const paidOrders = sumAmounts(
    orders.filter((order) => order.paymentStatus === "Pago"),
    "amount"
  );
  const openOrders = sumAmounts(
    orders.filter((order) => order.paymentStatus !== "Pago"),
    "amount"
  );
  const manualRevenuePaid = sumAmounts(
    entries.filter((entry) => entry.entryType === "Receita" && entry.status === "Pago"),
    "amount"
  );
  const manualRevenuePending = sumAmounts(
    entries.filter((entry) => entry.entryType === "Receita" && entry.status !== "Pago"),
    "amount"
  );
  const expensesPaid = sumAmounts(
    entries.filter((entry) => entry.entryType === "Despesa" && entry.status === "Pago"),
    "amount"
  );
  const expensesPending = sumAmounts(
    entries.filter((entry) => entry.entryType === "Despesa" && entry.status !== "Pago"),
    "amount"
  );
  const confirmedRevenue = paidOrders + manualRevenuePaid;
  const projectedRevenue = orderProjected + manualRevenuePaid + manualRevenuePending;
  const openReceivables = openOrders + manualRevenuePending;
  const operationalBalance = confirmedRevenue - expensesPaid;
  const projectedBalance = projectedRevenue - (expensesPaid + expensesPending);
  const averageTicket = orders.length ? orderProjected / orders.length : 0;
  const overdueReceivables = sumAmounts(
    receivables.filter((item) => item.date < today),
    "amount"
  ) +
    sumAmounts(
      entries.filter((entry) => entry.entryType === "Receita" && entry.status !== "Pago" && entry.entryDate < today),
      "amount"
    );

  const monthOrders = orders.filter((order) => order.date.slice(0, 7) === currentMonthKey);
  const monthEntries = entries.filter((entry) => entry.entryDate.slice(0, 7) === currentMonthKey);
  const monthRevenue =
    sumAmounts(monthOrders.filter((order) => order.paymentStatus === "Pago"), "amount") +
    sumAmounts(monthEntries.filter((entry) => entry.entryType === "Receita" && entry.status === "Pago"), "amount");
  const monthExpenses = sumAmounts(
    monthEntries.filter((entry) => entry.entryType === "Despesa" && entry.status === "Pago"),
    "amount"
  );
  const monthBalance = monthRevenue - monthExpenses;
  const margin = confirmedRevenue > 0 ? operationalBalance / confirmedRevenue : 0;

  const serviceBreakdown = buildBreakdownItems(orders, "service", "amount");
  const expenseBreakdown = buildBreakdownItems(
    entries.filter((entry) => entry.entryType === "Despesa"),
    "category",
    "amount"
  );

  const timeline = buildFinanceTimeline(receivables, entries);
  const planning = buildFinancePlanning({
    openReceivables,
    overdueReceivables,
    expensesPending,
    operationalBalance,
    monthBalance,
    topService: serviceBreakdown[0],
    financeModuleReady: state.meta.financeModuleReady,
  });

  return {
    receivables,
    entries,
    serviceBreakdown,
    expenseBreakdown,
    timeline,
    planning,
    summaryCards: [
      {
        label: "Saldo operacional",
        value: formatCurrency(operationalBalance),
        hint: "Receita confirmada menos despesas pagas",
        trend: operationalBalance >= 0 ? "Operacao saudavel" : "Caixa apertado",
        tone: operationalBalance >= 0 ? "is-positive" : "is-warning",
      },
      {
        label: "Receita confirmada",
        value: formatCurrency(confirmedRevenue),
        hint: "OS pagas + receitas extras liquidadas",
        trend: `${orders.filter((order) => order.paymentStatus === "Pago").length} OS recebidas`,
        tone: "is-positive",
      },
      {
        label: "A receber",
        value: formatCurrency(openReceivables),
        hint: "OS pendentes e receitas futuras",
        trend: `${receivables.length} cobrancas abertas`,
        tone: openReceivables > 0 ? "is-warning" : "is-neutral",
      },
      {
        label: "Despesas pagas",
        value: formatCurrency(expensesPaid),
        hint: "Saidas confirmadas no financeiro",
        trend: `${entries.filter((entry) => entry.entryType === "Despesa" && entry.status === "Pago").length} lancamentos pagos`,
        tone: "is-danger",
      },
      {
        label: "Despesas previstas",
        value: formatCurrency(expensesPending),
        hint: "Compromissos ainda em aberto",
        trend: expensesPending > 0 ? "Reserve caixa" : "Sem pressao imediata",
        tone: expensesPending > 0 ? "is-warning" : "is-neutral",
      },
      {
        label: "Ticket medio",
        value: formatCurrency(averageTicket),
        hint: "Media financeira por OS ativa",
        trend: `${orders.length} ordens no radar`,
        tone: "is-neutral",
      },
    ],
    healthItems: [
      {
        label: "Resultado do mes",
        description: "Entradas e saidas pagas no mes atual.",
        value: formatCurrency(monthBalance),
        tag: monthBalance >= 0 ? "Positivo" : "Atencao",
        tone: monthBalance >= 0 ? "status-pago" : "status-cancelado",
      },
      {
        label: "Previsao liquida",
        description: "Receita prevista menos despesas ja projetadas.",
        value: formatCurrency(projectedBalance),
        tag: projectedBalance >= 0 ? "Planejado" : "Risco",
        tone: projectedBalance >= 0 ? "status-concluido" : "status-cancelado",
      },
      {
        label: "Inadimplencia em atraso",
        description: "Valores que ja passaram da data prevista.",
        value: formatCurrency(overdueReceivables),
        tag: overdueReceivables > 0 ? "Cobrar" : "Em dia",
        tone: overdueReceivables > 0 ? "status-cancelado" : "status-pago",
      },
      {
        label: "Margem operacional",
        description: "Saldo operacional sobre a receita confirmada.",
        value: formatPercent(margin),
        tag: margin >= 0.2 ? "Saudavel" : "Monitorar",
        tone: margin >= 0.2 ? "status-pago" : "status-orcamento-enviado",
      },
    ],
  };
}

function renderBreakdownRows(items) {
  if (!items.length) {
    return `<div class="empty-state">Ainda nao ha dados suficientes para montar este comparativo.</div>`;
  }

  return items
    .map(
      (item) => `
        <div class="finance-breakdown-row">
          <div class="finance-breakdown-copy">
            <strong>${item.label}</strong>
            <span class="muted">${item.count} registro(s)</span>
          </div>
          <div class="finance-breakdown-metric">
            <strong>${formatCurrency(item.total)}</strong>
            <div class="finance-progress">
              <span class="finance-progress-bar" style="width:${Math.max(item.share, 8)}%"></span>
            </div>
          </div>
        </div>
      `
    )
    .join("");
}

function buildBreakdownItems(items, labelKey, amountKey) {
  const grouped = items.reduce((accumulator, item) => {
    const label = item[labelKey] || "Sem classificacao";

    if (!accumulator[label]) {
      accumulator[label] = { label, total: 0, count: 0 };
    }

    accumulator[label].total += Number(item[amountKey] || 0);
    accumulator[label].count += 1;
    return accumulator;
  }, {});

  const totals = Object.values(grouped).sort((left, right) => right.total - left.total);
  const grandTotal = totals.reduce((sum, item) => sum + item.total, 0) || 1;

  return totals.slice(0, 5).map((item) => ({
    ...item,
    share: Math.round((item.total / grandTotal) * 100),
  }));
}

function buildFinanceTimeline(receivables, entries) {
  const items = [
    ...receivables.map((item) => ({
      sortDate: item.date,
      title: `Recebimento ${item.code}`,
      description: `${item.customer} - ${item.service}`,
      amount: Number(item.amount || 0),
      amountTone: "finance-positive",
      when: formatDate(item.date),
    })),
    ...entries
      .filter((entry) => entry.status !== "Pago")
      .map((entry) => ({
        sortDate: entry.entryDate,
        title: `${entry.entryType} - ${entry.category}`,
        description: entry.description,
        amount: Number(entry.amount || 0),
        amountTone: entry.entryType === "Despesa" ? "finance-negative" : "finance-positive",
        when: formatDate(entry.entryDate),
      })),
  ];

  return items.sort((left, right) => left.sortDate.localeCompare(right.sortDate)).slice(0, 6);
}

function buildFinancePlanning(data) {
  const items = [];

  if (!data.financeModuleReady) {
    items.push({
      title: "Liberar lancamentos avulsos",
      description: "Rode o upgrade financeiro no Supabase para registrar despesas, impostos e receitas extras.",
      tag: "Configurar",
      tone: "status-orcamento-enviado",
    });
  }

  if (data.overdueReceivables > 0) {
    items.push({
      title: "Cobrar valores vencidos",
      description: `Existem ${formatCurrency(data.overdueReceivables)} em atraso e isso merece prioridade hoje.`,
      tag: "Urgente",
      tone: "status-cancelado",
    });
  }

  if (data.expensesPending > 0) {
    items.push({
      title: "Reservar caixa para despesas",
      description: `Separe ${formatCurrency(data.expensesPending)} para compromissos pendentes do operacional.`,
      tag: "Planejar",
      tone: "status-orcamento-enviado",
    });
  }

  if (data.topService) {
    items.push({
      title: "Servico lider de faturamento",
      description: `${data.topService.label} concentra ${formatCurrency(data.topService.total)} no periodo filtrado.`,
      tag: "Escalar",
      tone: "status-pago",
    });
  }

  items.push({
    title: "Resultado do mes",
    description:
      data.monthBalance >= 0
        ? `A operacao do mes esta positiva em ${formatCurrency(data.monthBalance)}.`
        : `O mes esta negativo em ${formatCurrency(Math.abs(data.monthBalance))}; vale rever custos e cobrancas.`,
    tag: data.monthBalance >= 0 ? "Saudavel" : "Atencao",
    tone: data.monthBalance >= 0 ? "status-pago" : "status-cancelado",
  });

  if (data.openReceivables <= 0) {
    items.push({
      title: "Carteira em dia",
      description: "Nao ha valores em aberto relevantes no momento.",
      tag: "Estavel",
      tone: "status-pago",
    });
  }

  return items.slice(0, 5);
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
  return state.leads.filter((lead) =>
    matchesSearch([lead.name, lead.phone, lead.service, lead.address, lead.notes, lead.source, lead.priority, lead.status])
  );
}

function getFilteredAppointments() {
  return state.appointments.filter((item) =>
    matchesSearch([item.customer, item.phone, item.service, item.address, item.notes, item.date, item.time, formatDate(item.date)])
  );
}

function getFilteredOrders() {
  return state.orders.filter((order) =>
    matchesSearch([
      order.code,
      order.customer,
      order.phone,
      order.service,
      order.address,
      order.notes,
      order.status,
      order.paymentStatus,
      formatDate(order.date),
      formatDateTime(order.date, order.time),
    ])
  );
}

function getAppointmentsForDate(date) {
  return state.appointments.filter((item) => item.date === date);
}

function getUpcomingAppointments(source = state.appointments) {
  return [...source]
    .sort(byDateTime)
    .filter((item) => new Date(`${item.date}T${item.time}:00`).getTime() >= Date.now());
}

function getAgendaDays() {
  const sorted = [...getFilteredAppointments()].sort(byDateTime);

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

function getFilteredActivities() {
  return state.activities.filter((activity) => matchesSearch([activity.title, activity.description, activity.time]));
}

function getFilteredReceivables() {
  return state.orders.filter(
    (order) =>
      order.paymentStatus !== "Pago" &&
      matchesSearch([
        order.code,
        order.customer,
        order.phone,
        order.service,
        order.address,
        order.paymentStatus,
        formatDate(order.date),
      ])
  );
}

function getFilteredFinancialEntries() {
  return state.financialEntries.filter((entry) =>
    matchesSearch([
      entry.entryType,
      entry.category,
      entry.description,
      entry.status,
      entry.paymentMethod,
      entry.reference,
      entry.entryDate,
      formatDate(entry.entryDate),
      formatCurrency(entry.amount),
    ])
  );
}

function matchesSearch(values) {
  if (!searchQuery) {
    return true;
  }

  const compactQuery = searchQuery.replace(/\s+/g, "");

  return values.some((value) => {
    const normalizedValue = normalizeSearchTerm(value);
    const compactValue = normalizedValue.replace(/\s+/g, "");

    return normalizedValue.includes(searchQuery) || compactValue.includes(compactQuery);
  });
}

function byDateTime(a, b) {
  return `${a.date}T${a.time}`.localeCompare(`${b.date}T${b.time}`);
}

function normalizeSearchTerm(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function getCurrentSectionId() {
  return document.querySelector(".page-section.is-active")?.id || "dashboard";
}

function getSearchSectionCounts() {
  return {
    dashboard: getFilteredActivities().length + getFilteredLeads().length + getFilteredAppointments().length,
    pipeline: getFilteredLeads().length,
    clientes: buildCustomerRows().length,
    agenda: getFilteredAppointments().length,
    ordens: getFilteredOrders().length,
    financeiro: getFilteredReceivables().length + getFilteredFinancialEntries().length,
  };
}

function findBestSectionForSearch() {
  if (!searchQuery) {
    return null;
  }

  const counts = getSearchSectionCounts();
  const currentSection = getCurrentSectionId();

  if (currentSection !== "dashboard" && (counts[currentSection] || 0) > 0) {
    return currentSection;
  }

  return ["pipeline", "clientes", "agenda", "ordens", "financeiro", "dashboard"].find(
    (section) => (counts[section] || 0) > 0
  ) || currentSection;
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

function mapFinancialEntry(row) {
  return {
    id: row.id,
    entryType: row.entry_type,
    category: row.category,
    description: row.description,
    entryDate: String(row.entry_date),
    amount: Number(row.amount || 0),
    status: row.status,
    paymentMethod: row.payment_method,
    reference: row.reference || "",
    createdAt: row.created_at,
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

  if (message.includes("financial_entries")) {
    return "O modulo financeiro precisa do upgrade SQL no Supabase para salvar lancamentos.";
  }

  return message;
}

function formatCurrency(value) {
  return new Intl.NumberFormat("pt-BR", {
    style: "currency",
    currency: "BRL",
  }).format(Number(value || 0));
}

function formatPercent(value) {
  return new Intl.NumberFormat("pt-BR", {
    style: "percent",
    minimumFractionDigits: 0,
    maximumFractionDigits: 1,
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

function sumAmounts(items, fieldName) {
  return items.reduce((sum, item) => sum + Number(item[fieldName] || 0), 0);
}

function isMissingFinancialTableError(error) {
  const message = String(error?.message || "");
  const details = String(error?.details || "");
  return (
    error?.code === "42P01" ||
    error?.code === "PGRST205" ||
    message.includes("financial_entries") ||
    message.includes("Could not find the table") ||
    details.includes("financial_entries")
  );
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
    pipeline: "Funil de vendas",
    clientes: "Base de clientes",
    agenda: "Agenda de atendimentos",
    ordens: "Ordens de servico",
    financeiro: "Financeiro e contabilidade",
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

async function deleteFinancialEntry(entryId) {
  if (!entryId || !state.meta.financeModuleReady) {
    return;
  }

  const entry = state.financialEntries.find((item) => item.id === entryId);

  if (!confirmAction(`Excluir o lancamento ${entry?.description || "selecionado"}?`)) {
    return;
  }

  try {
    const { error } = await supabaseClient.from("financial_entries").delete().eq("id", entryId);
    throwIfError(error);
    await insertActivity("Lancamento excluido", `${entry?.description || "Registro financeiro"} foi removido do CRM.`);
    await bootstrap();
    setActiveSection("financeiro");
    showFeedback("Lancamento excluido com sucesso.", "success");
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
