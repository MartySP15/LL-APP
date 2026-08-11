import React, { useState, useEffect, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  LayoutDashboard,
  Users,
  FileText,
  Plus,
  Trash2,
  Pencil,
  X,
  Printer,
  Search,
  ArrowLeft,
  Receipt,
  ChevronRight,
  AlertCircle,
  Settings,
  Home,
  LogOut,
  Download,
  Upload,
  FileSpreadsheet,
} from "lucide-react";
import { supabase } from "./supabaseClient";

/* ---------------------------------------------------------------
   Helpers
--------------------------------------------------------------- */

const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36);

const fmtMoney = (n) =>
  (Number.isFinite(n) ? n : 0).toLocaleString("en-IE", {
    style: "currency",
    currency: "EUR",
  });

const fmtDate = (d) => {
  if (!d) return "—";
  const dt = new Date(d + "T00:00:00");
  if (isNaN(dt.getTime())) return "—";
  return dt.toLocaleDateString("en-US", { year: "numeric", month: "short", day: "numeric" });
};

const todayISO = () => new Date().toISOString().slice(0, 10);

function computeTotals(order) {
  const subtotal = (order.items || []).reduce(
    (s, it) => s + (Number(it.qty) || 0) * (Number(it.price) || 0),
    0
  );
  const taxAmt = subtotal * ((Number(order.taxRate) || 0) / 100);
  const total = subtotal + taxAmt;
  const paid = Number(order.amountPaid) || 0;
  const balance = Math.max(0, total - paid);
  let paymentStatus = "unpaid";
  if (order.orderStatus !== "cancelled") {
    if (paid <= 0) paymentStatus = "unpaid";
    else if (paid >= total && total > 0) paymentStatus = "paid";
    else paymentStatus = "partial";
  }
  return { subtotal, taxAmt, total, paid, balance, paymentStatus };
}

const STATUS_LABEL = {
  draft: "Draft",
  sent: "Sent",
  fulfilled: "Fulfilled",
  cancelled: "Cancelled",
};

const NAME_KEYS = ["name", "customer", "customer name", "full name", "client", "client name"];
const EMAIL_KEYS = ["email", "email address", "e-mail"];
const PHONE_KEYS = ["phone", "phone number", "mobile", "tel", "telephone"];
const ADDRESS_KEYS = ["address", "billing address", "address 1"];

function pickField(row, keys) {
  const lowerMap = {};
  Object.keys(row).forEach((k) => {
    lowerMap[k.trim().toLowerCase()] = row[k];
  });
  for (const k of keys) {
    if (lowerMap[k] !== undefined && String(lowerMap[k]).trim() !== "") {
      return String(lowerMap[k]).trim();
    }
  }
  return "";
}

function normalizeImportRow(row) {
  return {
    name: pickField(row, NAME_KEYS),
    email: pickField(row, EMAIL_KEYS),
    phone: pickField(row, PHONE_KEYS),
    address: pickField(row, ADDRESS_KEYS),
  };
}

function exportCustomersToExcel(customers, orders) {
  const rows = customers.map((c) => ({
    Name: c.name,
    Email: c.email || "",
    Phone: c.phone || "",
    Address: c.address || "",
    Orders: orders.filter((o) => o.customerId === c.id).length,
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = [{ wch: 24 }, { wch: 26 }, { wch: 16 }, { wch: 32 }, { wch: 8 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Customers");
  XLSX.writeFile(wb, `customers-${todayISO()}.xlsx`);
}

function exportOrdersToExcel(orders, customers) {
  const rows = orders.map((o) => {
    const t = computeTotals(o);
    const customer = customers.find((c) => c.id === o.customerId);
    const itemsText = (o.items || [])
      .map((it) => `${it.description} (x${it.qty} @ ${fmtMoney(it.price)})`)
      .join("; ");
    return {
      "Order #": o.orderNumber,
      "Invoice #": o.invoiceNumber || "",
      Customer: customer?.name || "",
      Date: o.date || "",
      "Due Date": o.dueDate || "",
      Status: STATUS_LABEL[o.orderStatus] || o.orderStatus,
      "Payment Status": t.paymentStatus,
      Items: itemsText,
      Subtotal: Number(t.subtotal.toFixed(2)),
      "Tax %": o.taxRate || 0,
      Tax: Number(t.taxAmt.toFixed(2)),
      Total: Number(t.total.toFixed(2)),
      Paid: Number(t.paid.toFixed(2)),
      Balance: Number(t.balance.toFixed(2)),
      Notes: o.notes || "",
    };
  });
  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = [
    { wch: 12 }, { wch: 12 }, { wch: 22 }, { wch: 12 }, { wch: 12 },
    { wch: 12 }, { wch: 14 }, { wch: 40 }, { wch: 11 }, { wch: 8 },
    { wch: 10 }, { wch: 11 }, { wch: 11 }, { wch: 11 }, { wch: 30 },
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Orders");
  XLSX.writeFile(wb, `orders-${todayISO()}.xlsx`);
}

const DEFAULT_BUSINESS = {
  name: "Liveable Layouts",
  tagline: "Designed to work in real life",
  address: "",
  email: "",
  paymentTerms: "Due on receipt",
  paymentMethod: "",
  vatNote: "",
  signatureName: "",
  socialLinks: "",
  accent: "#2C6CB0",
};

/* ---------------------------------------------------------------
   Supabase <-> app field mapping
--------------------------------------------------------------- */

function orderFromDb(row) {
  return {
    id: row.id,
    orderNumber: row.order_number,
    invoiceNumber: row.invoice_number,
    customerId: row.customer_id,
    date: row.date,
    dueDate: row.due_date,
    taxRate: row.tax_rate,
    notes: row.notes,
    items: row.items || [],
    orderStatus: row.order_status,
    amountPaid: row.amount_paid,
  };
}

function orderToDb(order) {
  return {
    invoice_number: order.invoiceNumber ?? null,
    customer_id: order.customerId,
    date: order.date,
    due_date: order.dueDate || null,
    tax_rate: order.taxRate ?? 0,
    notes: order.notes ?? "",
    items: order.items ?? [],
    order_status: order.orderStatus ?? "draft",
    amount_paid: order.amountPaid ?? 0,
  };
}

function businessFromDb(row) {
  if (!row) return DEFAULT_BUSINESS;
  return {
    name: row.name || DEFAULT_BUSINESS.name,
    tagline: row.tagline || "",
    address: row.address || "",
    email: row.email || "",
    paymentTerms: row.payment_terms || "",
    paymentMethod: row.payment_method || "",
    vatNote: row.vat_note || "",
    signatureName: row.signature_name || "",
    socialLinks: row.social_links || "",
    accent: row.accent || DEFAULT_BUSINESS.accent,
  };
}

function businessToDb(b) {
  return {
    id: 1,
    name: b.name,
    tagline: b.tagline,
    address: b.address,
    email: b.email,
    payment_terms: b.paymentTerms,
    payment_method: b.paymentMethod,
    vat_note: b.vatNote,
    signature_name: b.signatureName,
    social_links: b.socialLinks,
    accent: b.accent,
  };
}

/* ---------------------------------------------------------------
   Root App
--------------------------------------------------------------- */

export default function App() {
  const [session, setSession] = useState(undefined); // undefined = checking, null = logged out, object = logged in
  const [customers, setCustomers] = useState([]);
  const [orders, setOrders] = useState([]);
  const [business, setBusiness] = useState(DEFAULT_BUSINESS);
  const [loaded, setLoaded] = useState(false);
  const [error, setError] = useState(null);

  const [view, setView] = useState("dashboard");
  const [activeOrderId, setActiveOrderId] = useState(null);
  const [invoiceOrderId, setInvoiceOrderId] = useState(null);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: listener } = supabase.auth.onAuthStateChange((_event, sess) => {
      setSession(sess);
    });
    return () => listener.subscription.unsubscribe();
  }, []);

  const fetchAll = useCallback(async () => {
    try {
      const [custRes, orderRes, bizRes] = await Promise.all([
        supabase.from("customers").select("*").order("name"),
        supabase.from("orders").select("*").order("date", { ascending: false }),
        supabase.from("business_settings").select("*").eq("id", 1).maybeSingle(),
      ]);
      if (custRes.error) throw custRes.error;
      if (orderRes.error) throw orderRes.error;
      if (bizRes.error) throw bizRes.error;
      setCustomers(custRes.data || []);
      setOrders((orderRes.data || []).map(orderFromDb));
      setBusiness(businessFromDb(bizRes.data));
      setError(null);
    } catch (e) {
      setError(e.message || "Couldn't reach the database.");
    } finally {
      setLoaded(true);
    }
  }, []);

  useEffect(() => {
    if (!session) return;
    fetchAll();
    // Live sync: refetch when any team member changes data.
    const channel = supabase
      .channel("orders-app-sync")
      .on("postgres_changes", { event: "*", schema: "public", table: "customers" }, fetchAll)
      .on("postgres_changes", { event: "*", schema: "public", table: "orders" }, fetchAll)
      .on("postgres_changes", { event: "*", schema: "public", table: "business_settings" }, fetchAll)
      .subscribe();
    return () => {
      supabase.removeChannel(channel);
    };
  }, [session, fetchAll]);

  /* ---- Customers ---- */

  const addCustomer = async (customer) => {
    const { error: err } = await supabase.from("customers").insert(customer);
    if (err) setError(err.message);
    else fetchAll();
  };

  const updateCustomer = async (id, patch) => {
    const { error: err } = await supabase.from("customers").update(patch).eq("id", id);
    if (err) setError(err.message);
    else fetchAll();
  };

  const importCustomers = async (rawRows) => {
    const rows = rawRows.map(normalizeImportRow).filter((r) => r.name);
    const toInsert = [];
    const toUpdate = [];
    let skipped = rawRows.length - rows.length;

    rows.forEach((row) => {
      const existing = customers.find((c) => {
        if (row.email && c.email) return c.email.trim().toLowerCase() === row.email.toLowerCase();
        return c.name.trim().toLowerCase() === row.name.toLowerCase();
      });
      if (existing) {
        const patch = {};
        if (row.email && row.email !== existing.email) patch.email = row.email;
        if (row.phone && row.phone !== existing.phone) patch.phone = row.phone;
        if (row.address && row.address !== existing.address) patch.address = row.address;
        if (Object.keys(patch).length > 0) toUpdate.push({ id: existing.id, patch });
      } else {
        toInsert.push({ name: row.name, email: row.email, phone: row.phone, address: row.address });
      }
    });

    try {
      if (toInsert.length > 0) {
        const { error: insErr } = await supabase.from("customers").insert(toInsert);
        if (insErr) throw insErr;
      }
      for (const u of toUpdate) {
        const { error: updErr } = await supabase.from("customers").update(u.patch).eq("id", u.id);
        if (updErr) throw updErr;
      }
      await fetchAll();
      setError(null);
      return { added: toInsert.length, updated: toUpdate.length, skipped };
    } catch (e) {
      setError(e.message || "Import failed.");
      return { added: 0, updated: 0, skipped: rawRows.length, failed: true };
    }
  };

  const deleteCustomer = async (id) => {
    const { error: err } = await supabase.from("customers").delete().eq("id", id);
    if (err) setError(err.message);
    else fetchAll();
  };

  /* ---- Orders ---- */

  const addOrder = async (order) => {
    const { error: err } = await supabase.from("orders").insert(orderToDb(order));
    if (err) setError(err.message);
    else fetchAll();
  };

  const updateOrder = async (id, patch) => {
    const dbPatch = {};
    if ("customerId" in patch) dbPatch.customer_id = patch.customerId;
    if ("date" in patch) dbPatch.date = patch.date;
    if ("dueDate" in patch) dbPatch.due_date = patch.dueDate || null;
    if ("taxRate" in patch) dbPatch.tax_rate = patch.taxRate;
    if ("notes" in patch) dbPatch.notes = patch.notes;
    if ("items" in patch) dbPatch.items = patch.items;
    if ("orderStatus" in patch) dbPatch.order_status = patch.orderStatus;
    if ("amountPaid" in patch) dbPatch.amount_paid = patch.amountPaid;
    if ("invoiceNumber" in patch) dbPatch.invoice_number = patch.invoiceNumber;
    const { error: err } = await supabase.from("orders").update(dbPatch).eq("id", id);
    if (err) setError(err.message);
    else fetchAll();
  };

  const deleteOrder = async (id) => {
    const { error: err } = await supabase.from("orders").delete().eq("id", id);
    if (err) setError(err.message);
    else fetchAll();
  };

  const generateInvoice = async (order) => {
    if (order.invoiceNumber) {
      setInvoiceOrderId(order.id);
      return;
    }
    const { data, error: rpcErr } = await supabase.rpc("next_invoice_number");
    if (rpcErr) {
      setError(rpcErr.message);
      return;
    }
    const patch = { invoice_number: data };
    if (order.orderStatus === "draft") patch.order_status = "sent";
    const { error: err } = await supabase.from("orders").update(patch).eq("id", order.id);
    if (err) setError(err.message);
    else await fetchAll();
    setInvoiceOrderId(order.id);
  };

  /* ---- Business settings ---- */

  const updateBusiness = async (patch) => {
    const merged = { ...business, ...patch };
    const { error: err } = await supabase.from("business_settings").upsert(businessToDb(merged));
    if (err) setError(err.message);
    else fetchAll();
  };

  if (session === undefined) {
    return (
      <Shell>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100%", color: "var(--ink-soft)" }}>
          Checking session…
        </div>
      </Shell>
    );
  }

  if (!session) {
    return (
      <Shell>
        <LoginScreen />
      </Shell>
    );
  }

  if (!loaded) {
    return (
      <Shell>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100%", color: "var(--ink-soft)" }}>
          Loading ledger…
        </div>
      </Shell>
    );
  }

  const activeOrder = orders.find((o) => o.id === activeOrderId) || null;
  const invoiceOrder = orders.find((o) => o.id === invoiceOrderId) || null;

  return (
    <Shell>
      <div className="app-grid">
        <Sidebar
          view={view}
          setView={(v) => { setView(v); setActiveOrderId(null); }}
          userEmail={session.user?.email}
          onSignOut={() => supabase.auth.signOut()}
        />
        <main className="main">
          {error && (
            <div className="save-banner">
              <AlertCircle size={14} /> {error}
            </div>
          )}
          {view === "dashboard" && (
            <Dashboard
              orders={orders}
              customers={customers}
              onOpenOrder={(id) => {
                setView("orders");
                setActiveOrderId(id);
              }}
            />
          )}
          {view === "customers" && (
            <CustomersView
              customers={customers}
              orders={orders}
              onAdd={addCustomer}
              onUpdate={updateCustomer}
              onDelete={deleteCustomer}
              onImport={importCustomers}
            />
          )}
          {view === "orders" && !activeOrder && (
            <OrdersListView
              orders={orders}
              customers={customers}
              onAdd={addOrder}
              onOpen={(id) => setActiveOrderId(id)}
              onGoToCustomers={() => setView("customers")}
            />
          )}
          {view === "orders" && activeOrder && (
            <OrderDetailView
              order={activeOrder}
              customers={customers}
              onBack={() => setActiveOrderId(null)}
              onUpdate={(patch) => updateOrder(activeOrder.id, patch)}
              onDelete={() => {
                deleteOrder(activeOrder.id);
                setActiveOrderId(null);
              }}
              onGenerateInvoice={() => generateInvoice(activeOrder)}
            />
          )}
          {view === "settings" && (
            <SettingsView business={business} onSave={updateBusiness} />
          )}
        </main>
      </div>
      {invoiceOrder && (
        <InvoiceModal
          order={invoiceOrder}
          customer={customers.find((c) => c.id === invoiceOrder.customerId)}
          business={business}
          onClose={() => setInvoiceOrderId(null)}
        />
      )}
    </Shell>
  );
}
function Shell({ children }) {
  return (
    <div className="ledger-root">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600;700&family=IBM+Plex+Sans:wght@400;500;600;700&display=swap');

        .ledger-root {
          --paper: #EFF1EA;
          --paper-raised: #F8F9F4;
          --paper-line: #D7DBCF;
          --ink: #1F2A33;
          --ink-soft: #5B6A70;
          --ink-faint: #93A0A0;
          --brass: #A9793D;
          --brass-dark: #86602F;
          --sage: #55805F;
          --sage-bg: #E4EBE1;
          --amber: #C0912E;
          --amber-bg: #F3E9D3;
          --rust: #AE5138;
          --rust-bg: #F1E0D9;
          --line: #CBD1C2;
          font-family: 'IBM Plex Sans', sans-serif;
          color: var(--ink);
          background: var(--paper);
          min-height: 100vh;
          width: 100%;
        }
        .ledger-root * { box-sizing: border-box; }
        .mono { font-family: 'IBM Plex Mono', monospace; font-variant-numeric: tabular-nums; }

        .app-grid { display: flex; min-height: 100vh; }

        .sidebar {
          width: 220px;
          flex-shrink: 0;
          background: #387ABC;
          color: #ffffff;
          display: flex;
          flex-direction: column;
          padding: 24px 0;
        }
        .nav-main { flex: 1; }
        .sidebar-account {
          border-top: 1px solid rgba(255,255,255,0.2);
          margin-top: 10px;
          padding-top: 8px;
        }
        .sidebar-email {
          padding: 4px 20px 6px;
          font-size: 11.5px;
          color: rgba(255,255,255,0.75);
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
        }
        .brand {
          padding: 0 20px 20px 20px;
          border-bottom: 1px solid rgba(255,255,255,0.25);
          margin-bottom: 12px;
        }
        .brand-mark {
          font-family: 'IBM Plex Mono', monospace;
          font-size: 11px;
          letter-spacing: 0.14em;
          color: #ffffff;
          opacity: 0.85;
          text-transform: uppercase;
        }
        .brand-name {
          font-size: 18px;
          font-weight: 600;
          margin-top: 4px;
          letter-spacing: -0.01em;
          color: #ffffff;
        }
        .nav-item {
          display: flex;
          align-items: center;
          gap: 10px;
          padding: 10px 20px;
          font-size: 14px;
          font-weight: 500;
          color: rgba(255,255,255,0.85);
          cursor: pointer;
          border-left: 3px solid transparent;
          transition: background 0.12s ease, color 0.12s ease;
        }
        .nav-item:hover { background: rgba(255,255,255,0.12); color: #fff; }
        .nav-item.active {
          color: #fff;
          background: rgba(255,255,255,0.16);
          border-left-color: #ffffff;
        }

        .main { flex: 1; padding: 32px 40px 60px; max-width: 1100px; }

        .save-banner {
          display: flex; align-items: center; gap: 8px;
          background: var(--rust-bg); color: var(--rust);
          font-size: 13px; padding: 8px 12px; border-radius: 4px; margin-bottom: 16px;
        }

        h1.page-title {
          font-size: 24px; font-weight: 600; letter-spacing: -0.01em; margin: 0 0 4px 0;
        }
        .page-sub { color: var(--ink-soft); font-size: 14px; margin-bottom: 28px; }

        .btn {
          display: inline-flex; align-items: center; gap: 6px;
          font-family: inherit; font-size: 13.5px; font-weight: 600;
          padding: 9px 14px; border-radius: 5px; border: 1px solid transparent;
          cursor: pointer; transition: all 0.12s ease; line-height: 1;
        }
        .btn-primary { background: var(--brass); color: #fff; }
        .btn-primary:hover { background: var(--brass-dark); }
        .btn-ghost { background: transparent; color: var(--ink); border-color: var(--line); }
        .btn-ghost:hover { background: var(--paper-raised); }
        .btn-danger { background: transparent; color: var(--rust); border-color: var(--rust-bg); }
        .btn-danger:hover { background: var(--rust-bg); }
        .btn-sm { padding: 6px 10px; font-size: 12.5px; }
        .icon-btn {
          display: inline-flex; align-items: center; justify-content: center;
          width: 30px; height: 30px; border-radius: 5px; border: 1px solid var(--line);
          background: var(--paper-raised); color: var(--ink-soft); cursor: pointer;
        }
        .icon-btn:hover { color: var(--ink); border-color: var(--ink-faint); }

        .cards-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 32px; }
        .stat-card {
          background: var(--paper-raised); border: 1px solid var(--paper-line);
          border-radius: 8px; padding: 16px 18px;
        }
        .stat-label {
          font-size: 11px; text-transform: uppercase; letter-spacing: 0.08em;
          color: var(--ink-soft); font-weight: 600; margin-bottom: 8px;
        }
        .stat-value { font-size: 22px; font-weight: 600; }

        .section-title {
          font-size: 13px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.06em;
          color: var(--ink-soft); margin: 0 0 12px 0;
        }

        table.ledger-table { width: 100%; border-collapse: collapse; }
        table.ledger-table thead th {
          text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: 0.06em;
          color: var(--ink-soft); font-weight: 600; padding: 0 12px 10px 12px;
          border-bottom: 1px solid var(--line);
        }
        table.ledger-table tbody td {
          padding: 12px; font-size: 14px; border-bottom: 1px solid var(--paper-line);
        }
        table.ledger-table tbody tr { cursor: pointer; }
        table.ledger-table tbody tr:hover { background: var(--paper-raised); }
        table.ledger-table tbody tr:last-child td { border-bottom: none; }

        .badge {
          display: inline-flex; align-items: center; padding: 3px 9px; border-radius: 20px;
          font-size: 11.5px; font-weight: 600; font-family: 'IBM Plex Mono', monospace;
          text-transform: uppercase; letter-spacing: 0.03em;
        }
        .badge-paid { background: var(--sage-bg); color: var(--sage); }
        .badge-partial { background: var(--amber-bg); color: var(--amber); }
        .badge-unpaid { background: var(--rust-bg); color: var(--rust); }
        .badge-neutral { background: var(--paper-line); color: var(--ink-soft); }

        .toolbar { display: flex; justify-content: space-between; align-items: center; gap: 12px; margin-bottom: 20px; }
        .search-box {
          display: flex; align-items: center; gap: 8px; background: var(--paper-raised);
          border: 1px solid var(--line); border-radius: 6px; padding: 7px 10px; width: 260px;
        }
        .search-box input { border: none; background: transparent; outline: none; font-size: 13.5px; width: 100%; color: var(--ink); }

        .empty-state {
          border: 1px dashed var(--line); border-radius: 8px; padding: 48px 24px;
          text-align: center; color: var(--ink-soft);
        }
        .empty-state h3 { color: var(--ink); font-size: 16px; margin: 12px 0 4px; }
        .empty-state p { font-size: 13.5px; margin: 0 0 16px; }

        .modal-overlay {
          position: fixed; inset: 0; background: rgba(20,26,30,0.45);
          display: flex; align-items: flex-start; justify-content: center;
          padding: 40px 20px; overflow-y: auto; z-index: 50;
        }
        .modal-card {
          background: #fff; border-radius: 10px; width: 100%; max-width: 560px;
          box-shadow: 0 20px 50px rgba(0,0,0,0.2); overflow: hidden;
        }
        .modal-card.wide { max-width: 720px; }
        .modal-header {
          display: flex; align-items: center; justify-content: space-between;
          padding: 18px 22px; border-bottom: 1px solid var(--paper-line);
        }
        .modal-header h2 { font-size: 16px; font-weight: 600; margin: 0; }
        .modal-body { padding: 22px; max-height: 70vh; overflow-y: auto; }
        .modal-footer {
          display: flex; justify-content: flex-end; gap: 10px;
          padding: 16px 22px; border-top: 1px solid var(--paper-line); background: var(--paper-raised);
        }

        .field { margin-bottom: 14px; }
        .field label {
          display: block; font-size: 12px; font-weight: 600; color: var(--ink-soft);
          text-transform: uppercase; letter-spacing: 0.04em; margin-bottom: 6px;
        }
        .field input, .field select, .field textarea {
          width: 100%; padding: 9px 10px; border: 1px solid var(--line); border-radius: 5px;
          font-family: inherit; font-size: 14px; color: var(--ink); background: #fff;
        }
        .field input:focus, .field select:focus, .field textarea:focus {
          outline: 2px solid var(--brass); outline-offset: 0; border-color: var(--brass);
        }
        .field-row { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }

        .items-table { width: 100%; border-collapse: collapse; margin-bottom: 10px; }
        .items-table th {
          font-size: 10.5px; text-transform: uppercase; color: var(--ink-soft); text-align: left;
          padding: 0 6px 6px; letter-spacing: 0.05em;
        }
        .items-table td { padding: 4px 6px; vertical-align: middle; }
        .items-table input { width: 100%; padding: 7px 8px; border: 1px solid var(--line); border-radius: 4px; font-size: 13.5px; font-family: inherit; }
        .items-table .qty-input, .items-table .price-input { text-align: right; font-family: 'IBM Plex Mono', monospace; }
        .row-remove { color: var(--ink-faint); cursor: pointer; background: none; border: none; padding: 6px; }
        .row-remove:hover { color: var(--rust); }

        .totals-box { margin-top: 10px; border-top: 1px solid var(--paper-line); padding-top: 10px; }
        .totals-line { display: flex; justify-content: space-between; font-size: 13.5px; padding: 3px 0; color: var(--ink-soft); }
        .totals-line.grand { color: var(--ink); font-weight: 700; font-size: 15px; padding-top: 6px; border-top: 1px solid var(--paper-line); margin-top: 4px; }

        .detail-header { display: flex; align-items: flex-start; justify-content: space-between; margin-bottom: 24px; }
        .back-link { display: inline-flex; align-items: center; gap: 6px; color: var(--ink-soft); font-size: 13px; cursor: pointer; margin-bottom: 14px; }
        .back-link:hover { color: var(--ink); }

        .detail-grid { display: grid; grid-template-columns: 1fr 320px; gap: 24px; }
        .panel { background: var(--paper-raised); border: 1px solid var(--paper-line); border-radius: 8px; padding: 20px; margin-bottom: 16px; }
        .panel h3.panel-title { font-size: 13px; text-transform: uppercase; letter-spacing: 0.05em; color: var(--ink-soft); margin: 0 0 14px; }

        .kv-row { display: flex; justify-content: space-between; padding: 6px 0; font-size: 13.5px; border-bottom: 1px solid var(--paper-line); }
        .kv-row:last-child { border-bottom: none; }
        .kv-row span:first-child { color: var(--ink-soft); }

        .status-select { font-family: inherit; font-size: 13px; padding: 6px 8px; border-radius: 5px; border: 1px solid var(--line); background: #fff; }

        /* Invoice / print styling — cream sheet with accent header, modeled on user's own template */
        .invoice-sheet {
          --iv-accent: #2C6CB0;
          background: #F7F3E1;
          width: 100%; max-width: 680px; margin: 0 auto;
          font-family: 'IBM Plex Sans', sans-serif; color: #1F2A2E;
          position: relative; overflow: hidden;
          border-radius: 4px;
          box-shadow: 0 1px 0 rgba(0,0,0,0.04);
        }
        .invoice-bar { height: 14px; background: var(--iv-accent); }
        .invoice-inner { padding: 40px 48px 44px; }

        .invoice-top { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 22px; }
        .invoice-logo-row { display: flex; align-items: center; gap: 10px; }
        .invoice-logo-icon {
          width: 40px; height: 40px; border: 1.6px solid #1F2A2E; border-radius: 3px;
          display: flex; align-items: center; justify-content: center; flex-shrink: 0;
        }
        .invoice-brand-text .name { font-size: 17px; font-weight: 800; letter-spacing: 0.01em; line-height: 1.1; }
        .invoice-brand-text .name .accent { color: var(--iv-accent); }
        .invoice-brand-text .tagline { font-size: 9.5px; letter-spacing: 0.08em; color: #6B7574; font-weight: 600; text-transform: uppercase; margin-top: 3px; }
        .invoice-title { font-size: 34px; font-weight: 300; letter-spacing: 0.06em; color: var(--iv-accent); text-transform: uppercase; }

        .invoice-hr { border: none; border-top: 1.5px solid #DAD3B8; margin: 0 0 18px; }

        .invoice-meta-row { display: flex; justify-content: flex-end; margin-bottom: 22px; }
        .invoice-meta { text-align: right; font-size: 12.5px; color: #4B5453; line-height: 1.8; }
        .invoice-meta .num { font-weight: 700; color: #1F2A2E; }

        .invoice-parties { display: flex; justify-content: space-between; margin-bottom: 26px; gap: 24px; }
        .invoice-parties .party { max-width: 260px; }
        .invoice-parties.right-align .party { text-align: right; }
        .invoice-parties .label { font-size: 11px; color: #6B7574; margin-bottom: 4px; }
        .invoice-parties .name { font-size: 14.5px; font-weight: 700; color: var(--iv-accent); }
        .invoice-parties .detail { font-size: 12.5px; color: #4B5453; line-height: 1.55; }

        table.invoice-items { width: 100%; border-collapse: collapse; margin-bottom: 4px; }
        table.invoice-items thead th {
          text-align: left; font-size: 11.5px; text-transform: uppercase; letter-spacing: 0.04em;
          color: #fff; background: var(--iv-accent); padding: 10px 10px; font-weight: 700;
        }
        table.invoice-items thead th:first-child { border-radius: 3px 0 0 3px; }
        table.invoice-items thead th:last-child { border-radius: 0 3px 3px 0; }
        table.invoice-items td { padding: 12px 10px; font-size: 13.5px; border-bottom: 1px solid #DAD3B8; }
        table.invoice-items tr:last-child td { border-bottom: 1.5px solid #DAD3B8; }
        table.invoice-items td.num, table.invoice-items th.num { text-align: right; }

        .invoice-totals { margin-left: auto; width: 220px; margin-top: 14px; }
        .invoice-totals .line { display: flex; justify-content: space-between; font-size: 13px; padding: 4px 0; color: #4B5453; }
        .invoice-totals .line.grand { color: var(--iv-accent); font-weight: 700; font-size: 19px; padding-top: 8px; }
        .invoice-totals .line.balance { color: #B0503B; font-weight: 700; }

        .invoice-pay-block { margin-top: 30px; font-size: 13px; line-height: 1.9; }
        .invoice-pay-block .ref { color: var(--iv-accent); font-weight: 700; font-size: 14px; margin-bottom: 4px; }
        .invoice-pay-block b { font-weight: 700; }
        .invoice-vat-note { margin-top: 16px; font-size: 12px; color: #6B7574; line-height: 1.6; }

        .invoice-footer { display: flex; justify-content: space-between; align-items: flex-end; margin-top: 34px; }
        .invoice-thanks { font-size: 15px; font-weight: 800; }
        .invoice-sig { text-align: right; font-size: 12.5px; }
        .invoice-sig .sig-name { font-weight: 700; text-transform: uppercase; letter-spacing: 0.02em; }
        .invoice-sig .sig-links { color: #6B7574; margin-top: 3px; }

        .badge-inline {
          display: inline-block; font-size: 10px; font-weight: 700; text-transform: uppercase;
          letter-spacing: 0.04em; padding: 2px 8px; border-radius: 10px; margin-left: 8px; vertical-align: middle;
        }
        .badge-inline.paid { background: var(--sage-bg); color: var(--sage); }
        .badge-inline.partial { background: var(--amber-bg); color: var(--amber); }
        .badge-inline.unpaid { background: var(--rust-bg); color: var(--rust); }

        @media print {
          body * { visibility: hidden; }
          #invoice-printable, #invoice-printable * { visibility: visible; }
          #invoice-printable { position: absolute; top: 0; left: 0; width: 100%; border: none; }
        }

        @media (max-width: 720px) {
          .app-grid { flex-direction: column; }
          .sidebar { width: 100%; flex-direction: row; overflow-x: auto; padding: 10px; }
          .brand { display: none; }
          .nav-item { border-left: none; border-bottom: 3px solid transparent; white-space: nowrap; }
          .nav-item.active { border-left: none; border-bottom-color: var(--amber); }
          .main { padding: 20px; }
          .cards-row { grid-template-columns: 1fr 1fr; }
          .detail-grid { grid-template-columns: 1fr; }
        }
      `}</style>
      {children}
    </div>
  );
}

/* ---------------------------------------------------------------
   Confirm dialog (native confirm() is blocked in the artifact sandbox)
--------------------------------------------------------------- */

function ConfirmModal({ title, message, confirmLabel = "Delete", onConfirm, onCancel }) {
  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div className="modal-card" style={{ maxWidth: 420 }} onClick={(e) => e.stopPropagation()}>
        <div className="modal-header">
          <h2>{title}</h2>
          <button className="icon-btn" onClick={onCancel}><X size={15} /></button>
        </div>
        <div className="modal-body">
          <p style={{ margin: 0, fontSize: 14, color: "var(--ink-soft)" }}>{message}</p>
        </div>
        <div className="modal-footer">
          <button className="btn btn-ghost" onClick={onCancel}>Cancel</button>
          <button className="btn btn-danger" style={{ background: "var(--rust)", color: "#fff", borderColor: "var(--rust)" }} onClick={onConfirm}>
            {confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   Login
--------------------------------------------------------------- */

function LoginScreen() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setLoading(true);
    setError(null);
    const { error: err } = await supabase.auth.signInWithPassword({ email, password });
    if (err) setError(err.message);
    setLoading(false);
  };

  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "center", minHeight: "100vh" }}>
      <form onSubmit={handleSubmit} className="modal-card" style={{ width: 360, padding: "32px 28px" }}>
        <div style={{ marginBottom: 20 }}>
          <div className="brand-mark" style={{ color: "var(--brass)" }}>Liveable Layouts</div>
          <h2 style={{ margin: "4px 0 0", fontSize: 19 }}>Sign in</h2>
        </div>
        {error && (
          <div className="save-banner" style={{ marginBottom: 14 }}>
            <AlertCircle size={14} /> {error}
          </div>
        )}
        <div className="field">
          <label>Email</label>
          <input type="email" required autoFocus value={email} onChange={(e) => setEmail(e.target.value)} placeholder="you@business.com" />
        </div>
        <div className="field">
          <label>Password</label>
          <input type="password" required value={password} onChange={(e) => setPassword(e.target.value)} placeholder="••••••••" />
        </div>
        <button className="btn btn-primary" type="submit" disabled={loading} style={{ width: "100%", justifyContent: "center", marginTop: 6 }}>
          {loading ? "Signing in…" : "Sign In"}
        </button>
        <p style={{ fontSize: 12, color: "var(--ink-soft)", marginTop: 14, marginBottom: 0 }}>
          Don't have an account? Ask your admin to add you in Supabase (Authentication → Users).
        </p>
      </form>
    </div>
  );
}

/* ---------------------------------------------------------------
   Sidebar
--------------------------------------------------------------- */

function Sidebar({ view, setView, userEmail, onSignOut }) {
  const items = [
    { key: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { key: "customers", label: "Customers", icon: Users },
    { key: "orders", label: "Orders & Invoices", icon: FileText },
  ];
  return (
    <nav className="sidebar">
      <div className="brand">
        <div className="brand-mark">Orders & Invoicing</div>
        <div className="brand-name">Liveable Layouts</div>
      </div>
      <div className="nav-main">
        {items.map((it) => (
          <div
            key={it.key}
            className={`nav-item ${view === it.key ? "active" : ""}`}
            onClick={() => setView(it.key)}
          >
            <it.icon size={16} />
            {it.label}
          </div>
        ))}
      </div>
      <div
        className={`nav-item ${view === "settings" ? "active" : ""}`}
        onClick={() => setView("settings")}
      >
        <Settings size={16} />
        Business Details
      </div>
      <div className="sidebar-account">
        {userEmail && <div className="sidebar-email">{userEmail}</div>}
        <div className="nav-item" onClick={onSignOut}>
          <LogOut size={16} />
          Sign Out
        </div>
      </div>
    </nav>
  );
}

/* ---------------------------------------------------------------
   Settings / Business Profile
--------------------------------------------------------------- */

function SettingsView({ business, onSave }) {
  const [form, setForm] = useState({ ...business });
  const [savedFlash, setSavedFlash] = useState(false);
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });

  useEffect(() => setForm({ ...business }), [business]);

  const save = () => {
    onSave(form);
    setSavedFlash(true);
    setTimeout(() => setSavedFlash(false), 1800);
  };

  return (
    <div>
      <h1 className="page-title">Business Details</h1>
      <p className="page-sub">This appears on every invoice you generate.</p>

      <div className="panel" style={{ maxWidth: 560 }}>
        <div className="field">
          <label>Business Name</label>
          <input value={form.name} onChange={set("name")} placeholder="Liveable Layouts" />
        </div>
        <div className="field">
          <label>Tagline</label>
          <input value={form.tagline} onChange={set("tagline")} placeholder="Designed to work in real life" />
        </div>
        <div className="field-row">
          <div className="field">
            <label>Email</label>
            <input value={form.email} onChange={set("email")} placeholder="you@business.com" />
          </div>
          <div className="field">
            <label>Accent Color</label>
            <div style={{ display: "flex", gap: 8 }}>
              <input type="color" style={{ width: 44, padding: 2 }} value={form.accent} onChange={set("accent")} />
              <input value={form.accent} onChange={set("accent")} />
            </div>
          </div>
        </div>
        <div className="field">
          <label>Address</label>
          <textarea rows={2} value={form.address} onChange={set("address")} placeholder="Cloghmore, Ballynahown, Co. Galway" />
        </div>
        <div className="field-row">
          <div className="field">
            <label>Payment Terms</label>
            <input value={form.paymentTerms} onChange={set("paymentTerms")} placeholder="Due before work commences" />
          </div>
          <div className="field">
            <label>Payment Method</label>
            <input value={form.paymentMethod} onChange={set("paymentMethod")} placeholder="Revolut (payment link provided)" />
          </div>
        </div>
        <div className="field">
          <label>VAT / Tax Note</label>
          <input value={form.vatNote} onChange={set("vatNote")} placeholder="Not currently registered for VAT" />
        </div>
        <div className="field-row">
          <div className="field">
            <label>Signature Name</label>
            <input value={form.signatureName} onChange={set("signatureName")} placeholder="Jessica Doe" />
          </div>
          <div className="field">
            <label>Social / Website Links</label>
            <input value={form.socialLinks} onChange={set("socialLinks")} placeholder="instagram.com/yourbusiness" />
          </div>
        </div>
        <button className="btn btn-primary" onClick={save}>Save Details</button>
        {savedFlash && <span style={{ marginLeft: 10, fontSize: 13, color: "var(--sage)" }}>Saved</span>}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   Dashboard
--------------------------------------------------------------- */

function Dashboard({ orders, customers, onOpenOrder }) {
  const stats = useMemo(() => {
    let revenue = 0;
    let outstanding = 0;
    let active = 0;
    orders.forEach((o) => {
      const t = computeTotals(o);
      revenue += t.paid;
      if (o.orderStatus !== "cancelled") outstanding += t.balance;
      if (o.orderStatus === "draft" || o.orderStatus === "sent") active += 1;
    });
    return { revenue, outstanding, active, customers: customers.length };
  }, [orders, customers]);

  const recent = [...orders]
    .sort((a, b) => (b.date || "").localeCompare(a.date || ""))
    .slice(0, 6);

  const customerName = (id) => customers.find((c) => c.id === id)?.name || "—";

  return (
    <div>
      <h1 className="page-title">Dashboard</h1>
      <p className="page-sub">Where things stand today, {fmtDate(todayISO())}.</p>

      <div className="cards-row">
        <div className="stat-card">
          <div className="stat-label">Revenue Collected</div>
          <div className="stat-value mono">{fmtMoney(stats.revenue)}</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Outstanding</div>
          <div className="stat-value mono" style={{ color: stats.outstanding > 0 ? "var(--rust)" : "var(--ink)" }}>
            {fmtMoney(stats.outstanding)}
          </div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Active Orders</div>
          <div className="stat-value mono">{stats.active}</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Customers</div>
          <div className="stat-value mono">{stats.customers}</div>
        </div>
      </div>

      <div className="section-title">Recent Orders</div>
      {recent.length === 0 ? (
        <div className="empty-state">
          <Receipt size={22} style={{ margin: "0 auto", color: "var(--ink-faint)" }} />
          <h3>No orders yet</h3>
          <p>Once you create an order, it'll show up here.</p>
        </div>
      ) : (
        <table className="ledger-table">
          <thead>
            <tr>
              <th>Order</th>
              <th>Customer</th>
              <th>Date</th>
              <th>Status</th>
              <th>Payment</th>
              <th style={{ textAlign: "right" }}>Total</th>
            </tr>
          </thead>
          <tbody>
            {recent.map((o) => {
              const t = computeTotals(o);
              return (
                <tr key={o.id} onClick={() => onOpenOrder(o.id)}>
                  <td className="mono">{o.orderNumber}</td>
                  <td>{customerName(o.customerId)}</td>
                  <td className="mono">{fmtDate(o.date)}</td>
                  <td><span className="badge badge-neutral">{STATUS_LABEL[o.orderStatus]}</span></td>
                  <td><PaymentBadge status={t.paymentStatus} /></td>
                  <td className="mono" style={{ textAlign: "right" }}>{fmtMoney(t.total)}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      )}
    </div>
  );
}

function PaymentBadge({ status }) {
  const map = {
    paid: ["badge-paid", "Paid"],
    partial: ["badge-partial", "Partial"],
    unpaid: ["badge-unpaid", "Unpaid"],
  };
  const [cls, label] = map[status] || map.unpaid;
  return <span className={`badge ${cls}`}>{label}</span>;
}

/* ---------------------------------------------------------------
   Customers
--------------------------------------------------------------- */

function CustomersView({ customers, orders, onAdd, onUpdate, onDelete, onImport }) {
  const [query, setQuery] = useState("");
  const [modalCustomer, setModalCustomer] = useState(undefined); // undefined = closed, null = new, obj = edit
  const [pendingDelete, setPendingDelete] = useState(null);
  const [importing, setImporting] = useState(false);
  const [importResult, setImportResult] = useState(null);
  const [pendingWorkbook, setPendingWorkbook] = useState(null); // { workbook, fileName } when picking a sheet
  const fileInputRef = React.useRef(null);

  const filtered = customers.filter((c) =>
    (c.name + " " + c.email).toLowerCase().includes(query.toLowerCase())
  );

  const orderCount = (id) => orders.filter((o) => o.customerId === id).length;

  const runImportForSheet = async (workbook, sheetName) => {
    setImporting(true);
    setImportResult(null);
    try {
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
      const result = await onImport(rows);
      setImportResult(result);
    } catch (err) {
      setImportResult({ failed: true, message: err.message || "Couldn't read that sheet." });
    } finally {
      setImporting(false);
      setPendingWorkbook(null);
    }
  };

  const handleFileChange = async (e) => {
    const file = e.target.files?.[0];
    e.target.value = ""; // allow re-selecting the same file later
    if (!file) return;
    setImportResult(null);
    try {
      const buf = await file.arrayBuffer();
      const workbook = XLSX.read(buf, { type: "array" });
      if (workbook.SheetNames.length > 1) {
        // Ask which sheet to use instead of guessing
        setPendingWorkbook({ workbook, fileName: file.name });
      } else {
        await runImportForSheet(workbook, workbook.SheetNames[0]);
      }
    } catch (err) {
      setImportResult({ failed: true, message: err.message || "Couldn't read that file." });
    }
  };

  return (
    <div>
      <h1 className="page-title">Customers</h1>
      <p className="page-sub">Everyone you've billed, in one place.</p>

      {importResult && (
        <div
          className="save-banner"
          style={
            importResult.failed
              ? {}
              : { background: "var(--sage-bg)", color: "var(--sage)" }
          }
        >
          {importResult.failed ? (
            <>
              <AlertCircle size={14} /> {importResult.message || "Import failed."}
            </>
          ) : (
            <>
              <AlertCircle size={14} /> Imported: {importResult.added} added, {importResult.updated} updated
              {importResult.skipped > 0 ? `, ${importResult.skipped} skipped (no name)` : ""}.
            </>
          )}
        </div>
      )}

      <div className="toolbar">
        <div className="search-box">
          <Search size={14} color="var(--ink-soft)" />
          <input placeholder="Search customers…" value={query} onChange={(e) => setQuery(e.target.value)} />
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls,.csv"
            style={{ display: "none" }}
            onChange={handleFileChange}
          />
          <button
            className="btn btn-ghost"
            onClick={() => fileInputRef.current?.click()}
            disabled={importing}
            style={importing ? { opacity: 0.5, cursor: "not-allowed" } : {}}
          >
            <Upload size={14} /> {importing ? "Importing…" : "Import Excel"}
          </button>
          <button
            className="btn btn-ghost"
            onClick={() => exportCustomersToExcel(customers, orders)}
            disabled={customers.length === 0}
            style={customers.length === 0 ? { opacity: 0.5, cursor: "not-allowed" } : {}}
          >
            <Download size={14} /> Export Excel
          </button>
          <button className="btn btn-primary" onClick={() => setModalCustomer(null)}>
            <Plus size={15} /> Add Customer
          </button>
        </div>
      </div>

      {filtered.length === 0 ? (
        <div className="empty-state">
          <Users size={22} style={{ margin: "0 auto", color: "var(--ink-faint)" }} />
          <h3>{customers.length === 0 ? "No customers yet" : "No matches"}</h3>
          <p>{customers.length === 0 ? "Add your first customer to start creating orders." : "Try a different search term."}</p>
          {customers.length === 0 && (
            <button className="btn btn-primary" onClick={() => setModalCustomer(null)}>
              <Plus size={15} /> Add Customer
            </button>
          )}
        </div>
      ) : (
        <table className="ledger-table">
          <thead>
            <tr>
              <th>Name</th>
              <th>Email</th>
              <th>Phone</th>
              <th>Orders</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((c) => (
              <tr key={c.id} onClick={() => setModalCustomer(c)}>
                <td style={{ fontWeight: 600 }}>{c.name}</td>
                <td>{c.email || "—"}</td>
                <td className="mono">{c.phone || "—"}</td>
                <td className="mono">{orderCount(c.id)}</td>
                <td onClick={(e) => e.stopPropagation()} style={{ textAlign: "right" }}>
                  <button
                    className="icon-btn"
                    onClick={() => setPendingDelete(c)}
                  >
                    <Trash2 size={14} />
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}

      {modalCustomer !== undefined && (
        <CustomerModal
          customer={modalCustomer}
          onClose={() => setModalCustomer(undefined)}
          onSave={(vals) => {
            if (modalCustomer) onUpdate(modalCustomer.id, vals);
            else onAdd(vals);
            setModalCustomer(undefined);
          }}
        />
      )}

      {pendingDelete && (
        <ConfirmModal
          title="Delete customer"
          message={`Delete ${pendingDelete.name}? This won't delete their orders.`}
          onCancel={() => setPendingDelete(null)}
          onConfirm={() => {
            onDelete(pendingDelete.id);
            setPendingDelete(null);
          }}
        />
      )}

      {pendingWorkbook && (
        <div className="modal-overlay" onClick={() => setPendingWorkbook(null)}>
          <div className="modal-card" style={{ maxWidth: 420 }} onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <h2>Which sheet?</h2>
              <button className="icon-btn" onClick={() => setPendingWorkbook(null)}><X size={15} /></button>
            </div>
            <div className="modal-body">
              <p style={{ marginTop: 0, fontSize: 13.5, color: "var(--ink-soft)" }}>
                <strong>{pendingWorkbook.fileName}</strong> has {pendingWorkbook.workbook.SheetNames.length} sheets. Pick the one with your customer list.
              </p>
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                {pendingWorkbook.workbook.SheetNames.map((name) => (
                  <button
                    key={name}
                    className="btn btn-ghost"
                    style={{ justifyContent: "flex-start" }}
                    onClick={() => runImportForSheet(pendingWorkbook.workbook, name)}
                  >
                    <FileSpreadsheet size={14} /> {name}
                  </button>
                ))}
              </div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-ghost" onClick={() => setPendingWorkbook(null)}>Cancel</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function CustomerModal({ customer, onClose, onSave }) {
  const [form, setForm] = useState({
    name: customer?.name || "",
    email: customer?.email || "",
    phone: customer?.phone || "",
    address: customer?.address || "",
  });
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });
  const canSave = form.name.trim().length > 0;

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-card" onClick={(e) => e.stopPropagation()}>
        <div className="modal-header">
          <h2>{customer ? "Edit Customer" : "Add Customer"}</h2>
          <button className="icon-btn" onClick={onClose}><X size={15} /></button>
        </div>
        <div className="modal-body">
          <div className="field">
            <label>Name</label>
            <input autoFocus value={form.name} onChange={set("name")} placeholder="Acme Roasting Co." />
          </div>
          <div className="field-row">
            <div className="field">
              <label>Email</label>
              <input value={form.email} onChange={set("email")} placeholder="billing@acme.com" />
            </div>
            <div className="field">
              <label>Phone</label>
              <input value={form.phone} onChange={set("phone")} placeholder="(555) 010-2200" />
            </div>
          </div>
          <div className="field">
            <label>Billing Address</label>
            <textarea rows={3} value={form.address} onChange={set("address")} placeholder="123 Market St, Suite 4&#10;Portland, OR 97201" />
          </div>
        </div>
        <div className="modal-footer">
          <button className="btn btn-ghost" onClick={onClose}>Cancel</button>
          <button className="btn btn-primary" disabled={!canSave} style={!canSave ? { opacity: 0.5, cursor: "not-allowed" } : {}} onClick={() => canSave && onSave(form)}>
            Save Customer
          </button>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   Orders list
--------------------------------------------------------------- */

function OrdersListView({ orders, customers, onAdd, onOpen, onGoToCustomers }) {
  const [query, setQuery] = useState("");
  const [showNew, setShowNew] = useState(false);

  const handleNewOrderClick = () => {
    if (customers.length === 0) {
      onGoToCustomers();
    } else {
      setShowNew(true);
    }
  };

  const customerName = (id) => customers.find((c) => c.id === id)?.name || "Unknown";

  const filtered = orders.filter((o) =>
    (o.orderNumber + " " + customerName(o.customerId)).toLowerCase().includes(query.toLowerCase())
  );
  const sorted = [...filtered].sort((a, b) => (b.date || "").localeCompare(a.date || ""));

  return (
    <div>
      <h1 className="page-title">Orders & Invoices</h1>
      <p className="page-sub">Track every order from draft to paid.</p>

      <div className="toolbar">
        <div className="search-box">
          <Search size={14} color="var(--ink-soft)" />
          <input placeholder="Search orders…" value={query} onChange={(e) => setQuery(e.target.value)} />
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button
            className="btn btn-ghost"
            onClick={() => exportOrdersToExcel(sorted, customers)}
            disabled={sorted.length === 0}
            style={sorted.length === 0 ? { opacity: 0.5, cursor: "not-allowed" } : {}}
          >
            <Download size={14} /> Export Excel
          </button>
          <button className="btn btn-primary" onClick={handleNewOrderClick}>
            <Plus size={15} /> New Order
          </button>
        </div>
      </div>

      {customers.length === 0 && (
        <div className="empty-state" style={{ marginBottom: 20 }}>
          <p style={{ margin: 0 }}>You'll need a customer before creating your first order.</p>
          <button className="btn btn-primary" style={{ marginTop: 12 }} onClick={onGoToCustomers}>
            <Plus size={15} /> Add a Customer
          </button>
        </div>
      )}

      {sorted.length === 0 && customers.length > 0 ? (
        <div className="empty-state">
          <Receipt size={22} style={{ margin: "0 auto", color: "var(--ink-faint)" }} />
          <h3>{orders.length === 0 ? "No orders yet" : "No matches"}</h3>
          <p>{orders.length === 0 ? "Create your first order to start invoicing." : "Try a different search term."}</p>
          {orders.length === 0 && (
            <button className="btn btn-primary" onClick={() => setShowNew(true)}>
              <Plus size={15} /> New Order
            </button>
          )}
        </div>
      ) : sorted.length > 0 && (
        <table className="ledger-table">
          <thead>
            <tr>
              <th>Order</th>
              <th>Customer</th>
              <th>Date</th>
              <th>Status</th>
              <th>Payment</th>
              <th style={{ textAlign: "right" }}>Total</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {sorted.map((o) => {
              const t = computeTotals(o);
              return (
                <tr key={o.id} onClick={() => onOpen(o.id)}>
                  <td className="mono">{o.orderNumber}</td>
                  <td>{customerName(o.customerId)}</td>
                  <td className="mono">{fmtDate(o.date)}</td>
                  <td><span className="badge badge-neutral">{STATUS_LABEL[o.orderStatus]}</span></td>
                  <td><PaymentBadge status={t.paymentStatus} /></td>
                  <td className="mono" style={{ textAlign: "right" }}>{fmtMoney(t.total)}</td>
                  <td style={{ color: "var(--ink-faint)" }}><ChevronRight size={16} /></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      )}

      {showNew && (
        <OrderModal
          customers={customers}
          onClose={() => setShowNew(false)}
          onSave={(vals) => {
            onAdd(vals);
            setShowNew(false);
          }}
        />
      )}
    </div>
  );
}

function blankItem() {
  return { id: uid(), description: "", qty: 1, price: 0 };
}

function OrderModal({ customers, order, onClose, onSave }) {
  const [form, setForm] = useState({
    customerId: order?.customerId || customers[0]?.id || "",
    date: order?.date || todayISO(),
    dueDate: order?.dueDate || "",
    taxRate: order?.taxRate ?? 0,
    notes: order?.notes || "",
    items: order?.items?.length ? order.items.map((i) => ({ ...i })) : [blankItem()],
  });

  const setField = (k, v) => setForm((f) => ({ ...f, [k]: v }));
  const setItem = (id, patch) =>
    setForm((f) => ({ ...f, items: f.items.map((it) => (it.id === id ? { ...it, ...patch } : it)) }));
  const addItem = () => setForm((f) => ({ ...f, items: [...f.items, blankItem()] }));
  const removeItem = (id) => setForm((f) => ({ ...f, items: f.items.filter((it) => it.id !== id) }));

  const totals = useMemo(() => computeTotals({ items: form.items, taxRate: form.taxRate, amountPaid: 0, orderStatus: "draft" }), [form.items, form.taxRate]);

  const canSave = form.customerId && form.items.some((i) => i.description.trim());

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-card wide" onClick={(e) => e.stopPropagation()}>
        <div className="modal-header">
          <h2>{order ? "Edit Order" : "New Order"}</h2>
          <button className="icon-btn" onClick={onClose}><X size={15} /></button>
        </div>
        <div className="modal-body">
          <div className="field-row">
            <div className="field">
              <label>Customer</label>
              <select value={form.customerId} onChange={(e) => setField("customerId", e.target.value)}>
                {customers.map((c) => (
                  <option key={c.id} value={c.id}>{c.name}</option>
                ))}
              </select>
            </div>
            <div className="field">
              <label>Order Date</label>
              <input type="date" value={form.date} onChange={(e) => setField("date", e.target.value)} />
            </div>
          </div>
          <div className="field-row">
            <div className="field">
              <label>Due Date</label>
              <input type="date" value={form.dueDate} onChange={(e) => setField("dueDate", e.target.value)} />
            </div>
            <div className="field">
              <label>Tax Rate (%)</label>
              <input type="number" min="0" step="0.1" value={form.taxRate} onChange={(e) => setField("taxRate", e.target.value)} />
            </div>
          </div>

          <div className="field">
            <label style={{ marginBottom: 8 }}>Line Items</label>
            <table className="items-table">
              <thead>
                <tr>
                  <th style={{ width: "50%" }}>Description</th>
                  <th style={{ width: "15%" }}>Qty</th>
                  <th style={{ width: "20%" }}>Price</th>
                  <th style={{ width: "10%" }}></th>
                </tr>
              </thead>
              <tbody>
                {form.items.map((it) => (
                  <tr key={it.id}>
                    <td>
                      <input
                        value={it.description}
                        onChange={(e) => setItem(it.id, { description: e.target.value })}
                        placeholder="Item or service"
                      />
                    </td>
                    <td>
                      <input
                        className="qty-input"
                        type="number"
                        min="0"
                        step="1"
                        value={it.qty}
                        onChange={(e) => setItem(it.id, { qty: e.target.value })}
                      />
                    </td>
                    <td>
                      <input
                        className="price-input"
                        type="number"
                        min="0"
                        step="0.01"
                        value={it.price}
                        onChange={(e) => setItem(it.id, { price: e.target.value })}
                      />
                    </td>
                    <td>
                      <button className="row-remove" onClick={() => removeItem(it.id)} disabled={form.items.length === 1}>
                        <Trash2 size={14} />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            <button className="btn btn-ghost btn-sm" onClick={addItem}>
              <Plus size={13} /> Add Line
            </button>
          </div>

          <div className="totals-box">
            <div className="totals-line"><span>Subtotal</span><span className="mono">{fmtMoney(totals.subtotal)}</span></div>
            <div className="totals-line"><span>Tax ({Number(form.taxRate) || 0}%)</span><span className="mono">{fmtMoney(totals.taxAmt)}</span></div>
            <div className="totals-line grand"><span>Total</span><span className="mono">{fmtMoney(totals.total)}</span></div>
          </div>

          <div className="field" style={{ marginTop: 14 }}>
            <label>Notes</label>
            <textarea rows={2} value={form.notes} onChange={(e) => setField("notes", e.target.value)} placeholder="Payment terms, delivery notes, etc." />
          </div>
        </div>
        <div className="modal-footer">
          <button className="btn btn-ghost" onClick={onClose}>Cancel</button>
          <button
            className="btn btn-primary"
            disabled={!canSave}
            style={!canSave ? { opacity: 0.5, cursor: "not-allowed" } : {}}
            onClick={() => {
              if (!canSave) return;
              const cleanItems = form.items
                .filter((i) => i.description.trim())
                .map((i) => ({ ...i, qty: Number(i.qty) || 0, price: Number(i.price) || 0 }));
              onSave({ ...form, items: cleanItems, taxRate: Number(form.taxRate) || 0 });
            }}
          >
            Save Order
          </button>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   Order Detail
--------------------------------------------------------------- */

function OrderDetailView({ order, customers, onBack, onUpdate, onDelete, onGenerateInvoice }) {
  const [editing, setEditing] = useState(false);
  const [paymentInput, setPaymentInput] = useState(String(order.amountPaid || 0));
  const [confirmingDelete, setConfirmingDelete] = useState(false);
  const customer = customers.find((c) => c.id === order.customerId);
  const totals = computeTotals(order);

  useEffect(() => {
    setPaymentInput(String(order.amountPaid || 0));
  }, [order.id]);

  return (
    <div>
      <div className="back-link" onClick={onBack}>
        <ArrowLeft size={14} /> Back to orders
      </div>
      <div className="detail-header">
        <div>
          <h1 className="page-title mono">{order.orderNumber}</h1>
          <p className="page-sub" style={{ marginBottom: 0 }}>
            {customer?.name || "Unknown customer"} · {fmtDate(order.date)}
          </p>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button className="btn btn-ghost" onClick={() => setEditing(true)}>
            <Pencil size={14} /> Edit
          </button>
          <button className="btn btn-primary" onClick={onGenerateInvoice}>
            <Printer size={14} /> {order.invoiceNumber ? "View Invoice" : "Generate Invoice"}
          </button>
          <button
            className="btn btn-danger"
            onClick={() => setConfirmingDelete(true)}
          >
            <Trash2 size={14} />
          </button>
        </div>
      </div>

      <div className="detail-grid">
        <div>
          <div className="panel">
            <h3 className="panel-title">Line Items</h3>
            <table className="ledger-table">
              <thead>
                <tr>
                  <th>Description</th>
                  <th style={{ textAlign: "right" }}>Qty</th>
                  <th style={{ textAlign: "right" }}>Price</th>
                  <th style={{ textAlign: "right" }}>Amount</th>
                </tr>
              </thead>
              <tbody>
                {order.items.map((it) => (
                  <tr key={it.id} style={{ cursor: "default" }}>
                    <td>{it.description}</td>
                    <td className="mono" style={{ textAlign: "right" }}>{it.qty}</td>
                    <td className="mono" style={{ textAlign: "right" }}>{fmtMoney(it.price)}</td>
                    <td className="mono" style={{ textAlign: "right" }}>{fmtMoney(it.qty * it.price)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div className="totals-box">
              <div className="totals-line"><span>Subtotal</span><span className="mono">{fmtMoney(totals.subtotal)}</span></div>
              <div className="totals-line"><span>Tax ({order.taxRate || 0}%)</span><span className="mono">{fmtMoney(totals.taxAmt)}</span></div>
              <div className="totals-line grand"><span>Total</span><span className="mono">{fmtMoney(totals.total)}</span></div>
            </div>
          </div>
          {order.notes && (
            <div className="panel">
              <h3 className="panel-title">Notes</h3>
              <p style={{ fontSize: 13.5, color: "var(--ink-soft)", margin: 0, whiteSpace: "pre-wrap" }}>{order.notes}</p>
            </div>
          )}
        </div>

        <div>
          <div className="panel">
            <h3 className="panel-title">Status</h3>
            <div className="field">
              <label>Order Status</label>
              <select
                className="status-select"
                style={{ width: "100%" }}
                value={order.orderStatus}
                onChange={(e) => onUpdate({ orderStatus: e.target.value })}
              >
                {Object.entries(STATUS_LABEL).map(([k, v]) => (
                  <option key={k} value={k}>{v}</option>
                ))}
              </select>
            </div>
            <div className="kv-row">
              <span>Payment</span>
              <PaymentBadge status={totals.paymentStatus} />
            </div>
            <div className="kv-row">
              <span>Due</span>
              <span className="mono">{fmtDate(order.dueDate)}</span>
            </div>
            <div className="kv-row">
              <span>Invoice #</span>
              <span className="mono">{order.invoiceNumber || "Not issued"}</span>
            </div>
          </div>

          <div className="panel">
            <h3 className="panel-title">Record Payment</h3>
            <div className="kv-row">
              <span>Total Due</span>
              <span className="mono">{fmtMoney(totals.total)}</span>
            </div>
            <div className="kv-row">
              <span>Balance</span>
              <span className="mono" style={{ color: totals.balance > 0 ? "var(--rust)" : "var(--sage)" }}>{fmtMoney(totals.balance)}</span>
            </div>
            <div className="field" style={{ marginTop: 12, marginBottom: 8 }}>
              <label>Amount Paid</label>
              <input
                type="number"
                min="0"
                step="0.01"
                value={paymentInput}
                onChange={(e) => setPaymentInput(e.target.value)}
              />
            </div>
            <button
              className="btn btn-primary btn-sm"
              style={{ width: "100%", justifyContent: "center" }}
              onClick={() => onUpdate({ amountPaid: Number(paymentInput) || 0 })}
            >
              Update Payment
            </button>
          </div>

          {customer && (
            <div className="panel">
              <h3 className="panel-title">Bill To</h3>
              <div style={{ fontSize: 13.5, fontWeight: 600, marginBottom: 4 }}>{customer.name}</div>
              <div style={{ fontSize: 13, color: "var(--ink-soft)" }}>{customer.email}</div>
              <div style={{ fontSize: 13, color: "var(--ink-soft)" }}>{customer.phone}</div>
              {customer.address && (
                <div style={{ fontSize: 13, color: "var(--ink-soft)", whiteSpace: "pre-wrap", marginTop: 4 }}>{customer.address}</div>
              )}
            </div>
          )}
        </div>
      </div>

      {editing && (
        <OrderModal
          customers={customers}
          order={order}
          onClose={() => setEditing(false)}
          onSave={(vals) => {
            onUpdate(vals);
            setEditing(false);
          }}
        />
      )}

      {confirmingDelete && (
        <ConfirmModal
          title="Delete order"
          message={`Delete order ${order.orderNumber}? This can't be undone.`}
          onCancel={() => setConfirmingDelete(false)}
          onConfirm={() => {
            setConfirmingDelete(false);
            onDelete();
          }}
        />
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   Invoice Modal (printable)
--------------------------------------------------------------- */

function InvoiceModal({ order, customer, business, onClose }) {
  const totals = computeTotals(order);
  const b = business || {};
  const accent = b.accent || "#2C6CB0";
  const statusLabel = { paid: "Paid", partial: "Partial paid", unpaid: "Unpaid" }[totals.paymentStatus];

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div style={{ width: "100%", maxWidth: 760 }} onClick={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginBottom: 12 }}>
          <button className="btn btn-ghost" style={{ background: "#fff" }} onClick={onClose}>
            <X size={14} /> Close
          </button>
          <button className="btn btn-primary" onClick={() => window.print()}>
            <Printer size={14} /> Print / Save PDF
          </button>
        </div>

        <div className="invoice-sheet" id="invoice-printable" style={{ "--iv-accent": accent }}>
          <div className="invoice-bar" />
          <div className="invoice-inner">
            <div className="invoice-top">
              <div className="invoice-logo-row">
                <div className="invoice-logo-icon">
                  <Home size={20} color={accent} strokeWidth={1.6} />
                </div>
                <div className="invoice-brand-text">
                  <div className="name">{b.name || "Your Business"}</div>
                  {b.tagline && <div className="tagline">{b.tagline}</div>}
                </div>
              </div>
              <div className="invoice-title">Invoice</div>
            </div>

            <hr className="invoice-hr" />

            <div className="invoice-meta-row">
              <div className="invoice-meta">
                <div><b>Invoice No:</b> <span className="num">{order.invoiceNumber}</span></div>
                <div>Invoice Date: {fmtDate(order.date)}</div>
                <div>Due: {order.dueDate ? fmtDate(order.dueDate) : (b.paymentTerms || "—")}</div>
              </div>
            </div>

            <div className="invoice-parties">
              <div className="party">
                <div className="label">Issued by:</div>
                <div className="name">{b.name || "Your Business"}</div>
                {b.address && <div className="detail" style={{ whiteSpace: "pre-wrap" }}>{b.address}</div>}
                {b.email && <div className="detail">Email: {b.email}</div>}
              </div>
              <div className="party" style={{ textAlign: "right" }}>
                <div className="label">Bill to:</div>
                <div className="name">{customer?.name || "—"}</div>
                {customer?.email && <div className="detail">Email: {customer.email}</div>}
                {customer?.phone && <div className="detail">{customer.phone}</div>}
                {customer?.address && <div className="detail" style={{ whiteSpace: "pre-wrap" }}>{customer.address}</div>}
              </div>
            </div>

            <table className="invoice-items">
              <thead>
                <tr>
                  <th>Service</th>
                  <th className="num">Qty</th>
                  <th className="num">Price</th>
                  <th className="num">Total</th>
                </tr>
              </thead>
              <tbody>
                {order.items.map((it) => (
                  <tr key={it.id}>
                    <td>{it.description}</td>
                    <td className="num">{it.qty}</td>
                    <td className="num">{fmtMoney(it.price)}</td>
                    <td className="num">{fmtMoney(it.qty * it.price)}</td>
                  </tr>
                ))}
              </tbody>
            </table>

            <div className="invoice-totals">
              <div className="line"><span>Subtotal:</span><span>{fmtMoney(totals.subtotal)}</span></div>
              <div className="line"><span>Tax:</span><span>{order.taxRate ? `${fmtMoney(totals.taxAmt)}` : "N/A"}</span></div>
              <div className="line grand"><span>Total:</span><span>{fmtMoney(totals.total)}</span></div>
              {totals.paid > 0 && (
                <>
                  <div className="line"><span>Paid:</span><span>{fmtMoney(totals.paid)}</span></div>
                  <div className="line balance"><span>Balance Due:</span><span>{fmtMoney(totals.balance)}</span></div>
                </>
              )}
            </div>

            <div className="invoice-pay-block">
              <div className="ref">
                Payment Reference: {order.invoiceNumber}
                <span className={`badge-inline ${totals.paymentStatus}`}>{statusLabel}</span>
              </div>
              {b.paymentTerms && <div><b>Payment Terms:</b> {b.paymentTerms}</div>}
              {b.paymentMethod && <div><b>Payment Method:</b> {b.paymentMethod}</div>}
            </div>

            {order.notes && (
              <div className="invoice-vat-note"><b>Notes:</b> {order.notes}</div>
            )}
            {b.vatNote && <div className="invoice-vat-note">{b.vatNote}</div>}

            <div className="invoice-footer">
              <div className="invoice-thanks">Thank you for business!</div>
              {(b.signatureName || b.socialLinks) && (
                <div className="invoice-sig">
                  {b.signatureName && <div className="sig-name">{b.signatureName}</div>}
                  {b.socialLinks && <div className="sig-links">{b.socialLinks}</div>}
                </div>
              )}
            </div>
          </div>
          <div className="invoice-bar" />
        </div>
      </div>
    </div>
  );
}
