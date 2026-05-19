import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

const CATEGORIES_INCOME = ["Gaji", "Bonus", "Freelance", "Investasi", "Penjualan", "Lainnya"];
const CATEGORIES_EXPENSE = ["Makan & Minum", "Transport", "Belanja", "Tagihan", "Kesehatan", "Hiburan", "Pendidikan", "Lainnya"];

const formatRp = (n) =>
  new Intl.NumberFormat("id-ID", { style: "currency", currency: "IDR", minimumFractionDigits: 0 }).format(n || 0);

const today = () => new Date().toISOString().split("T")[0];

const STORAGE_KEY = "keuangan_data_v1";

export default function App() {
  const [transactions, setTransactions] = useState([]);
  const [view, setView] = useState("dashboard"); // dashboard | form | history | export
  const [type, setType] = useState("income");
  const [form, setForm] = useState({ amount: "", category: "", note: "", date: today() });
  const [filterMonth, setFilterMonth] = useState(new Date().toISOString().slice(0, 7));
  const [toast, setToast] = useState(null);
  const [deleteId, setDeleteId] = useState(null);

  // Load from storage
  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) setTransactions(JSON.parse(raw));
    } catch {}
  }, []);

  // Save to storage
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(transactions));
    } catch {}
  }, [transactions]);

  const showToast = (msg, color = "#22c55e") => {
    setToast({ msg, color });
    setTimeout(() => setToast(null), 2200);
  };

  const filtered = transactions.filter((t) => t.date.startsWith(filterMonth));
  const totalIn = filtered.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0);
  const totalOut = filtered.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0);
  const balance = totalIn - totalOut;

  const addTransaction = () => {
    const amt = parseFloat(form.amount.replace(/\D/g, ""));
    if (!amt || amt <= 0) return showToast("Masukkan jumlah yang valid", "#ef4444");
    if (!form.category) return showToast("Pilih kategori", "#ef4444");
    const tx = {
      id: Date.now(),
      type,
      amount: amt,
      category: form.category,
      note: form.note,
      date: form.date,
    };
    setTransactions((prev) => [tx, ...prev]);
    setForm({ amount: "", category: "", note: "", date: today() });
    setView("dashboard");
    showToast(type === "income" ? "✅ Pemasukan dicatat!" : "✅ Pengeluaran dicatat!");
  };

  const deleteTransaction = (id) => {
    setTransactions((prev) => prev.filter((t) => t.id !== id));
    setDeleteId(null);
    showToast("🗑️ Transaksi dihapus", "#f59e0b");
  };

  const exportExcel = () => {
    const months = [...new Set(transactions.map((t) => t.date.slice(0, 7)))].sort().reverse();
    const wb = XLSX.utils.book_new();

    // Summary sheet
    const summaryData = [["Bulan", "Pemasukan", "Pengeluaran", "Saldo"]];
    months.forEach((m) => {
      const mTx = transactions.filter((t) => t.date.startsWith(m));
      const inc = mTx.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0);
      const exp = mTx.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0);
      summaryData.push([m, inc, exp, inc - exp]);
    });
    const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
    wsSummary["!cols"] = [{ wch: 12 }, { wch: 18 }, { wch: 18 }, { wch: 18 }];
    XLSX.utils.book_append_sheet(wb, wsSummary, "Ringkasan");

    // All transactions sheet
    const rows = [["Tanggal", "Jenis", "Kategori", "Jumlah", "Catatan"]];
    [...transactions].sort((a, b) => b.date.localeCompare(a.date)).forEach((t) => {
      rows.push([t.date, t.type === "income" ? "Pemasukan" : "Pengeluaran", t.category, t.amount, t.note]);
    });
    const wsAll = XLSX.utils.aoa_to_sheet(rows);
    wsAll["!cols"] = [{ wch: 12 }, { wch: 14 }, { wch: 18 }, { wch: 18 }, { wch: 30 }];
    XLSX.utils.book_append_sheet(wb, wsAll, "Semua Transaksi");

    // Per-month sheets
    months.slice(0, 6).forEach((m) => {
      const mTx = transactions.filter((t) => t.date.startsWith(m)).sort((a, b) => b.date.localeCompare(a.date));
      const mRows = [["Tanggal", "Jenis", "Kategori", "Jumlah", "Catatan"]];
      mTx.forEach((t) => mRows.push([t.date, t.type === "income" ? "Pemasukan" : "Pengeluaran", t.category, t.amount, t.note]));
      const inc = mTx.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0);
      const exp = mTx.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0);
      mRows.push([], ["", "Total Pemasukan", "", inc, ""], ["", "Total Pengeluaran", "", exp, ""], ["", "Saldo", "", inc - exp, ""]);
      const ws = XLSX.utils.aoa_to_sheet(mRows);
      ws["!cols"] = [{ wch: 12 }, { wch: 14 }, { wch: 18 }, { wch: 18 }, { wch: 30 }];
      XLSX.utils.book_append_sheet(wb, ws, m);
    });

    XLSX.writeFile(wb, `Keuangan_${new Date().toISOString().slice(0, 10)}.xlsx`);
    showToast("📊 File Excel berhasil diunduh!");
  };

  const fmtAmt = (v) => {
    const raw = v.replace(/\D/g, "");
    if (!raw) return "";
    return parseInt(raw, 10).toLocaleString("id-ID");
  };

  const recentTx = transactions.slice(0, 5);

  return (
    <div style={{ fontFamily: "'Nunito', sans-serif", background: "#0f0f14", minHeight: "100vh", maxWidth: 430, margin: "0 auto", color: "#f0f0f5", position: "relative", paddingBottom: 80 }}>
      <link href="https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800;900&display=swap" rel="stylesheet" />

      {/* Toast */}
      {toast && (
        <div style={{ position: "fixed", top: 20, left: "50%", transform: "translateX(-50%)", background: toast.color, color: "#fff", borderRadius: 14, padding: "10px 22px", fontWeight: 700, zIndex: 999, fontSize: 14, boxShadow: "0 4px 20px rgba(0,0,0,0.4)", whiteSpace: "nowrap" }}>
          {toast.msg}
        </div>
      )}

      {/* Delete Confirm Modal */}
      {deleteId && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.7)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 998 }}>
          <div style={{ background: "#1e1e2e", borderRadius: 20, padding: 28, width: 300, textAlign: "center" }}>
            <div style={{ fontSize: 40, marginBottom: 10 }}>🗑️</div>
            <div style={{ fontWeight: 800, fontSize: 17, marginBottom: 8 }}>Hapus transaksi?</div>
            <div style={{ color: "#888", fontSize: 13, marginBottom: 22 }}>Data ini akan dihapus permanen.</div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => setDeleteId(null)} style={{ flex: 1, padding: "11px", borderRadius: 12, border: "1.5px solid #333", background: "transparent", color: "#aaa", fontWeight: 700, cursor: "pointer" }}>Batal</button>
              <button onClick={() => deleteTransaction(deleteId)} style={{ flex: 1, padding: "11px", borderRadius: 12, border: "none", background: "#ef4444", color: "#fff", fontWeight: 700, cursor: "pointer" }}>Hapus</button>
            </div>
          </div>
        </div>
      )}

      {/* Header */}
      <div style={{ background: "linear-gradient(135deg, #6366f1 0%, #8b5cf6 60%, #a855f7 100%)", padding: "28px 22px 60px", borderRadius: "0 0 36px 36px", position: "relative", overflow: "hidden" }}>
        <div style={{ position: "absolute", top: -40, right: -40, width: 150, height: 150, background: "rgba(255,255,255,0.06)", borderRadius: "50%" }} />
        <div style={{ position: "absolute", bottom: -20, left: -30, width: 120, height: 120, background: "rgba(255,255,255,0.05)", borderRadius: "50%" }} />
        <div style={{ fontWeight: 900, fontSize: 20, letterSpacing: 0.5, marginBottom: 4 }}>💰 KasKu</div>
        <div style={{ fontSize: 12, opacity: 0.75, marginBottom: 22 }}>Pencatat Keuangan Pribadi</div>
        <div style={{ background: "rgba(255,255,255,0.12)", backdropFilter: "blur(10px)", borderRadius: 20, padding: "18px 22px" }}>
          <div style={{ fontSize: 12, opacity: 0.8, marginBottom: 4 }}>Saldo Bulan Ini</div>
          <div style={{ fontWeight: 900, fontSize: 32, letterSpacing: -1 }}>{formatRp(balance)}</div>
          <div style={{ display: "flex", gap: 20, marginTop: 14 }}>
            <div>
              <div style={{ fontSize: 11, opacity: 0.75 }}>↑ Pemasukan</div>
              <div style={{ fontWeight: 800, fontSize: 15, color: "#86efac" }}>{formatRp(totalIn)}</div>
            </div>
            <div style={{ width: 1, background: "rgba(255,255,255,0.2)" }} />
            <div>
              <div style={{ fontSize: 11, opacity: 0.75 }}>↓ Pengeluaran</div>
              <div style={{ fontWeight: 800, fontSize: 15, color: "#fca5a5" }}>{formatRp(totalOut)}</div>
            </div>
          </div>
        </div>
        {/* Month filter */}
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 14 }}>
          <span style={{ fontSize: 12, opacity: 0.8 }}>📅 Bulan:</span>
          <input type="month" value={filterMonth} onChange={(e) => setFilterMonth(e.target.value)}
            style={{ background: "rgba(255,255,255,0.15)", border: "none", borderRadius: 10, padding: "6px 12px", color: "#fff", fontFamily: "inherit", fontWeight: 700, fontSize: 13, cursor: "pointer" }} />
        </div>
      </div>

      {/* Content */}
      <div style={{ padding: "0 16px", marginTop: -20 }}>

        {/* DASHBOARD */}
        {view === "dashboard" && (
          <>
            {/* Quick Actions */}
            <div style={{ display: "flex", gap: 10, marginBottom: 20 }}>
              <button onClick={() => { setType("income"); setView("form"); }} style={{ flex: 1, background: "linear-gradient(135deg, #22c55e, #16a34a)", border: "none", borderRadius: 16, padding: "16px", color: "#fff", fontWeight: 800, fontSize: 15, cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", gap: 4 }}>
                <span style={{ fontSize: 24 }}>➕</span>
                <span>Pemasukan</span>
              </button>
              <button onClick={() => { setType("expense"); setView("form"); }} style={{ flex: 1, background: "linear-gradient(135deg, #ef4444, #dc2626)", border: "none", borderRadius: 16, padding: "16px", color: "#fff", fontWeight: 800, fontSize: 15, cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", gap: 4 }}>
                <span style={{ fontSize: 24 }}>➖</span>
                <span>Pengeluaran</span>
              </button>
            </div>

            {/* Recent */}
            <div style={{ background: "#1a1a26", borderRadius: 20, padding: "18px 16px", marginBottom: 16 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
                <span style={{ fontWeight: 800, fontSize: 15 }}>Transaksi Terbaru</span>
                <button onClick={() => setView("history")} style={{ background: "none", border: "none", color: "#8b5cf6", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>Lihat Semua →</button>
              </div>
              {recentTx.length === 0 ? (
                <div style={{ textAlign: "center", padding: "24px 0", color: "#555", fontSize: 14 }}>
                  <div style={{ fontSize: 36, marginBottom: 8 }}>📭</div>
                  Belum ada transaksi
                </div>
              ) : recentTx.map((t) => (
                <TxItem key={t.id} t={t} onDelete={() => setDeleteId(t.id)} />
              ))}
            </div>

            {/* Export Button */}
            <button onClick={exportExcel} style={{ width: "100%", background: "#1a1a26", border: "1.5px solid #2d2d40", borderRadius: 16, padding: "14px", color: "#a78bfa", fontWeight: 800, fontSize: 14, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>
              <span style={{ fontSize: 20 }}>📊</span> Export ke Excel (.xlsx)
            </button>
          </>
        )}

        {/* FORM */}
        {view === "form" && (
          <div style={{ background: "#1a1a26", borderRadius: 20, padding: "22px 18px", marginTop: 4 }}>
            <div style={{ fontWeight: 900, fontSize: 17, marginBottom: 18, display: "flex", alignItems: "center", gap: 10 }}>
              <span style={{ fontSize: 26 }}>{type === "income" ? "📥" : "📤"}</span>
              Tambah {type === "income" ? "Pemasukan" : "Pengeluaran"}
            </div>

            {/* Type Toggle */}
            <div style={{ display: "flex", background: "#12121c", borderRadius: 14, padding: 4, marginBottom: 18 }}>
              {["income", "expense"].map((t) => (
                <button key={t} onClick={() => { setType(t); setForm((f) => ({ ...f, category: "" })); }}
                  style={{ flex: 1, padding: "10px", borderRadius: 10, border: "none", background: type === t ? (t === "income" ? "#22c55e" : "#ef4444") : "transparent", color: type === t ? "#fff" : "#666", fontWeight: 800, cursor: "pointer", transition: "all 0.2s", fontSize: 13 }}>
                  {t === "income" ? "➕ Pemasukan" : "➖ Pengeluaran"}
                </button>
              ))}
            </div>

            <label style={labelStyle}>Jumlah (Rp)</label>
            <div style={{ position: "relative", marginBottom: 14 }}>
              <span style={{ position: "absolute", left: 14, top: "50%", transform: "translateY(-50%)", color: "#6366f1", fontWeight: 800 }}>Rp</span>
              <input value={form.amount} onChange={(e) => setForm((f) => ({ ...f, amount: fmtAmt(e.target.value) }))}
                placeholder="0" inputMode="numeric"
                style={{ ...inputStyle, paddingLeft: 44, fontSize: 20, fontWeight: 800 }} />
            </div>

            <label style={labelStyle}>Kategori</label>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 16 }}>
              {(type === "income" ? CATEGORIES_INCOME : CATEGORIES_EXPENSE).map((c) => (
                <button key={c} onClick={() => setForm((f) => ({ ...f, category: c }))}
                  style={{ padding: "8px 14px", borderRadius: 10, border: form.category === c ? "none" : "1.5px solid #2d2d40", background: form.category === c ? "#6366f1" : "transparent", color: form.category === c ? "#fff" : "#888", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>
                  {c}
                </button>
              ))}
            </div>

            <label style={labelStyle}>Catatan (Opsional)</label>
            <input value={form.note} onChange={(e) => setForm((f) => ({ ...f, note: e.target.value }))}
              placeholder="Contoh: Makan siang bersama klien"
              style={{ ...inputStyle, marginBottom: 14 }} />

            <label style={labelStyle}>Tanggal</label>
            <input type="date" value={form.date} onChange={(e) => setForm((f) => ({ ...f, date: e.target.value }))}
              style={{ ...inputStyle, marginBottom: 22 }} />

            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => setView("dashboard")} style={{ flex: 1, padding: 14, borderRadius: 14, border: "1.5px solid #2d2d40", background: "transparent", color: "#888", fontWeight: 700, cursor: "pointer" }}>Batal</button>
              <button onClick={addTransaction} style={{ flex: 2, padding: 14, borderRadius: 14, border: "none", background: type === "income" ? "linear-gradient(135deg,#22c55e,#16a34a)" : "linear-gradient(135deg,#ef4444,#dc2626)", color: "#fff", fontWeight: 800, fontSize: 15, cursor: "pointer" }}>
                Simpan
              </button>
            </div>
          </div>
        )}

        {/* HISTORY */}
        {view === "history" && (
          <div style={{ background: "#1a1a26", borderRadius: 20, padding: "18px 16px", marginTop: 4 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16 }}>
              <button onClick={() => setView("dashboard")} style={{ background: "none", border: "none", color: "#8b5cf6", fontWeight: 800, fontSize: 18, cursor: "pointer" }}>←</button>
              <span style={{ fontWeight: 900, fontSize: 16 }}>Semua Transaksi</span>
              <span style={{ marginLeft: "auto", fontSize: 13, color: "#555" }}>{filtered.length} data</span>
            </div>
            {filtered.length === 0 ? (
              <div style={{ textAlign: "center", padding: "32px 0", color: "#555" }}>
                <div style={{ fontSize: 40 }}>📭</div>
                <div style={{ marginTop: 8 }}>Tidak ada transaksi di bulan ini</div>
              </div>
            ) : [...filtered].sort((a, b) => b.date.localeCompare(a.date)).map((t) => (
              <TxItem key={t.id} t={t} onDelete={() => setDeleteId(t.id)} />
            ))}
          </div>
        )}
      </div>

      {/* Bottom Nav */}
      <nav style={{ position: "fixed", bottom: 0, left: "50%", transform: "translateX(-50%)", width: "100%", maxWidth: 430, background: "#13131f", borderTop: "1px solid #1e1e2e", display: "flex", padding: "10px 0 16px" }}>
        {[
          { id: "dashboard", icon: "🏠", label: "Beranda" },
          { id: "form", icon: "➕", label: "Catat" },
          { id: "history", icon: "📋", label: "Riwayat" },
        ].map((nav) => (
          <button key={nav.id} onClick={() => { if (nav.id === "form") { setType("income"); } setView(nav.id); }}
            style={{ flex: 1, background: "none", border: "none", display: "flex", flexDirection: "column", alignItems: "center", gap: 3, cursor: "pointer", opacity: view === nav.id ? 1 : 0.4, transition: "opacity 0.2s" }}>
            <span style={{ fontSize: 22 }}>{nav.icon}</span>
            <span style={{ fontSize: 11, fontWeight: 700, color: view === nav.id ? "#8b5cf6" : "#fff" }}>{nav.label}</span>
          </button>
        ))}
      </nav>
    </div>
  );
}

function TxItem({ t, onDelete }) {
  const isIn = t.type === "income";
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "12px 0", borderBottom: "1px solid #1e1e2e" }}>
      <div style={{ width: 42, height: 42, borderRadius: 14, background: isIn ? "rgba(34,197,94,0.15)" : "rgba(239,68,68,0.15)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, flexShrink: 0 }}>
        {isIn ? "📥" : "📤"}
      </div>
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontWeight: 800, fontSize: 14 }}>{t.category}</div>
        <div style={{ fontSize: 12, color: "#555", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.note || t.date}</div>
      </div>
      <div style={{ textAlign: "right" }}>
        <div style={{ fontWeight: 900, fontSize: 14, color: isIn ? "#86efac" : "#fca5a5" }}>
          {isIn ? "+" : "-"}{formatRp(t.amount)}
        </div>
        <div style={{ fontSize: 11, color: "#444" }}>{t.date}</div>
      </div>
      <button onClick={onDelete} style={{ background: "none", border: "none", color: "#444", cursor: "pointer", fontSize: 18, padding: "0 4px" }}>×</button>
    </div>
  );
}

const inputStyle = {
  width: "100%", background: "#12121c", border: "1.5px solid #2d2d40", borderRadius: 14,
  padding: "13px 16px", color: "#f0f0f5", fontFamily: "inherit", fontWeight: 600, fontSize: 15,
  boxSizing: "border-box", outline: "none",
};

const labelStyle = { display: "block", fontSize: 12, fontWeight: 700, color: "#666", marginBottom: 8, textTransform: "uppercase", letterSpacing: 0.5 };