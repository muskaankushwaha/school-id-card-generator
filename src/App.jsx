
import { useState, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── Utility helpers ────────────────────────────────────────────────────────
const uid = () => Math.random().toString(36).slice(2, 9);

const fileToDataURL = (file) =>
  new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result);
    r.onerror = rej;
    r.readAsDataURL(file);
  });

const emptyStudent = () => ({
  id: uid(),
  schoolName: "",
  studentName: "",
  fatherName: "",
  motherName: "",
  dob: "",
  course: "",
  mobile: "",
  aadhaar: "",
  address: "",
  schoolContact: "",
  rollNumber: "",
  photo: null,
  principalSig: null,
  schoolLogo: null,
});

// ─── ID Card Component ───────────────────────────────────────────────────────
function IDCard({ student, scale = 1, forExport = false }) {
  const cardStyle = {
    width: 340,
    minHeight: 540,
    transform: `scale(${scale})`,
    transformOrigin: "top left",
    fontFamily: "'Georgia', serif",
  };

  return (
    <div
      className="id-card-root"
      style={cardStyle}
      data-card-id={student.id}
    >
      {/* Header */}
      <div
        style={{
          background: "linear-gradient(135deg,#1a3a5c 0%,#0d6efd 100%)",
          borderRadius: "12px 12px 0 0",
          padding: "14px 12px 10px",
          display: "flex",
          alignItems: "center",
          gap: 10,
        }}
      >
        {student.schoolLogo ? (
          <img
            src={student.schoolLogo}
            alt="logo"
            style={{ width: 52, height: 52, borderRadius: 8, objectFit: "contain", background: "#fff", padding: 2 }}
          />
        ) : (
          <div
            style={{
              width: 52, height: 52, borderRadius: 8, background: "rgba(255,255,255,0.18)",
              display: "flex", alignItems: "center", justifyContent: "center",
              fontSize: 22, color: "#fff",
            }}
          >🏫</div>
        )}
        <div style={{ flex: 1 }}>
          <div style={{ color: "#fff", fontWeight: 700, fontSize: 13, lineHeight: 1.3, textShadow: "0 1px 3px rgba(0,0,0,.3)" }}>
            {student.schoolName || "School / College Name"}
          </div>
          <div style={{ color: "rgba(255,255,255,.75)", fontSize: 10, marginTop: 2 }}>
            STUDENT IDENTITY CARD
          </div>
        </div>
      </div>

      {/* Body */}
      <div
        style={{
          background: "#fff",
          border: "1.5px solid #d0e4f7",
          borderTop: "none",
          padding: "16px 14px",
        }}
      >
        {/* Photo + Name */}
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", marginBottom: 14 }}>
          <div
            style={{
              width: 90, height: 110, borderRadius: 8,
              background: "#eaf1fb", overflow: "hidden",
              border: "3px solid #0d6efd", boxShadow: "0 2px 12px rgba(13,110,253,.18)",
            }}
          >
            {student.photo ? (
              <img src={student.photo} alt="student" style={{ width: "100%", height: "100%", objectFit: "cover" }} />
            ) : (
              <div style={{ width: "100%", height: "100%", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 38, color: "#a0b8d8" }}>👤</div>
            )}
          </div>
          <div style={{ marginTop: 8, fontWeight: 700, fontSize: 15, color: "#1a3a5c", letterSpacing: 0.3, textAlign: "center" }}>
            {student.studentName || "Student Name"}
          </div>
          <div style={{ fontSize: 11, color: "#0d6efd", marginTop: 2 }}>
            {student.course || "Class / Course"}
          </div>
        </div>

        {/* Details */}
        <div
          style={{
            background: "#f0f7ff",
            borderRadius: 8,
            padding: "10px 12px",
            fontSize: 11,
            lineHeight: 1.8,
            color: "#2c3e50",
          }}
        >
          {[
            ["Father", student.fatherName],
            ["Mother", student.motherName],
            ["DOB", student.dob],
            ["Mobile", student.mobile],
            ["Aadhaar", student.aadhaar ? `XXXX-XXXX-${student.aadhaar.slice(-4)}` : ""],
            ["Roll No.", student.rollNumber],
            ["Address", student.address],
          ]
            .filter(([, v]) => v)
            .map(([label, val]) => (
              <div key={label} style={{ display: "flex", gap: 6 }}>
                <span style={{ fontWeight: 600, minWidth: 58, color: "#1a3a5c" }}>{label}:</span>
                <span style={{ flex: 1, wordBreak: "break-word" }}>{val}</span>
              </div>
            ))}
        </div>

        {/* Signature */}
        <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 12, paddingRight: 4 }}>
          <div style={{ textAlign: "center" }}>
            {student.principalSig ? (
              <img
                src={student.principalSig}
                alt="sig"
                style={{ height: 36, maxWidth: 100, objectFit: "contain", display: "block", margin: "0 auto" }}
              />
            ) : (
              <div style={{ height: 28, borderBottom: "1.5px solid #1a3a5c", width: 80, margin: "0 auto" }} />
            )}
            <div style={{ fontSize: 9, color: "#5a7a9a", marginTop: 3 }}>Principal's Signature</div>
          </div>
        </div>
      </div>

      {/* Footer */}
      <div
        style={{
          background: "linear-gradient(135deg,#1a3a5c 0%,#0d6efd 100%)",
          borderRadius: "0 0 12px 12px",
          padding: "7px 12px",
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
        }}
      >
        <div style={{ color: "rgba(255,255,255,.8)", fontSize: 9 }}>☎ {student.schoolContact || "Contact Number"}</div>
        <div style={{ color: "rgba(255,255,255,.55)", fontSize: 8 }}>VALID FOR CURRENT SESSION</div>
      </div>
    </div>
  );
}

// ─── Drag-Drop Zone ──────────────────────────────────────────────────────────
function DropZone({ label, accept, multiple, onFiles, icon }) {
  const [drag, setDrag] = useState(false);
  const inp = useRef();
  const handle = (e) => {
    e.preventDefault();
    setDrag(false);
    const files = Array.from(e.dataTransfer?.files || e.target.files || []);
    if (files.length) onFiles(files);
    if (inp.current) inp.current.value = "";
  };
  return (
    <div
      onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
      onDragLeave={() => setDrag(false)}
      onDrop={handle}
      onClick={() => inp.current?.click()}
      style={{
        border: `2px dashed ${drag ? "#0d6efd" : "#93c5fd"}`,
        borderRadius: 10,
        padding: "20px 16px",
        textAlign: "center",
        cursor: "pointer",
        background: drag ? "#eff6ff" : "#f8fbff",
        transition: "all .2s",
        userSelect: "none",
      }}
    >
      <div style={{ fontSize: 28 }}>{icon}</div>
      <div style={{ fontSize: 13, color: "#1a3a5c", fontWeight: 600, marginTop: 6 }}>{label}</div>
      <div style={{ fontSize: 11, color: "#6b9cc7", marginTop: 3 }}>Drag & drop or click to browse</div>
      <input ref={inp} type="file" accept={accept} multiple={multiple} onChange={handle} style={{ display: "none" }} />
    </div>
  );
}

// ─── Field ───────────────────────────────────────────────────────────────────
function Field({ label, value, onChange, type = "text", required }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
      <label style={{ fontSize: 12, fontWeight: 600, color: "#1a3a5c" }}>
        {label}{required && <span style={{ color: "#e74c3c" }}> *</span>}
      </label>
      <input
        type={type}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        required={required}
        style={{
          padding: "8px 12px",
          borderRadius: 8,
          border: "1.5px solid #cce0f5",
          fontSize: 13,
          color: "#1a3a5c",
          background: "#f8fbff",
          outline: "none",
          transition: "border .2s",
        }}
        onFocus={(e) => (e.target.style.border = "1.5px solid #0d6efd")}
        onBlur={(e) => (e.target.style.border = "1.5px solid #cce0f5")}
      />
    </div>
  );
}

// ─── Toast ────────────────────────────────────────────────────────────────────
function Toast({ msg, type, onClose }) {
  useEffect(() => {
    const t = setTimeout(onClose, 3500);
    return () => clearTimeout(t);
  }, []);
  const colors = { success: "#10b981", error: "#ef4444", info: "#0d6efd" };
  return (
    <div
      style={{
        position: "fixed", bottom: 24, right: 24, zIndex: 9999,
        background: colors[type] || "#333", color: "#fff",
        padding: "12px 20px", borderRadius: 10,
        boxShadow: "0 4px 20px rgba(0,0,0,.2)",
        fontSize: 13, fontWeight: 600, maxWidth: 320,
        animation: "slideUp .3s ease",
      }}
    >
      {msg}
    </div>
  );
}

// ─── PDF Export (pure canvas/print approach) ─────────────────────────────────
async function exportToPDF(students) {
  // Use html2canvas-like approach via a hidden iframe print
  const win = window.open("", "_blank", "width=900,height=700");
  if (!win) return;

  const cardsHtml = students
    .map((s) => {
      const photo = s.photo ? `<img src="${s.photo}" style="width:100%;height:100%;object-fit:cover" />` : "👤";
      const sig = s.principalSig ? `<img src="${s.principalSig}" style="height:36px;max-width:100px;object-fit:contain" />` : `<div style="height:28px;border-bottom:1.5px solid #1a3a5c;width:80px;margin:0 auto"></div>`;
      const logo = s.schoolLogo ? `<img src="${s.schoolLogo}" style="width:52px;height:52px;border-radius:8px;object-fit:contain;background:#fff;padding:2px" />` : `<div style="width:52px;height:52px;border-radius:8px;background:rgba(255,255,255,.18);display:flex;align-items:center;justify-content:center;font-size:22px">🏫</div>`;
      const rows = [
        ["Father", s.fatherName], ["Mother", s.motherName], ["DOB", s.dob],
        ["Mobile", s.mobile], ["Roll No.", s.rollNumber], ["Address", s.address],
      ].filter(([, v]) => v).map(([l, v]) => `<div style="display:flex;gap:6px;font-size:11px;line-height:1.8"><span style="font-weight:600;min-width:58px;color:#1a3a5c">${l}:</span><span>${v}</span></div>`).join("");

      return `
      <div style="width:340px;font-family:Georgia,serif;margin:8px;display:inline-block;vertical-align:top;box-shadow:0 4px 20px rgba(0,0,0,.15);border-radius:12px;overflow:hidden;page-break-inside:avoid">
        <div style="background:linear-gradient(135deg,#1a3a5c,#0d6efd);padding:14px 12px 10px;display:flex;align-items:center;gap:10px">
          ${logo}
          <div><div style="color:#fff;font-weight:700;font-size:13px">${s.schoolName || "School Name"}</div><div style="color:rgba(255,255,255,.75);font-size:10px;margin-top:2px">STUDENT IDENTITY CARD</div></div>
        </div>
        <div style="background:#fff;border:1.5px solid #d0e4f7;border-top:none;padding:16px 14px">
          <div style="text-align:center;margin-bottom:14px">
            <div style="width:90px;height:110px;border-radius:8px;background:#eaf1fb;overflow:hidden;border:3px solid #0d6efd;margin:0 auto">${photo}</div>
            <div style="margin-top:8px;font-weight:700;font-size:15px;color:#1a3a5c">${s.studentName || "Student Name"}</div>
            <div style="font-size:11px;color:#0d6efd;margin-top:2px">${s.course || ""}</div>
          </div>
          <div style="background:#f0f7ff;border-radius:8px;padding:10px 12px">${rows}</div>
          <div style="display:flex;justify-content:flex-end;margin-top:12px">
            <div style="text-align:center">${sig}<div style="font-size:9px;color:#5a7a9a;margin-top:3px">Principal's Signature</div></div>
          </div>
        </div>
        <div style="background:linear-gradient(135deg,#1a3a5c,#0d6efd);padding:7px 12px;display:flex;justify-content:space-between">
          <div style="color:rgba(255,255,255,.8);font-size:9px">☎ ${s.schoolContact || ""}</div>
          <div style="color:rgba(255,255,255,.55);font-size:8px">VALID FOR CURRENT SESSION</div>
        </div>
      </div>`;
    })
    .join("");

  win.document.write(`<!DOCTYPE html><html><head><title>ID Cards</title>
  <style>
    @page { size: A4; margin: 15mm; }
    body { font-family: Georgia, serif; }
    @media print { button { display:none; } }
  </style>
  </head><body>
  <div style="text-align:center;margin-bottom:16px">
    <h2 style="color:#1a3a5c;font-size:18px">Student ID Cards</h2>
    <p style="color:#666;font-size:12px">Total: ${students.length} cards</p>
  </div>
  <div style="display:flex;flex-wrap:wrap;justify-content:center;gap:16px">${cardsHtml}</div>
  <div style="text-align:center;margin-top:20px">
    <button onclick="window.print()" style="background:#0d6efd;color:#fff;border:none;padding:12px 32px;border-radius:8px;font-size:15px;cursor:pointer;font-weight:600">🖨️ Print / Save as PDF</button>
  </div>
  </body></html>`);
  win.document.close();
}

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function App() {
  const [tab, setTab] = useState("form"); // form | bulk | preview | manage
  const [darkMode, setDarkMode] = useState(false);
  const [students, setStudents] = useState([emptyStudent()]);
  const [activeIdx, setActiveIdx] = useState(0);
  const [search, setSearch] = useState("");
  const [toast, setToast] = useState(null);
  const [loading, setLoading] = useState(false);
  const [bulkPhotos, setBulkPhotos] = useState({});

  const showToast = (msg, type = "success") => setToast({ msg, type });

  const cur = students[activeIdx] || students[0];

  const updateField = (key) => (val) => {
    setStudents((prev) =>
      prev.map((s, i) => (i === activeIdx ? { ...s, [key]: val } : s))
    );
  };

  const handleImageField = async (key, files) => {
    if (!files[0]) return;
    const url = await fileToDataURL(files[0]);
    updateField(key)(url);
  };

  // ── Bulk CSV/Excel upload ────────────────────────────────────────────────
  const handleCSV = async (files) => {
    const file = files[0];
    if (!file) return;
    setLoading(true);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
      if (!rows.length) { showToast("No data found in file", "error"); return; }

      const mapped = rows.map((row) => {
        const get = (...keys) => keys.map((k) => row[k] || row[k.toLowerCase()] || "").find(Boolean) || "";
        return {
          ...emptyStudent(),
          schoolName: get("schoolName", "School Name", "school_name", "School"),
          studentName: get("studentName", "Student Name", "student_name", "Name"),
          fatherName: get("fatherName", "Father Name", "father_name", "Father"),
          motherName: get("motherName", "Mother Name", "mother_name", "Mother"),
          dob: get("dob", "DOB", "Date of Birth", "dateOfBirth"),
          course: get("course", "Course", "Class", "class"),
          mobile: get("mobile", "Mobile", "Phone"),
          aadhaar: get("aadhaar", "Aadhaar"),
          address: get("address", "Address"),
          schoolContact: get("schoolContact", "School Contact", "school_contact"),
          rollNumber: get("rollNumber", "Roll Number", "roll_no", "Roll No"),
        };
      });

      setStudents(mapped);
      setActiveIdx(0);
      showToast(`✅ Imported ${mapped.length} student records!`);
    } catch (e) {
      showToast("Failed to parse file: " + e.message, "error");
    } finally {
      setLoading(false);
    }
  };

  // ── Bulk photo upload ────────────────────────────────────────────────────
  const handleBulkPhotos = async (files) => {
    setLoading(true);
    const map = {};
    for (const f of files) {
      const nameWithoutExt = f.name.replace(/\.[^/.]+$/, "");
      const url = await fileToDataURL(f);
      map[nameWithoutExt.toLowerCase()] = url;
    }
    setBulkPhotos(map);

    // Auto-match to students
    setStudents((prev) =>
      prev.map((s, idx) => {
        const roll = (s.rollNumber || "").toLowerCase();
        const seq = String(idx + 1);
        const matched = map[roll] || map[seq] || map[s.studentName?.toLowerCase()] || null;
        return matched ? { ...s, photo: matched } : s;
      })
    );
    showToast(`📸 Matched ${Object.keys(map).length} photos!`);
    setLoading(false);
  };

  // ── Filtered students ───────────────────────────────────────────────────
  const filtered = students.filter(
    (s) =>
      !search ||
      s.studentName.toLowerCase().includes(search.toLowerCase()) ||
      s.rollNumber.toLowerCase().includes(search.toLowerCase())
  );

  // ─────────────────────────────────────────────────────────────────────────
  const bg = darkMode ? "#0f172a" : "#eef5ff";
  const panel = darkMode ? "#1e293b" : "#fff";
  const text = darkMode ? "#e2e8f0" : "#1a3a5c";
  const border = darkMode ? "#334155" : "#d0e4f7";
  const accent = "#0d6efd";

  const navItems = [
    { id: "form", label: "➕ Single Card", icon: "📝" },
    { id: "bulk", label: "📦 Bulk Upload", icon: "📦" },
    { id: "preview", label: "👁 Preview All", icon: "👁" },
    { id: "manage", label: "🗂 Manage", icon: "🗂" },
  ];

  return (
    <div style={{ minHeight: "100vh", background: bg, fontFamily: "'Segoe UI', sans-serif", color: text }}>
      <style>{`
        * { box-sizing: border-box; margin: 0; padding: 0; }
        @keyframes slideUp { from { transform: translateY(20px); opacity: 0; } to { transform: translateY(0); opacity: 1; } }
        @keyframes spin { to { transform: rotate(360deg); } }
        ::-webkit-scrollbar { width: 6px; } ::-webkit-scrollbar-thumb { background: #93c5fd; border-radius: 6px; }
        input:focus { box-shadow: 0 0 0 3px rgba(13,110,253,.15); }
      `}</style>

      {/* ── Header ── */}
      <header
        style={{
          background: `linear-gradient(135deg,#1a3a5c 0%,#0d6efd 100%)`,
          padding: "0 24px",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          height: 60,
          boxShadow: "0 2px 16px rgba(13,110,253,.3)",
          position: "sticky", top: 0, zIndex: 100,
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <span style={{ fontSize: 26 }}>🏫</span>
          <div>
            <div style={{ color: "#fff", fontWeight: 800, fontSize: 17, letterSpacing: 0.5 }}>
              ScholarCard Pro
            </div>
            <div style={{ color: "rgba(255,255,255,.65)", fontSize: 10 }}>
              School ID Card Generator
            </div>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <span style={{ color: "rgba(255,255,255,.7)", fontSize: 12 }}>
            {students.length} student{students.length !== 1 ? "s" : ""}
          </span>
          <button
            onClick={() => setDarkMode(!darkMode)}
            style={{ background: "rgba(255,255,255,.15)", border: "none", borderRadius: 8, padding: "6px 10px", cursor: "pointer", fontSize: 16, color: "#fff" }}
          >
            {darkMode ? "☀️" : "🌙"}
          </button>
          <button
            onClick={() => exportToPDF(students)}
            style={{ background: "#fff", color: "#0d6efd", border: "none", borderRadius: 8, padding: "7px 14px", cursor: "pointer", fontWeight: 700, fontSize: 12, display: "flex", alignItems: "center", gap: 5 }}
          >
            📄 Export PDF
          </button>
        </div>
      </header>

      {/* ── Nav ── */}
      <div
        style={{
          background: panel,
          borderBottom: `1.5px solid ${border}`,
          display: "flex",
          padding: "0 24px",
          gap: 4,
          overflowX: "auto",
        }}
      >
        {navItems.map((n) => (
          <button
            key={n.id}
            onClick={() => setTab(n.id)}
            style={{
              padding: "12px 18px",
              border: "none",
              background: "none",
              cursor: "pointer",
              fontWeight: tab === n.id ? 700 : 500,
              fontSize: 13,
              color: tab === n.id ? accent : (darkMode ? "#94a3b8" : "#6b9cc7"),
              borderBottom: tab === n.id ? `2.5px solid ${accent}` : "2.5px solid transparent",
              whiteSpace: "nowrap",
              transition: "all .2s",
            }}
          >
            {n.label}
          </button>
        ))}
      </div>

      {/* ── Loading overlay ── */}
      {loading && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.4)", zIndex: 9998, display: "flex", alignItems: "center", justifyContent: "center", flexDirection: "column", gap: 14 }}>
          <div style={{ width: 48, height: 48, border: "4px solid rgba(255,255,255,.3)", borderTopColor: "#fff", borderRadius: "50%", animation: "spin 1s linear infinite" }} />
          <div style={{ color: "#fff", fontSize: 15, fontWeight: 600 }}>Processing…</div>
        </div>
      )}

      {/* ── Content ── */}
      <div style={{ maxWidth: 1200, margin: "0 auto", padding: "24px 16px" }}>

        {/* ═══════════ FORM TAB ═══════════ */}
        {tab === "form" && (
          <div style={{ display: "grid", gridTemplateColumns: "1fr auto", gap: 24, alignItems: "start" }}>
            {/* Left: Form */}
            <div
              style={{
                background: panel,
                borderRadius: 16,
                border: `1.5px solid ${border}`,
                padding: 24,
                boxShadow: "0 4px 24px rgba(13,110,253,.06)",
              }}
            >
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
                <h2 style={{ fontSize: 17, fontWeight: 700, color: text }}>
                  Student Details
                  <span style={{ marginLeft: 10, fontSize: 12, color: "#6b9cc7", fontWeight: 400 }}>
                    #{activeIdx + 1} of {students.length}
                  </span>
                </h2>
                <div style={{ display: "flex", gap: 8 }}>
                  <button
                    onClick={() => {
                      const ns = [...students, emptyStudent()];
                      setStudents(ns);
                      setActiveIdx(ns.length - 1);
                    }}
                    style={{ background: "#eff6ff", color: accent, border: `1.5px solid ${accent}`, borderRadius: 8, padding: "6px 12px", cursor: "pointer", fontSize: 12, fontWeight: 600 }}
                  >
                    + Add Student
                  </button>
                  {students.length > 1 && (
                    <button
                      onClick={() => {
                        const ns = students.filter((_, i) => i !== activeIdx);
                        setStudents(ns);
                        setActiveIdx(Math.max(0, activeIdx - 1));
                      }}
                      style={{ background: "#fff1f2", color: "#e74c3c", border: "1.5px solid #e74c3c", borderRadius: 8, padding: "6px 12px", cursor: "pointer", fontSize: 12, fontWeight: 600 }}
                    >
                      🗑 Remove
                    </button>
                  )}
                </div>
              </div>

              {/* Student selector chips */}
              {students.length > 1 && (
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 18 }}>
                  {students.map((s, i) => (
                    <button
                      key={s.id}
                      onClick={() => setActiveIdx(i)}
                      style={{
                        padding: "4px 10px",
                        borderRadius: 20,
                        border: `1.5px solid ${i === activeIdx ? accent : border}`,
                        background: i === activeIdx ? "#eff6ff" : panel,
                        color: i === activeIdx ? accent : text,
                        fontSize: 11,
                        cursor: "pointer",
                        fontWeight: i === activeIdx ? 700 : 400,
                      }}
                    >
                      {s.studentName || `Student ${i + 1}`}
                    </button>
                  ))}
                </div>
              )}

              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
                <div style={{ gridColumn: "1/-1" }}>
                  <Field label="School / College Name" value={cur.schoolName} onChange={updateField("schoolName")} required />
                </div>
                <Field label="Student Name" value={cur.studentName} onChange={updateField("studentName")} required />
                <Field label="Roll Number" value={cur.rollNumber} onChange={updateField("rollNumber")} />
                <Field label="Father's Name" value={cur.fatherName} onChange={updateField("fatherName")} />
                <Field label="Mother's Name" value={cur.motherName} onChange={updateField("motherName")} />
                <Field label="Date of Birth" value={cur.dob} onChange={updateField("dob")} type="date" />
                <Field label="Class / Course" value={cur.course} onChange={updateField("course")} />
                <Field label="Student Mobile" value={cur.mobile} onChange={updateField("mobile")} />
                <Field label="Aadhaar Number" value={cur.aadhaar} onChange={updateField("aadhaar")} />
                <Field label="School Contact" value={cur.schoolContact} onChange={updateField("schoolContact")} />
                <div style={{ gridColumn: "1/-1" }}>
                  <Field label="Full Address" value={cur.address} onChange={updateField("address")} />
                </div>

                {/* Image uploads */}
                {[
                  { key: "photo", label: "Student Photo", icon: "🖼️" },
                  { key: "principalSig", label: "Principal Signature", icon: "✍️" },
                  { key: "schoolLogo", label: "School Logo", icon: "🏫" },
                ].map(({ key, label, icon }) => (
                  <div key={key} style={{ gridColumn: key === "photo" ? "1/-1" : "auto" }}>
                    <div style={{ fontSize: 12, fontWeight: 600, color: text, marginBottom: 6 }}>{label}</div>
                    {cur[key] ? (
                      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                        <img
                          src={cur[key]}
                          alt={label}
                          style={{ width: key === "photo" ? 64 : 52, height: key === "photo" ? 80 : 40, objectFit: key === "photo" ? "cover" : "contain", borderRadius: 6, border: `1.5px solid ${border}` }}
                        />
                        <button
                          onClick={() => updateField(key)(null)}
                          style={{ background: "#fff1f2", color: "#e74c3c", border: "none", borderRadius: 6, padding: "4px 8px", cursor: "pointer", fontSize: 11 }}
                        >
                          Remove
                        </button>
                      </div>
                    ) : (
                      <DropZone
                        label={`Upload ${label}`}
                        accept="image/*"
                        multiple={false}
                        icon={icon}
                        onFiles={(files) => handleImageField(key, files)}
                      />
                    )}
                  </div>
                ))}
              </div>

              <button
                onClick={() => { showToast("✅ Card saved successfully!"); }}
                style={{
                  marginTop: 20, width: "100%",
                  background: `linear-gradient(135deg,#1a3a5c,${accent})`,
                  color: "#fff", border: "none", borderRadius: 10,
                  padding: "12px", fontWeight: 700, fontSize: 14,
                  cursor: "pointer", letterSpacing: 0.3,
                }}
              >
                💾 Save Card
              </button>
            </div>

            {/* Right: Live Preview */}
            <div style={{ position: "sticky", top: 80 }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: "#6b9cc7", marginBottom: 10, textAlign: "center", letterSpacing: 1, textTransform: "uppercase" }}>
                Live Preview
              </div>
              <div style={{ boxShadow: "0 8px 40px rgba(13,110,253,.18)", borderRadius: 12, overflow: "hidden" }}>
                <IDCard student={cur} />
              </div>
              <button
                onClick={() => exportToPDF([cur])}
                style={{
                  marginTop: 14, width: 340,
                  background: "#10b981", color: "#fff",
                  border: "none", borderRadius: 10, padding: "10px",
                  fontWeight: 700, fontSize: 13, cursor: "pointer",
                }}
              >
                ⬇️ Download This Card
              </button>
            </div>
          </div>
        )}

        {/* ═══════════ BULK TAB ═══════════ */}
        {tab === "bulk" && (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
            {/* CSV Upload */}
            <div style={{ background: panel, borderRadius: 16, border: `1.5px solid ${border}`, padding: 24, boxShadow: "0 4px 24px rgba(13,110,253,.06)" }}>
              <h3 style={{ fontSize: 16, fontWeight: 700, marginBottom: 6, color: text }}>📊 Upload Student Data</h3>
              <p style={{ fontSize: 12, color: "#6b9cc7", marginBottom: 16 }}>
                Upload Excel (.xlsx) or CSV file. Columns: studentName, rollNumber, fatherName, motherName, dob, course, mobile, aadhaar, address, schoolName, schoolContact
              </p>
              <DropZone
                label="Drop Excel / CSV file here"
                accept=".xlsx,.xls,.csv"
                multiple={false}
                icon="📊"
                onFiles={handleCSV}
              />
              <button
                onClick={() => {
                  // Generate sample CSV download
                  const headers = ["studentName", "rollNumber", "fatherName", "motherName", "dob", "course", "mobile", "aadhaar", "address", "schoolName", "schoolContact"];
                  const sample = [
                    ["Ravi Kumar", "101", "Suresh Kumar", "Sunita Devi", "2005-06-15", "Class 10 A", "9876543210", "123456789012", "123 MG Road, Delhi", "Delhi Public School", "011-23456789"],
                    ["Priya Sharma", "102", "Rajesh Sharma", "Meena Sharma", "2006-03-22", "Class 9 B", "9876543211", "234567890123", "456 Lake View, Mumbai", "Delhi Public School", "011-23456789"],
                  ];
                  const ws = XLSX.utils.aoa_to_sheet([headers, ...sample]);
                  const wb = XLSX.utils.book_new();
                  XLSX.utils.book_append_sheet(wb, ws, "Students");
                  XLSX.writeFile(wb, "sample_students.xlsx");
                  showToast("📥 Sample file downloaded!");
                }}
                style={{ marginTop: 12, width: "100%", background: "#eff6ff", color: accent, border: `1.5px solid ${accent}`, borderRadius: 8, padding: "9px", cursor: "pointer", fontWeight: 600, fontSize: 12 }}
              >
                ⬇️ Download Sample Template
              </button>
            </div>

            {/* Photo Upload */}
            <div style={{ background: panel, borderRadius: 16, border: `1.5px solid ${border}`, padding: 24, boxShadow: "0 4px 24px rgba(13,110,253,.06)" }}>
              <h3 style={{ fontSize: 16, fontWeight: 700, marginBottom: 6, color: text }}>📸 Bulk Photo Upload</h3>
              <p style={{ fontSize: 12, color: "#6b9cc7", marginBottom: 16 }}>
                Name photos by Roll Number (e.g., 101.jpg, 102.jpg) or sequential number (1.jpg, 2.jpg). Photos auto-match to student records.
              </p>
              <DropZone
                label="Drop student photos here"
                accept="image/*"
                multiple={true}
                icon="📷"
                onFiles={handleBulkPhotos}
              />
              {Object.keys(bulkPhotos).length > 0 && (
                <div style={{ marginTop: 12, padding: "10px 12px", background: "#f0fdf4", borderRadius: 8, border: "1px solid #86efac", fontSize: 12, color: "#166534" }}>
                  ✅ {Object.keys(bulkPhotos).length} photos loaded and matched
                </div>
              )}
            </div>

            {/* Status */}
            <div style={{ gridColumn: "1/-1", background: panel, borderRadius: 16, border: `1.5px solid ${border}`, padding: 24, boxShadow: "0 4px 24px rgba(13,110,253,.06)" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
                <h3 style={{ fontSize: 16, fontWeight: 700, color: text }}>📋 Imported Students ({students.length})</h3>
                <button
                  onClick={() => exportToPDF(students)}
                  style={{ background: `linear-gradient(135deg,#1a3a5c,${accent})`, color: "#fff", border: "none", borderRadius: 8, padding: "9px 18px", cursor: "pointer", fontWeight: 700, fontSize: 13 }}
                >
                  📄 Export All as PDF
                </button>
              </div>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead>
                    <tr style={{ background: "#f0f7ff" }}>
                      {["#", "Photo", "Name", "Roll No.", "Course", "Father", "Mobile", "Actions"].map((h) => (
                        <th key={h} style={{ padding: "8px 10px", textAlign: "left", color: "#1a3a5c", fontWeight: 700, borderBottom: `1.5px solid ${border}` }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {students.map((s, i) => (
                      <tr key={s.id} style={{ borderBottom: `1px solid ${border}` }}>
                        <td style={{ padding: "8px 10px", color: "#6b9cc7" }}>{i + 1}</td>
                        <td style={{ padding: "8px 10px" }}>
                          {s.photo ? (
                            <img src={s.photo} alt="p" style={{ width: 32, height: 40, objectFit: "cover", borderRadius: 4 }} />
                          ) : (
                            <span style={{ fontSize: 18 }}>👤</span>
                          )}
                        </td>
                        <td style={{ padding: "8px 10px", fontWeight: 600, color: text }}>{s.studentName || "—"}</td>
                        <td style={{ padding: "8px 10px", color: "#6b9cc7" }}>{s.rollNumber || "—"}</td>
                        <td style={{ padding: "8px 10px", color: text }}>{s.course || "—"}</td>
                        <td style={{ padding: "8px 10px", color: text }}>{s.fatherName || "—"}</td>
                        <td style={{ padding: "8px 10px", color: "#6b9cc7" }}>{s.mobile || "—"}</td>
                        <td style={{ padding: "8px 10px" }}>
                          <button
                            onClick={() => { setActiveIdx(i); setTab("form"); }}
                            style={{ background: "#eff6ff", color: accent, border: "none", borderRadius: 6, padding: "4px 8px", cursor: "pointer", fontSize: 11, fontWeight: 600 }}
                          >
                            Edit
                          </button>
                          <button
                            onClick={() => exportToPDF([s])}
                            style={{ marginLeft: 4, background: "#f0fdf4", color: "#10b981", border: "none", borderRadius: 6, padding: "4px 8px", cursor: "pointer", fontSize: 11, fontWeight: 600 }}
                          >
                            PDF
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* ═══════════ PREVIEW TAB ═══════════ */}
        {tab === "preview" && (
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
              <div>
                <h2 style={{ fontSize: 18, fontWeight: 700, color: text }}>All ID Cards Preview</h2>
                <p style={{ fontSize: 12, color: "#6b9cc7", marginTop: 2 }}>{students.length} card{students.length !== 1 ? "s" : ""} generated</p>
              </div>
              <div style={{ display: "flex", gap: 10 }}>
                <input
                  placeholder="🔍 Search student or roll no..."
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  style={{ padding: "8px 14px", borderRadius: 8, border: `1.5px solid ${border}`, fontSize: 12, background: panel, color: text, outline: "none", width: 220 }}
                />
                <button
                  onClick={() => exportToPDF(filtered)}
                  style={{ background: `linear-gradient(135deg,#1a3a5c,${accent})`, color: "#fff", border: "none", borderRadius: 8, padding: "9px 18px", cursor: "pointer", fontWeight: 700, fontSize: 13 }}
                >
                  📄 Export {filtered.length} Cards
                </button>
              </div>
            </div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 24, justifyContent: "flex-start" }}>
              {filtered.map((s, i) => (
                <div key={s.id} style={{ position: "relative" }}>
                  <div style={{ boxShadow: "0 4px 24px rgba(13,110,253,.12)", borderRadius: 12, overflow: "hidden" }}>
                    <IDCard student={s} scale={0.85} />
                  </div>
                  <div style={{ marginTop: 8, display: "flex", gap: 6, justifyContent: "center" }}>
                    <button
                      onClick={() => { const idx = students.indexOf(s); setActiveIdx(idx); setTab("form"); }}
                      style={{ background: "#eff6ff", color: accent, border: "none", borderRadius: 6, padding: "5px 10px", cursor: "pointer", fontSize: 11, fontWeight: 600 }}
                    >
                      ✏️ Edit
                    </button>
                    <button
                      onClick={() => exportToPDF([s])}
                      style={{ background: "#f0fdf4", color: "#10b981", border: "none", borderRadius: 6, padding: "5px 10px", cursor: "pointer", fontSize: 11, fontWeight: 600 }}
                    >
                      📥 PDF
                    </button>
                  </div>
                </div>
              ))}
            </div>
            {filtered.length === 0 && (
              <div style={{ textAlign: "center", padding: "60px 0", color: "#6b9cc7" }}>
                <div style={{ fontSize: 48 }}>🔍</div>
                <div style={{ marginTop: 12, fontWeight: 600 }}>No students found</div>
              </div>
            )}
          </div>
        )}

        {/* ═══════════ MANAGE TAB ═══════════ */}
        {tab === "manage" && (
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(280px,1fr))", gap: 20 }}>
            {/* Stats */}
            {[
              { label: "Total Students", value: students.length, icon: "👥", color: "#0d6efd" },
              { label: "With Photos", value: students.filter((s) => s.photo).length, icon: "📸", color: "#10b981" },
              { label: "With Logos", value: students.filter((s) => s.schoolLogo).length, icon: "🏫", color: "#f59e0b" },
              { label: "Complete Cards", value: students.filter((s) => s.studentName && s.photo && s.course).length, icon: "✅", color: "#8b5cf6" },
            ].map((stat) => (
              <div
                key={stat.label}
                style={{
                  background: panel, borderRadius: 14,
                  border: `1.5px solid ${border}`,
                  padding: 20,
                  boxShadow: "0 4px 16px rgba(13,110,253,.05)",
                  display: "flex", alignItems: "center", gap: 16,
                }}
              >
                <div style={{ width: 52, height: 52, borderRadius: 12, background: `${stat.color}18`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 24 }}>
                  {stat.icon}
                </div>
                <div>
                  <div style={{ fontSize: 28, fontWeight: 800, color: stat.color }}>{stat.value}</div>
                  <div style={{ fontSize: 12, color: "#6b9cc7", marginTop: 2 }}>{stat.label}</div>
                </div>
              </div>
            ))}

            {/* Actions */}
            <div style={{ gridColumn: "1/-1", background: panel, borderRadius: 14, border: `1.5px solid ${border}`, padding: 24 }}>
              <h3 style={{ fontWeight: 700, fontSize: 16, color: text, marginBottom: 16 }}>⚡ Quick Actions</h3>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
                {[
                  { label: "➕ New Student", action: () => { const ns = [...students, emptyStudent()]; setStudents(ns); setActiveIdx(ns.length - 1); setTab("form"); }, color: accent },
                  { label: "📄 Export All PDF", action: () => exportToPDF(students), color: "#10b981" },
                  { label: "🧹 Clear All", action: () => { if (confirm("Clear all students?")) { setStudents([emptyStudent()]); setActiveIdx(0); setTab("form"); } }, color: "#ef4444" },
                  { label: "📊 Download Template", action: () => { const headers = ["studentName","rollNumber","fatherName","motherName","dob","course","mobile","aadhaar","address","schoolName","schoolContact"]; const ws = XLSX.utils.aoa_to_sheet([headers]); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Students"); XLSX.writeFile(wb, "template.xlsx"); showToast("📥 Template downloaded!"); }, color: "#f59e0b" },
                ].map((a) => (
                  <button
                    key={a.label}
                    onClick={a.action}
                    style={{
                      background: `${a.color}15`, color: a.color,
                      border: `1.5px solid ${a.color}`,
                      borderRadius: 8, padding: "9px 16px",
                      cursor: "pointer", fontWeight: 700, fontSize: 13,
                    }}
                  >
                    {a.label}
                  </button>
                ))}
              </div>
            </div>

            {/* Usage guide */}
            <div style={{ gridColumn: "1/-1", background: panel, borderRadius: 14, border: `1.5px solid ${border}`, padding: 24 }}>
              <h3 style={{ fontWeight: 700, fontSize: 16, color: text, marginBottom: 14 }}>📖 How to Use</h3>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(220px,1fr))", gap: 14 }}>
                {[
                  { step: "1", title: "Single Card", desc: "Go to 'Single Card' tab, fill in student details, upload photo and logo, preview live, then download." },
                  { step: "2", title: "Bulk Upload", desc: "Download template, fill student data in Excel, upload CSV/XLSX. Then upload photos named by roll number (101.jpg)." },
                  { step: "3", title: "Auto-Matching", desc: "Photos are auto-matched by roll number or sequential number. 101.jpg → Roll 101, 1.jpg → Student 1." },
                  { step: "4", title: "Export PDF", desc: "Click 'Export PDF' from any tab. A print window opens — click Print and select 'Save as PDF'." },
                ].map((s) => (
                  <div key={s.step} style={{ background: "#f8fbff", borderRadius: 10, padding: 16, border: `1px solid ${border}` }}>
                    <div style={{ width: 28, height: 28, borderRadius: "50%", background: accent, color: "#fff", fontWeight: 800, fontSize: 13, display: "flex", alignItems: "center", justifyContent: "center", marginBottom: 10 }}>
                      {s.step}
                    </div>
                    <div style={{ fontWeight: 700, fontSize: 13, color: "#1a3a5c", marginBottom: 6 }}>{s.title}</div>
                    <div style={{ fontSize: 12, color: "#6b9cc7", lineHeight: 1.6 }}>{s.desc}</div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}
      </div>

      {toast && <Toast msg={toast.msg} type={toast.type} onClose={() => setToast(null)} />}
    </div>
  );
}
