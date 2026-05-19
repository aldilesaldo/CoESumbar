/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { 
  LayoutDashboard, 
  PlusCircle, 
  ListOrdered, 
  Calendar as CalendarIcon, 
  LogOut, 
  RefreshCcw, 
  FileSpreadsheet, 
  Download, 
  Upload, 
  ChevronLeft, 
  ChevronRight, 
  Search, 
  MapPin, 
  Tag, 
  User, 
  Phone, 
  Info,
  CheckCircle2,
  Clock,
  Trash2,
  Edit2,
  ExternalLink,
  PieChart,
  BarChart3,
  HelpCircle
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

// ── KONSTANTA ─────────────────────────────────────────────────
const KABUPATEN_KOTA = [
  "Kota Padang","Kota Bukittinggi","Kota Payakumbuh","Kota Padang Panjang",
  "Kota Solok","Kota Pariaman","Kota Sawahlunto",
  "Kabupaten Agam","Kabupaten Tanah Datar","Kabupaten Padang Pariaman",
  "Kabupaten Pesisir Selatan","Kabupaten Solok","Kabupaten Solok Selatan",
  "Kabupaten Sijunjung","Kabupaten Dharmasraya","Kabupaten Lima Puluh Kota",
  "Kabupaten Pasaman","Kabupaten Pasaman Barat","Kabupaten Kepulauan Mentawai",
];

const KATEGORI = [
  "Festival Budaya","Olahraga & Adventure","Kuliner","Seni & Pertunjukan",
  "Religi & Tradisi","MICE","Alam & Ekowisata","Pameran & Expo",
];

const BULAN = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];

const HOLIDAYS_2026: Record<string, string> = {
  "2026-01-01": "Tahun Baru 2026 Masehi",
  "2026-01-16": "Isra Mi'raj Nabi Muhammad SAW",
  "2026-02-17": "Tahun Baru Imlek 2577 Kongzili",
  "2026-03-19": "Hari Suci Nyepi (Tahun Baru Saka 1948)",
  "2026-03-20": "Hari Raya Idul Fitri 1447 H",
  "2026-03-21": "Hari Raya Idul Fitri 1447 H",
  "2026-03-23": "Cuti Bersama Idul Fitri",
  "2026-03-24": "Cuti Bersama Idul Fitri",
  "2026-04-03": "Wafat Yesus Kristus",
  "2026-05-01": "Hari Buruh Internasional",
  "2026-05-11": "Kenaikan Yesus Kristus",
  "2026-05-22": "Hari Raya Waisak 2570 BE",
  "2026-05-27": "Hari Raya Idul Adha 1447 H",
  "2026-06-01": "Hari Lahir Pancasila",
  "2026-06-16": "Tahun Baru Islam 1448 H",
  "2026-08-17": "Hari Kemerdekaan RI",
  "2026-08-25": "Maulid Nabi Muhammad SAW",
  "2026-12-25": "Hari Raya Natal",
  "2026-12-26": "Cuti Bersama Natal",
};

const STATUS_BADGE: Record<string, { bg: string; color: string; border: string }> = {
  "Draft":    { bg:"rgba(71, 85, 105, 0.2)", color:"#94A3B8", border:"rgba(148, 163, 184, 0.2)" },
  "Diajukan": { bg:"rgba(16, 185, 129, 0.15)", color:"#10B981", border:"rgba(16, 185, 129, 0.3)" },
};

const ACCOUNTS = [
  { username:"admin.provinsi", password:"sumbarrancak", role:"provinsi", kabupatenKota:"Provinsi Sumatera Barat", nama:"Provinsi Sumatera Barat" },
  ...KABUPATEN_KOTA.map((k)=>({
    username: "admin." + k.toLowerCase().replace(/\s+/g,".").replace(/\//g,""),
    password: "sumbarrancak",
    role: "kabkota",
    kabupatenKota: k,
    nama: `Admin Dispar ${k}`,
  })),
];

interface EventData {
  id: number;
  namaEvent: string;
  kabupatenKota: string;
  kategori: string;
  tanggalMulai: string;
  tanggalSelesai?: string;
  lokasi?: string;
  deskripsi?: string;
  targetWisatawan?: string;
  kontakNama?: string;
  kontakHP?: string;
  anggaran?: string;
  status: string;
  createdAt?: string;
}

const initialForm: Omit<EventData, 'id'> = {
  namaEvent:"", kabupatenKota:"", kategori:"", tanggalMulai:"",
  tanggalSelesai:"", lokasi:"", deskripsi:"", targetWisatawan:"",
  kontakNama:"", kontakHP:"", anggaran:"", status:"Draft",
};

const STORAGE_KEY  = "coe_sumbar_v2_events";
const SESSION_KEY  = "coe_sumbar_session";

// ── API HELPERS (Proxy through Server) ───────────────────────────────
async function sheetSaveEvent(event: EventData): Promise<{success: boolean, detail?: string}> {
  try {
    const res = await fetch("/api/events", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action: "save", event }),
    });
    if (res.ok) return { success: true };
    const data = await res.json();
    return { success: false, detail: data.detail || data.error || `Status ${res.status}` };
  } catch(e: any) { 
    console.warn("Sheet save failed:", e); 
    return { success: false, detail: e.message };
  }
}

async function sheetDeleteEvent(id: number) {
  try {
    const res = await fetch("/api/events", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action: "delete", id }),
    });
    return res.ok;
  } catch(e) { 
    console.warn("Sheet delete failed:", e); 
    return false;
  }
}

async function sheetGetAll(): Promise<EventData[] | null> {
  try {
    const res = await fetch("/api/events");
    const data = await res.json();
    if (Array.isArray(data)) return data;
    if (data && data.events && Array.isArray(data.events)) return data.events;
    return null;
  } catch(e) { console.warn("Sheet getAll failed:", e); return null; }
}

function getSeedData(): EventData[] {
  return [
    { id:1, namaEvent:"Pesona Budaya Hoyak Tabuik Piaman", kabupatenKota:"Kota Pariaman",
      kategori:"Festival Budaya", tanggalMulai:"2025-07-05", tanggalSelesai:"2025-07-14",
      lokasi:"Pantai Gandoriah, Pariaman", deskripsi:"Festival tradisi tabuik yang digelar setiap Muharram.",
      targetWisatawan:"5000", kontakNama:"Dedi Rahmat", kontakHP:"081234567890",
      anggaran:"500000000", status:"Diajukan", createdAt:"2025-01-10" },
    { id:2, namaEvent:"Tour de Singkarak", kabupatenKota:"Kabupaten Solok",
      kategori:"Olahraga & Adventure", tanggalMulai:"2025-10-15", tanggalSelesai:"2025-10-22",
      lokasi:"Danau Singkarak & sekitarnya", deskripsi:"Balap sepeda internasional mengelilingi alam Sumatera Barat.",
      targetWisatawan:"20000", kontakNama:"Roni Amir", kontakHP:"082345678901",
      anggaran:"2000000000", status:"Diajukan", createdAt:"2025-02-01" },
  ];
}

function exportToExcel(events: EventData[]) {
  const rows = events.map((e,i)=>({
    "No":i+1,"Nama Event":e.namaEvent,"Provinsi / Kabupaten / Kota":e.kabupatenKota,
    "Kategori":e.kategori,"Tanggal Mulai":e.tanggalMulai,"Tanggal Selesai":e.tanggalSelesai||"-",
    "Lokasi / Venue":e.lokasi||"-","Deskripsi":e.deskripsi||"-",
    "Target Wisatawan":e.targetWisatawan?parseInt(e.targetWisatawan):0,
    "Estimasi Anggaran (Rp)":e.anggaran?parseInt(e.anggaran):0,
    "Narahubung":e.kontakNama||"-","No. HP":e.kontakHP||"-",
    "Status":e.status,"Tanggal Input":e.createdAt||"-",
  }));
  const ws=XLSX.utils.json_to_sheet(rows);
  ws["!cols"]=[{wch:5},{wch:35},{wch:25},{wch:20},{wch:14},{wch:14},{wch:30},{wch:40},{wch:18},{wch:22},{wch:20},{wch:16},{wch:28},{wch:12},{wch:14}];
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"Calendar of Events");
  XLSX.writeFile(wb,`CoE_SumBar_${new Date().toISOString().split("T")[0]}.xlsx`);
}

// ══════════════════════════════════════════════════════════════
//  ROOT
// ══════════════════════════════════════════════════════════════
export default function App() {
  const [user, setUser] = useState(() => {
    try { const r=sessionStorage.getItem(SESSION_KEY); return r?JSON.parse(r):null; }
    catch { return null; }
  });

  const login  = (u: any)=>{ sessionStorage.setItem(SESSION_KEY,JSON.stringify(u)); setUser(u); };
  const logout = ()=>{ sessionStorage.removeItem(SESSION_KEY); setUser(null); };

  if (!user) return <LoginScreen onLogin={login} />;
  return <MainApp user={user} onLogout={logout} />;
}

// ══════════════════════════════════════════════════════════════
//  LOGIN SCREEN
// ══════════════════════════════════════════════════════════════
function LoginScreen({ onLogin }: { onLogin: (u: any) => void }) {
  const [uname, setUname] = useState("");
  const [pass,  setPass]  = useState("");
  const [err,   setErr]   = useState("");

  const doLogin = (override?: any) => {
    const targetU = override?.username || uname;
    const targetP = override?.password || pass;
    const acc = ACCOUNTS.find(a=>a.username===targetU.trim()&&a.password===targetP);
    if (!acc) { setErr("Username atau password salah."); return; }
    onLogin(acc);
  };

  return (
    <div className="min-h-screen bg-[#f8efe3] flex items-center justify-center p-4 relative overflow-hidden" 
         style={{ backgroundImage: `radial-gradient(circle at 10% 20%, rgba(184, 134, 11, 0.05) 0%, transparent 40%), radial-gradient(circle at 90% 80%, rgba(5, 150, 105, 0.05) 0%, transparent 40%)` }}>
      
      <motion.div 
        initial={{ opacity: 0, y: 20 }}
        animate={{ opacity: 1, y: 0 }}
        className="w-full max-w-[420px] bg-white border border-slate-200 rounded-3xl p-8 shadow-[0_32px_80px_rgba(0,0,0,0.1)] relative z-10 overflow-hidden"
      >
        <div className="absolute top-0 left-0 w-full h-1 bg-gradient-to-r from-transparent via-[#D4AF37] to-transparent" />
        
        <div className="text-center mb-6">
          <motion.div 
            initial={{ scale: 0.8 }}
            animate={{ scale: 1 }}
            className="w-16 h-16 bg-amber-50 rounded-2xl flex items-center justify-center mx-auto mb-4 border border-amber-100 shadow-sm"
          >
            <MapPin className="text-[#B8860B] w-8 h-8" />
          </motion.div>
          <div className="text-[10px] tracking-[4px] text-amber-700 uppercase mb-1 font-bold">Dinas Pariwisata</div>
          <h1 className="text-xl font-black text-slate-800 tracking-tight">Sumatera Barat</h1>
          <div className="text-[11px] text-slate-400 mt-1 font-medium">Sistem Pendataan Calendar of Events</div>
        </div>

        {/* Login Form for Province */}
        <div className="space-y-3 mb-6 bg-slate-50/50 p-5 rounded-2xl border border-slate-100">
          <div className="text-[9px] text-amber-600 mb-2 uppercase tracking-widest font-black flex items-center gap-2">
            <span className="w-1.5 h-1.5 rounded-full bg-amber-500"></span>
            Provinsi Sumatera Barat
          </div>
          <input 
            type="text" 
            placeholder="Username Provinsi" 
            className="w-full bg-white border border-slate-200 rounded-xl px-4 py-2.5 text-xs outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-50 shadow-sm transition-all"
            value={uname}
            onChange={(e) => setUname(e.target.value)}
          />
          <input 
            type="password" 
            placeholder="Password" 
            className="w-full bg-white border border-slate-200 rounded-xl px-4 py-2.5 text-xs outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-50 shadow-sm transition-all"
            value={pass}
            onChange={(e) => setPass(e.target.value)}
          />
          <button 
            onClick={() => doLogin()}
            className="w-full bg-slate-800 text-white rounded-xl py-2.5 text-xs font-bold hover:bg-slate-900 transition-all shadow-md active:scale-[0.98]"
          >
            Masuk Akses Provinsi
          </button>
        </div>

        {err && <div className="mb-6 text-[10px] text-red-600 bg-red-50 p-3 rounded-xl border border-red-100 font-semibold text-center">{err}</div>}

        <div className="space-y-4">
          <div className="pt-4 border-t border-slate-100">
            <div className="text-[10px] text-slate-400 font-bold uppercase tracking-[2px] mb-4 text-center opacity-70">Akses Cepat Kabupaten / Kota</div>
            
            <div className="space-y-4 max-h-[250px] overflow-y-auto px-1 custom-scroll bg-slate-50/30 rounded-2xl p-4 border border-slate-50">
              <div className="grid grid-cols-2 gap-2">
                {ACCOUNTS.filter(a => a.role === "kabkota" && a.kabupatenKota !== "Provinsi Sumatera Barat").map(a => (
                  <QuickLoginBtn 
                    key={a.username} 
                    label={a.kabupatenKota} 
                    onClick={() => doLogin({ username: a.username, password: a.password })} 
                  />
                ))}
              </div>
            </div>
          </div>
        </div>
      </motion.div>
      <div className="absolute bottom-6 left-0 w-full text-center">
        <span className="text-[10px] text-slate-300 font-bold uppercase tracking-[4px]">Design by Minangkaos</span>
      </div>
    </div>
  );
}

function QuickLoginBtn({ label, onClick, isProv }: { label: string; onClick: () => void; isProv?: boolean }) {
  return (
    <button 
      onClick={onClick}
      className={`w-full text-left p-3 rounded-xl border transition-all text-[11px] font-bold flex items-center gap-2 group
        ${isProv 
          ? "bg-amber-50 border-amber-200 text-amber-700 hover:bg-amber-100 hover:shadow-sm" 
          : "bg-slate-50 border-slate-100 text-slate-500 hover:border-emerald-200 hover:text-emerald-700 hover:bg-emerald-50 hover:shadow-sm"
        }`}
    >
      <div className={`w-1.5 h-1.5 rounded-full ${isProv ? "bg-amber-500 shadow-[0_0_8px_rgba(184,134,11,0.5)]" : "bg-emerald-400 opacity-40 group-hover:opacity-100"}`} />
      <span className="truncate">{label}</span>
    </button>
  );
}

// ══════════════════════════════════════════════════════════════
//  MAIN APP
// ══════════════════════════════════════════════════════════════
function MainApp({ user, onLogout }: { user: any; onLogout: () => void }) {
  const isProvinsi = user.role === "provinsi";

  const [events, setEvents]               = useState<EventData[]>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      return saved ? JSON.parse(saved) : [];
    } catch { return []; }
  });
  const [view,   setView]                 = useState("dashboard");
  const [form,   setForm]                 = useState<EventData | Omit<EventData, 'id'>>(initialForm);
  const [editId, setEditId]               = useState<number | null>(null);
  const [selEvent, setSelEvent]           = useState<EventData | null>(null);
  const [filterKab, setFilterKab]         = useState("");
  const [filterKat, setFilterKat]         = useState("");
  const [searchQ, setSearchQ]             = useState("");
  const [calYear,  setCalYear]            = useState(2026);
  const [calMonth, setCalMonth]           = useState(new Date().getMonth());
  const [toast,   setToast]               = useState<{ msg: string; type: string }|null>(null);
  const [delConfirm, setDelConfirm]       = useState<number | null>(null);
  const [ajukanConfirm, setAjukanConfirm] = useState<number | null>(null);
  const [sheetLoading, setSheetLoading]   = useState(true);
  const [lastSync, setLastSync]           = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const exportToJSON = () => {
    const blob = new Blob([JSON.stringify(events, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Backup_CoE_SumBar_${new Date().toISOString().split("T")[0]}.json`;
    a.click();
    URL.revokeObjectURL(url);
    showToast("Data dieksport ke JSON. ✓");
  };

  const handleImportJSON = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (event) => {
      try {
        const jsonData = JSON.parse(event.target?.result as string);
        if (!Array.isArray(jsonData)) throw new Error("Format file tidak valid.");
        if (window.confirm(`${jsonData.length} event ditemukan. Lanjutkan restorasi?`)) {
          setEvents(jsonData);
          for (const ev of jsonData) {
            await sheetSaveEvent(ev);
          }
          showToast("Restorasi data berhasil! ✓");
        }
      } catch (err) { alert("Gagal mengimpor: format file tidak sesuai."); }
      e.target.value = "";
    };
    reader.readAsText(file);
  };

  const fetchUpdatedData = async (silent = false) => {
    if (!silent) setSheetLoading(true);
    const sheetEvents = await sheetGetAll();
    if (sheetEvents && Array.isArray(sheetEvents)) {
      const cleaned = sheetEvents.filter(e => e && e.namaEvent && e.namaEvent.toString().trim() !== "");
      
      // Merge logic: keep local events that might not be in the sheet yet
      // This solves the "disappearing" issue during sync latency
      setEvents(prev => {
        const newEventsMap = new Map(cleaned.map(e => [String(e.id), e]));
        const merged = [...cleaned];
        
        // Add local events that aren't in the fetch result but were created recently (within last 30s)
        const thirtySecondsAgo = Date.now() - 30000;
        prev.forEach(p => {
          if (!newEventsMap.has(String(p.id)) && p.id > thirtySecondsAgo) {
            merged.push(p);
          }
        });
        
        return merged;
      });

      localStorage.setItem(STORAGE_KEY, JSON.stringify(cleaned));
      setLastSync(new Date().toLocaleTimeString());
    }
    if (!silent) setSheetLoading(false);
  };

  useEffect(()=>{
    fetchUpdatedData();
    const interval = setInterval(() => fetchUpdatedData(true), 10000);
    return () => clearInterval(interval);
  },[]);

  const showToast=(msg: string,type="success")=>{ 
    setToast({msg,type}); 
    setTimeout(()=>setToast(null),4000); 
  };

  const filtered = useMemo(() => {
    return (events || []).filter(e => {
      if (!e) return false;
      const q = (searchQ || "").toLowerCase();
      const matchSearch = (e.namaEvent || "").toLowerCase().includes(q) || (e.lokasi || "").toLowerCase().includes(q);
      const matchKab = !filterKab || e.kabupatenKota === filterKab;
      const matchKat = !filterKat || e.kategori === filterKat;
      return matchSearch && matchKab && matchKat;
    });
  }, [events, searchQ, filterKab, filterKat]);

  const stats = useMemo(() => {
    const totalTarget = events.reduce((s, e) => s + parseInt(e.targetWisatawan || "0"), 0);
    const kabCounts = KABUPATEN_KOTA.map(k => ({ 
      name: k, 
      count: events.filter(e => e.kabupatenKota === k).length 
    })).filter(k => k.count > 0).sort((a,b) => b.count - a.count);

    return {
      total: events.length,
      diajukan: events.filter(e => e.status === "Diajukan").length,
      target: totalTarget,
      kabCounts
    };
  }, [events]);

  const handleSubmit = async () => {
    const finalForm = !isProvinsi ? { ...form, kabupatenKota: user.kabupatenKota } : { ...form };

    if (!finalForm.namaEvent || !finalForm.kabupatenKota || !finalForm.tanggalMulai) {
      showToast("Lengkapi data wajib (Nama, Wilayah, Tanggal)!", "error"); return;
    }

    setSheetLoading(true);
    if (editId) {
      const updated = {...finalForm, id:editId} as EventData;
      const res = await sheetSaveEvent(updated);
      if (res.success) {
        setEvents(ev=>ev.map(e=>e.id===editId?updated:e));
        showToast("Event diperbarui. ✓");
        setView("list"); setForm(initialForm); setEditId(null);
      } else {
        showToast(`Gagal menyimpan: ${res.detail || "Kesalahan Sheet"}`, "error");
      }
    } else {
      const newEvent = {...finalForm, id:Date.now(), createdAt:new Date().toISOString()} as EventData;
      const res = await sheetSaveEvent(newEvent);
      if (res.success) {
        setEvents(ev=>[...ev, newEvent]);
        showToast("Event ditambahkan. ✓");
        setView("list"); setForm(initialForm); setEditId(null);
      } else {
        showToast(`Gagal menyimpan: ${res.detail || "Kesalahan Sheet"}`, "error");
      }
    }
    setSheetLoading(false);
  };

  const handleEdit = (e: EventData) => {
    setForm({...e}); setEditId(e.id); setView("form");
  };

  const handleDelete = async (id: any) => {
    if (!id) return;
    setSheetLoading(true);
    const success = await sheetDeleteEvent(Number(id));
    if (success) {
      setEvents(ev=>ev.filter(e=>String(e.id)!==String(id)));
      setDelConfirm(null); 
      showToast("Event berhasil dihapus. ✓");
      if (view==="detail") setView("list");
    } else {
      showToast("Gagal menghapus event dari server.", "error");
    }
    setSheetLoading(false);
  };

  const handleAjukan = async (id: any) => {
    const event = events.find(e => String(e.id) === String(id));
    if (!event) return;
    const updated = {...event, status: "Diajukan"};
    setEvents(ev=>ev.map(e=>String(e.id)===String(id)?updated:e));
    await sheetSaveEvent(updated);
    if (selEvent && String(selEvent.id) === String(id)) setSelEvent(updated);
    setAjukanConfirm(null);
    showToast("Event berhasil diajukan ke Provinsi! ✓");
  };

  const canEdit = (e: EventData) => isProvinsi || (e.kabupatenKota === user.kabupatenKota && e.status === "Draft");

  const getDIM=(y: number,m: number)=>new Date(y,m+1,0).getDate();
  const getFD =(y: number,m: number)=>new Date(y,m,1).getDay();
  const eInDay=(day: number)=>events.filter(e=>{
    if(!e.tanggalMulai) return false;
    const s=new Date(e.tanggalMulai), en=e.tanggalSelesai?new Date(e.tanggalSelesai):s;
    const d=new Date(calYear,calMonth,day); return d>=s&&d<=en;
  });

  return (
    <div className="min-h-screen bg-[#f8efe3] text-slate-900 flex font-sans selection:bg-[#D4AF37] selection:text-white">
      
      {/* ── SIDEBAR ── */}
      <aside className="w-64 bg-white border-r border-slate-200 flex flex-col fixed inset-y-0 z-50 shadow-sm">
        <div className="p-8 border-b border-slate-100">
          <div className="flex items-center gap-3 mb-6">
            <div className="w-10 h-10 bg-[#D4AF37]/10 rounded-xl flex items-center justify-center border border-[#D4AF37]/20">
              <MapPin className="text-[#D4AF37] w-5 h-5" />
            </div>
            <div>
              <div className="text-[8px] tracking-[4px] text-[#B8860B] uppercase font-bold text-opacity-80">Sumatera Barat</div>
              <div className="text-sm font-bold tracking-tight text-slate-800">CoE Admin v2</div>
            </div>
          </div>

          <div className={`p-4 rounded-2xl border ${isProvinsi ? "bg-amber-50 border-amber-200" : "bg-emerald-50 border-emerald-200"}`}>
            <div className={`text-[9px] uppercase font-bold mb-1 tracking-widest ${isProvinsi ? "text-amber-700" : "text-emerald-700"}`}>
              {isProvinsi ? "Dinas Provinsi" : "Admin Wilayah"}
            </div>
            <div className="text-xs font-bold text-slate-700 truncate">{user.nama}</div>
            {!isProvinsi && <div className="text-[10px] text-slate-400 mt-1 truncate">{user.kabupatenKota}</div>}
          </div>
        </div>

        <nav className="flex-1 p-4 space-y-1">
          <MenuBtn key="dash" active={view==="dashboard"} icon={LayoutDashboard} label="Dashboard" onClick={()=>setView("dashboard")} />
          <MenuBtn key="form" active={view==="form"} icon={PlusCircle} label="Tambah Event" onClick={()=>{setForm(initialForm); setEditId(null); setView("form");}} />
          <MenuBtn key="list" active={view==="list"} icon={ListOrdered} label="Daftar Event" onClick={()=>setView("list")} />
          <MenuBtn key="cal" active={view==="calendar"} icon={CalendarIcon} label="Kalender" onClick={()=>setView("calendar")} />
        </nav>

        <div className="p-4 border-t border-slate-100 space-y-2">
          {isProvinsi && (
            <div className="p-3 bg-slate-50 rounded-2xl border border-slate-200 space-y-3">
              <div className="text-[9px] uppercase font-bold text-slate-500 tracking-widest opacity-60">Pusat Data</div>
              
              <button onClick={()=>exportToExcel(events)} className="flex items-center gap-2.5 w-full p-2.5 rounded-xl text-[10px] text-amber-700 border border-amber-200 bg-amber-50 hover:bg-amber-100 transition-all font-bold">
                <FileSpreadsheet size={14} /> Eksport Excel/CSV
              </button>

              <button onClick={exportToJSON} className="flex items-center gap-2.5 w-full p-2.5 rounded-xl text-[10px] text-blue-700 border border-blue-200 bg-blue-50 hover:bg-blue-100 transition-all font-bold">
                <Download size={14} /> Backup Ke JSON
              </button>

              <button onClick={()=>fileInputRef.current?.click()} className="flex items-center gap-2.5 w-full p-2.5 rounded-xl text-[10px] text-emerald-700 border border-emerald-200 bg-emerald-50 hover:bg-emerald-100 transition-all font-bold">
                <Upload size={14} /> Import JSON
              </button>
              
              <input type="file" ref={fileInputRef} onChange={handleImportJSON} accept=".json" className="hidden" />
            </div>
          )}
          <button onClick={onLogout} className="flex items-center gap-2.5 w-full p-3 rounded-xl text-xs text-red-600 border border-red-200 bg-red-50 hover:bg-red-100 transition-all font-bold">
            <LogOut size={16} /> Keluar
          </button>
          <div className="pt-2 text-center">
            <span className="text-[9px] text-slate-300 font-bold uppercase tracking-[2px]">Design by Minangkaos</span>
          </div>
        </div>
      </aside>

      {/* ── CONTENT ── */}
      <main className="flex-1 ml-64 p-8 relative">
        <header className="flex justify-between items-center mb-10">
          <div>
            <h1 className="text-3xl font-bold tracking-tight text-slate-800">{view.charAt(0).toUpperCase() + view.slice(1)}</h1>
            <p className="text-slate-500 text-xs mt-1">
              Data terintegrasi secara realtime • Sync terakhir: {lastSync || "Menunggu..."}
            </p>
          </div>
          <div className="flex items-center gap-4">
            {sheetLoading && (
              <div className="flex items-center gap-2 text-emerald-600 text-[10px] font-bold uppercase tracking-widest bg-emerald-50 px-3 py-1.5 rounded-full border border-emerald-200">
                <RefreshCcw size={12} className="animate-spin" /> Sedang Sinkron
              </div>
            )}
            <button onClick={()=>fetchUpdatedData()} className="p-3 bg-white border border-slate-200 rounded-2xl hover:bg-slate-50 transition-all shadow-sm group">
              <RefreshCcw size={18} className="text-[#B8860B] group-active:rotate-180 transition-transform duration-500" />
            </button>
          </div>
        </header>

        <AnimatePresence mode="wait">
          <motion.div
            key={view}
            initial={{ opacity: 0, y: 10 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -10 }}
            transition={{ duration: 0.2 }}
          >
            {view==="dashboard" && (
              <div key="view-dash" className="space-y-8">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  <StatCard key="stat-total" label="Total Event" value={stats.total} icon={LayoutDashboard} color="#B8860B" />
                  <StatCard key="stat-target" label="Target Wisatawan" value={stats.target.toLocaleString("id-ID")} icon={User} color="#D4AF37" />
                </div>

                <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                  <div className="bg-white border border-slate-200 rounded-3xl p-8 shadow-sm">
                    <div className="flex items-center justify-between mb-8">
                      <h3 className="font-bold text-lg flex items-center gap-2 text-slate-800"><BarChart3 size={20} className="text-[#B8860B]" /> Distribusi Wilayah</h3>
                      <button onClick={()=>setView("list")} className="text-[10px] uppercase font-bold text-slate-400 hover:text-[#B8860B] transition-colors">Lihat Semua</button>
                    </div>
                    <div className="space-y-4 max-h-[400px] overflow-y-auto pr-2 custom-scroll">
                      {stats.kabCounts.map((k, i) => (
                        <div key={k.name} className="flex items-center gap-4">
                          <div className="text-[10px] font-sans text-slate-300 w-6">{(i+1).toString().padStart(2, '0')}</div>
                          <div className="flex-1">
                            <div className="flex justify-between text-xs mb-1.5">
                              <span className="font-bold text-slate-600">{k.name}</span>
                              <span className="text-[#B8860B] font-bold">{k.count}</span>
                            </div>
                            <div className="h-1.5 bg-slate-100 rounded-full overflow-hidden">
                              <motion.div 
                                initial={{ width: 0 }} 
                                animate={{ width: `${(k.count / stats.total) * 100}%` }}
                                className="h-full bg-gradient-to-r from-[#D4AF37] to-[#B8860B] rounded-full" 
                              />
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>

                  <div className="bg-white border border-slate-200 rounded-3xl p-8 shadow-sm">
                    <div className="flex items-center justify-between mb-8">
                      <h3 className="font-bold text-lg flex items-center gap-2 text-slate-800"><Clock size={20} className="text-emerald-600" /> Event Terbaru</h3>
                    </div>
                    <div className="space-y-2">
                      {events.slice(-6).reverse().map(e => (
                        <div key={e.id} onClick={()=>{setSelEvent(e); setView("detail");}} className="flex items-center gap-4 p-4 rounded-2xl bg-slate-50 border border-slate-100 hover:border-amber-200 hover:bg-amber-50/30 transition-all cursor-pointer group">
                          <div className="w-10 h-10 rounded-xl bg-white border border-slate-100 flex items-center justify-center flex-shrink-0 group-hover:scale-110 transition-transform shadow-sm">
                            <Tag size={16} className="text-slate-400" />
                          </div>
                          <div className="flex-1 min-w-0">
                            <h4 className="text-xs font-bold text-slate-700 truncate capitalize">{e.namaEvent}</h4>
                            <p className="text-[10px] text-slate-500 mt-0.5">{e.kabupatenKota}</p>
                          </div>
                          <StatusBadge status={e.status} />
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              </div>
            )}

            {view==="list" && (
              <div key="view-list" className="bg-white border border-slate-200 rounded-3xl overflow-hidden shadow-sm">
                <div className="p-8 border-b border-slate-100 bg-slate-50/50">
                   <div className="flex flex-wrap gap-4">
                      <div className="flex-1 min-w-[280px] relative">
                        <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
                        <input value={searchQ} onChange={e=>setSearchQ(e.target.value)} placeholder="Cari event, lokasi, atau deskripsi..." className="w-full bg-white border border-slate-200 rounded-2xl p-3.5 pl-12 text-sm focus:border-amber-400 outline-none transition-all placeholder:text-slate-400 shadow-sm" />
                      </div>
                      <div className="flex gap-4">
                        <select value={filterKab} onChange={e=>setFilterKab(e.target.value)} className="bg-white border border-slate-200 rounded-2xl p-3 px-4 text-xs font-semibold outline-none focus:border-amber-400 shadow-sm">
                          <option value="">Semua Wilayah</option>
                          {KABUPATEN_KOTA.map(k=><option key={k} value={k}>{k}</option>)}
                        </select>
                        <select value={filterKat} onChange={e=>setFilterKat(e.target.value)} className="bg-white border border-slate-200 rounded-2xl p-3 px-4 text-xs font-semibold outline-none focus:border-amber-400 shadow-sm">
                          <option value="">Semua Kategori</option>
                          {KATEGORI.map(k=><option key={k} value={k}>{k}</option>)}
                        </select>
                      </div>
                   </div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="text-[10px] uppercase font-bold text-slate-500 tracking-widest bg-slate-50/80">
                        <th className="px-8 py-5">Nama Event & Wilayah</th>
                        <th className="px-6 py-5">Kategori</th>
                        <th className="px-6 py-5">Tanggal</th>
                        <th className="px-8 py-5 text-right">Aksi</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {filtered.map(e => <EventRow key={e.id} e={e} onClick={() => {setSelEvent(e); setView("detail");}} canEdit={canEdit(e)} handleEdit={handleEdit} setDelConfirm={setDelConfirm} />)}
                    </tbody>
                  </table>
                </div>
                {filtered.length === 0 && (
                  <div className="p-20 text-center text-[#94A3B8]">
                    <Search className="mx-auto mb-4 opacity-10" size={48} />
                    <p className="text-sm">Tidak ditemukan event dengan kriteria tersebut.</p>
                  </div>
                )}
              </div>
            )}

            {view==="form" && (
              <div key="view-form" className="max-w-3xl mx-auto bg-white border border-slate-200 rounded-3xl p-10 shadow-lg">
                <div className="flex items-center gap-4 mb-10 pb-6 border-b border-slate-100">
                   <div className="w-12 h-12 bg-amber-50 rounded-2xl flex items-center justify-center border border-amber-200 text-amber-700">
                      {editId ? <Edit2 size={24} /> : <PlusCircle size={24} />}
                   </div>
                   <div>
                     <h3 className="text-xl font-bold text-slate-800">{editId ? "Ubah Data Event" : "Tambah Event Baru"}</h3>
                     <p className="text-slate-500 text-xs">Pastikan seluruh data yang berbintang (*) terisi dengan benar.</p>
                   </div>
                </div>

                <div className="space-y-8">
                   <FormGrid title="Informasi Utama">
                     <Inp label="Nama Event *" full val={form.namaEvent} onChange={v=>setForm(f=>({...f,namaEvent:v}))} />
                     <div className="col-span-1">
                        <label className="text-[11px] uppercase font-bold text-slate-400 mb-2 block tracking-wider">Wilayah *</label>
                        {isProvinsi ? (
                          <select value={form.kabupatenKota} onChange={e=>setForm(f=>({...f,kabupatenKota:e.target.value}))} className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3.5 text-sm outline-none focus:border-amber-400">
                            <option value="">-- Pilih Wilayah --</option>
                            {KABUPATEN_KOTA.map(k=><option key={k} value={k}>{k}</option>)}
                          </select>
                        ) : (
                          <input value={user.kabupatenKota} disabled className="w-full bg-slate-50 border border-slate-100 text-slate-400 p-3.5 rounded-xl text-sm cursor-not-allowed" />
                        )}
                     </div>
                     <div className="col-span-1">
                        <label className="text-[11px] uppercase font-bold text-slate-400 mb-2 block tracking-wider">Kategori *</label>
                        <select value={form.kategori} onChange={e=>setForm(f=>({...f,kategori:e.target.value}))} className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3.5 text-sm outline-none focus:border-amber-400">
                          <option value="">-- Pilih Kategori --</option>
                          {KATEGORI.map(k=><option key={k} value={k}>{k}</option>)}
                        </select>
                     </div>
                     <Inp label="Tanggal Mulai *" type="date" val={form.tanggalMulai} onChange={v=>setForm(f=>({...f,tanggalMulai:v}))} />
                     <Inp label="Tanggal Selesai" type="date" val={form.tanggalSelesai || ""} onChange={v=>setForm(f=>({...f,tanggalSelesai:v}))} />
                     <Inp label="Lokasi / Venue" val={form.lokasi || ""} onChange={v=>setForm(f=>({...f,lokasi:v}))} full />
                     <div className="col-span-2">
                        <label className="text-[11px] uppercase font-bold text-slate-400 mb-2 block tracking-wider">Deskripsi Singkat</label>
                        <textarea value={form.deskripsi} onChange={e=>setForm(f=>({...f,deskripsi:e.target.value}))} rows={4} className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3.5 text-sm outline-none focus:border-amber-400 resize-none placeholder:text-slate-400" />
                     </div>
                   </FormGrid>

                   <FormGrid title="Teknis & Anggaran">
                     <Inp label="Target Wisatawan" type="number" val={form.targetWisatawan || ""} onChange={v=>setForm(f=>({...f,targetWisatawan:v}))} />
                     <Inp label="Estimasi Anggaran (Rp)" type="number" val={form.anggaran || ""} onChange={v=>setForm(f=>({...f,anggaran:v}))} />
                   </FormGrid>

                   <FormGrid title="Informasi Narahubung">
                     <Inp label="Nama Penanggung Jawab" val={form.kontakNama || ""} onChange={v=>setForm(f=>({...f,kontakNama:v}))} />
                     <Inp label="Nomer Telp / WA" val={form.kontakHP || ""} onChange={v=>setForm(f=>({...f,kontakHP:v}))} />
                   </FormGrid>

                   <div className="flex gap-4 pt-4">
                      <button onClick={handleSubmit} className="flex-1 py-4 bg-gradient-to-r from-amber-500 to-amber-700 text-white rounded-2xl font-bold uppercase tracking-widest shadow-lg shadow-amber-500/20 hover:scale-[1.01] transition-all">
                        {editId ? "Simpan Perubahan" : "Simpan Data Event"}
                      </button>
                      <button onClick={()=>setView("list")} className="px-10 py-4 bg-slate-100 border border-slate-200 rounded-2xl font-bold uppercase tracking-widest text-slate-500 hover:bg-slate-200 transition-all">
                        Batal
                      </button>
                   </div>
                </div>
              </div>
            )}

            {view==="calendar" && (
               <div key="view-cal" className="bg-white border border-slate-200 rounded-3xl p-10 shadow-sm">
                 <div className="flex items-center justify-between mb-10">
                    <div className="flex items-center gap-6">
                       <button onClick={()=>{if(calMonth===0){setCalMonth(11);setCalYear(y=>y-1);}else setCalMonth(m=>m-1);}} className="w-12 h-12 bg-white border border-slate-200 rounded-2xl flex items-center justify-center hover:bg-slate-50 transition-all text-[#B8860B] shadow-sm"><ChevronLeft size={24} /></button>
                       <h3 className="text-2xl font-bold w-48 text-center text-slate-800">{BULAN[calMonth]} {calYear}</h3>
                       <button onClick={()=>{if(calMonth===11){setCalMonth(0);setCalYear(y=>y+1);}else setCalMonth(m=>m+1);}} className="w-12 h-12 bg-white border border-slate-200 rounded-2xl flex items-center justify-center hover:bg-slate-50 transition-all text-[#B8860B] shadow-sm"><ChevronRight size={24} /></button>
                    </div>
                    <div className="flex items-center gap-4 text-[11px] text-slate-400 font-bold uppercase tracking-wider">
                        <div className="flex items-center gap-2"><div className="w-3 h-3 bg-amber-400 rounded-sm" /> Berlangsung</div>
                        <div className="flex items-center gap-2"><div className="w-3 h-3 bg-slate-100 rounded-sm" /> Kosong</div>
                    </div>
                 </div>
                 <div className="grid grid-cols-7 gap-1">
                    {["Min","Sen","Sel","Rab","Kam","Jum","Sab"].map(d=> (
                      <div key={d} className="text-center text-[10px] uppercase font-bold text-slate-400 py-4 bg-slate-50/50 border border-slate-100 rounded-xl mb-2">{d}</div>
                    ))}
                    {Array(getFD(calYear,calMonth)).fill(null).map((_,i)=><div key={`pad-${i}`} className="min-h-[120px] rounded-2xl border border-transparent opacity-10" />)}
                    {Array(getDIM(calYear,calMonth)).fill(null).map((_,i)=>{
                      const day=i+1;
                      const dayEvs=eInDay(day);
                      const isToday=new Date().getFullYear()===calYear&&new Date().getMonth()===calMonth&&new Date().getDate()===day;
                      const dateObj = new Date(calYear, calMonth, day);
                      const isSunday = dateObj.getDay() === 0;
                      const dateStr = `${calYear}-${(calMonth + 1).toString().padStart(2, '0')}-${day.toString().padStart(2, '0')}`;
                      const holiday = HOLIDAYS_2026[dateStr];

                      return (
                        <div key={`day-${day}`} className={`min-h-[120px] rounded-2xl border p-3 flex flex-col gap-1.5 transition-all group hover:z-10 hover:scale-[1.03] hover:shadow-xl ${isToday ? "bg-amber-50 border-amber-200" : (holiday || isSunday) ? "bg-red-50/50 border-red-100" : "bg-white border-slate-100 hover:border-slate-200"}`}>
                           <div className="flex justify-between items-start">
                             <div className={`text-2xl font-black opacity-80 ${isToday ? "text-amber-600" : (holiday || isSunday) ? "text-red-500" : "text-slate-200 group-hover:text-slate-400"}`}>{day}</div>
                             {holiday && <div className="text-[7px] md:text-[8px] font-black text-red-600 bg-red-100/50 px-1.5 py-0.5 rounded uppercase leading-tight mt-1 text-right whitespace-normal break-words flex-1 max-w-[65%]" title={holiday}>{holiday}</div>}
                           </div>
                           <div className="flex-1 space-y-1 overflow-y-auto max-h-[80px] custom-scroll">
                              {dayEvs.map((e, idx)=>(
                                <div key={`${e.id}-${idx}`} onClick={()=>{setSelEvent(e); setView("detail");}} className="text-[10px] leading-tight bg-slate-50 p-1.5 rounded-lg border border-slate-100 text-slate-600 hover:bg-amber-100 hover:text-amber-800 hover:border-amber-200 cursor-pointer truncate font-medium">
                                  {e.namaEvent}
                                </div>
                              ))}
                           </div>
                        </div>
                      );
                    })}
                 </div>
               </div>
            )}

            {view==="detail" && selEvent && (
              <div key="view-detail" className="max-w-4xl mx-auto space-y-8">
                 <button onClick={()=>setView("list")} className="flex items-center gap-2 text-xs font-bold text-amber-700 bg-amber-50 px-5 py-2.5 rounded-xl border border-amber-200 hover:bg-amber-100 transition-all">
                    <ChevronLeft size={16} /> Kembali ke Daftar
                 </button>
                 
                 <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
                   <div className="lg:col-span-2 space-y-8">
                      <div className="bg-white border border-slate-200 rounded-3xl p-10 relative overflow-hidden shadow-sm">
                         <div className="absolute top-0 left-0 w-2 h-full bg-[#B8860B]" />
                         <div className="flex items-start justify-between mb-8">
                           <div>
                              <StatusBadge status={selEvent.status} />
                              <h2 className="text-3xl font-black mt-4 tracking-tight text-slate-800 capitalize">{selEvent.namaEvent}</h2>
                              <div className="flex items-center gap-4 mt-6 text-sm text-slate-500 font-medium">
                                <div className="flex items-center gap-2 pr-4 border-r border-slate-200"><MapPin size={16} className="text-amber-600" /> {selEvent.kabupatenKota}</div>
                                <div className="flex items-center gap-2"><Tag size={16} className="text-amber-600" /> {selEvent.kategori}</div>
                              </div>
                           </div>
                         </div>
                         <div className="space-y-6">
                            <div>
                               <label className="text-[10px] uppercase font-bold text-slate-400 mb-2 block tracking-widest">Deskripsi</label>
                               <div className="text-slate-600 leading-relaxed text-sm whitespace-pre-wrap">{selEvent.deskripsi || "Tidak ada deskripsi tersedia."}</div>
                            </div>
                            <div className="grid grid-cols-2 gap-6 pt-6 border-t border-slate-100">
                               <div>
                                  <label className="text-[10px] uppercase font-bold text-slate-400 mb-2 block tracking-widest">Lokasi</label>
                                  <div className="text-sm font-bold text-slate-700">{selEvent.lokasi || "-"}</div>
                                </div>
                               <div>
                                  <label className="text-[10px] uppercase font-bold text-slate-400 mb-2 block tracking-widest">Periode</label>
                                  <div className="text-sm font-bold text-slate-700">{selEvent.tanggalMulai} {selEvent.tanggalSelesai ? ` - ${selEvent.tanggalSelesai}` : ""}</div>
                               </div>
                            </div>
                         </div>
                      </div>
                   </div>

                   <div className="space-y-8">
                      <div className="bg-white border border-slate-200 rounded-3xl p-8 shadow-sm">
                         <h4 className="font-bold text-slate-700 mb-6 flex items-center gap-2 border-b border-slate-100 pb-4 uppercase text-[10px] tracking-widest"><Info size={18} className="text-[#B8860B]" /> Data Teknis</h4>
                         <div className="space-y-6">
                            <DetailRow label="Target Pengunjung" val={selEvent.targetWisatawan ? Number(selEvent.targetWisatawan).toLocaleString("id-ID") + " Wisatawan" : "-"} />
                            <DetailRow label="Anggaran Dana" val={"Rp " + Number(selEvent.anggaran || 0).toLocaleString("id-ID")} />
                            <DetailRow label="Kontak Person" val={selEvent.kontakNama || "-"} />
                            <DetailRow label="Nomer Ponsel (WA)" val={selEvent.kontakHP || "-"} />
                         </div>
                      </div>

                      {canEdit(selEvent) && (
                        <div className="bg-slate-50 p-8 rounded-3xl border border-slate-100 space-y-4">
                           {selEvent.status === "Draft" && (
                             <button onClick={()=>handleAjukan(selEvent.id)} className="w-full py-4 bg-emerald-600 text-white rounded-2xl font-bold uppercase tracking-widest text-[11px] hover:scale-[1.02] shadow-lg shadow-emerald-600/20 transition-all">Ajukan ke Provinsi</button>
                           )}
                           <div className="grid grid-cols-2 gap-4">
                              <button onClick={()=>handleEdit(selEvent)} className="py-3 bg-white border border-slate-200 rounded-xl text-xs font-bold uppercase tracking-wider text-slate-600 hover:bg-slate-50 transition-all shadow-sm"><Edit2 size={14} className="inline mr-2" /> Edit</button>
                              <button onClick={(ev)=>{ev.stopPropagation(); setDelConfirm(selEvent.id);}} className="py-3 bg-red-50 border border-red-100 text-red-600 rounded-xl text-xs font-bold uppercase tracking-wider hover:bg-red-100 transition-all shadow-sm"><Trash2 size={14} className="inline mr-2" /> Hapus</button>
                           </div>
                        </div>
                      )}
                   </div>
                 </div>
              </div>
            )}
            
            {view==="reports" && <div className="p-20 text-center text-[#94A3B8]">Fitur Laporan Analitik mendalam sedang dalam pengembangan sistem.</div>}
            {view==="help" && <div className="p-20 text-center text-[#94A3B8]">Bantuan pengoperasian CMS CoE Sumbar sedang disusun oleh tim teknis.</div>}

          </motion.div>
        </AnimatePresence>
      </main>

      {/* ── ALERTS & TOASTS ── */}
      {delConfirm !== null && (
        <div className="fixed inset-0 z-[1000] flex items-center justify-center bg-slate-900/40 backdrop-blur-sm p-4">
          <motion.div initial={{ scale: 0.9, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} className="bg-white border border-slate-200 rounded-3xl p-10 max-w-sm w-full text-center shadow-2xl">
            <div className="w-16 h-16 bg-red-50 rounded-full flex items-center justify-center mx-auto mb-6 border border-red-100 text-red-600"><Trash2 size={32} /></div>
            <h3 className="text-xl font-bold mb-2 text-slate-800">Hapus Event?</h3>
            <p className="text-xs text-slate-500 mb-8">Data akan dihapus secara permanen dari server Google Sheets.</p>
            <div className="flex gap-4">
              <button onClick={()=>{ if(delConfirm !== null) handleDelete(delConfirm); }} className="flex-1 py-3 bg-red-600 text-white rounded-xl font-bold text-xs uppercase cursor-pointer hover:bg-red-700 transition-colors shadow-lg shadow-red-600/20">Ya, Hapus</button>
              <button onClick={()=>setDelConfirm(null)} className="flex-1 py-3 bg-slate-100 text-slate-500 rounded-xl font-bold text-xs uppercase cursor-pointer hover:bg-slate-200 transition-colors">Batal</button>
            </div>
          </motion.div>
        </div>
      )}

      {toast && (
        <motion.div initial={{ x: 100, opacity: 0 }} animate={{ x: 10, opacity: 1 }} className={`fixed bottom-8 right-8 z-[2000] p-4 px-6 rounded-2xl border shadow-2xl flex items-center gap-3 transition-all ${toast.type==="error" ? "bg-red-50 border-red-200 text-red-800" : "bg-white border-slate-200 text-slate-800"}`}>
          <div className={toast.type==="error" ? "text-red-500" : "text-emerald-500"}>{toast.type==="error" ? <Info size={18} /> : <CheckCircle2 size={18} />}</div>
          <div className="text-xs font-bold uppercase tracking-widest">{toast.msg}</div>
        </motion.div>
      )}
    </div>
  );
}

// ── SUBCOMPONENTS ─────────────────────────────────────────────
function MenuBtn({ icon: Icon, label, active, onClick }: { icon: any; label: string; active: boolean; onClick: () => void }) {
  return (
    <button onClick={onClick} className={`flex items-center gap-3 w-full p-3.5 px-4 rounded-2xl text-[13px] font-bold tracking-tight transition-all border outline-none ${active ? "bg-amber-50 border-amber-200 text-amber-700 shadow-sm" : "bg-transparent border-transparent text-slate-400 hover:bg-slate-50 hover:text-slate-700"}`}>
      <Icon size={18} className={active ? "text-amber-600" : "text-slate-400 opacity-70"} />
      {label}
    </button>
  );
}

function StatCard({ label, value, icon: Icon, color }: { label: string; value: string | number; icon: any; color: string }) {
  return (
    <div className="bg-white border border-slate-200 rounded-3xl p-8 relative overflow-hidden group hover:border-amber-200 transition-all shadow-sm">
       <div className="absolute -right-4 -bottom-4 opacity-[0.05] group-hover:opacity-[0.1] transition-opacity text-slate-200"><Icon size={140} /></div>
       <div style={{ backgroundColor: `${color}10`, borderColor: `${color}30` }} className="w-12 h-12 rounded-2xl flex items-center justify-center border mb-6 relative z-10 shadow-sm">
          <Icon size={24} style={{ color }} />
       </div>
       <div className="text-3xl font-black tracking-tight mb-1 relative z-10 text-slate-800">{value}</div>
       <div className="text-[10px] uppercase font-bold text-slate-400 tracking-[2px] relative z-10">{label}</div>
    </div>
  );
}

function Inp({ label, type="text", val, onChange, full, placeholder }: { label: string; type?: string; val: string | number; onChange: (v: string)=>void; full?: boolean; placeholder?: string }) {
  return (
    <div className={full ? "col-span-2" : "col-span-1"}>
       <label className="text-[11px] uppercase font-bold text-slate-400 mb-2 block tracking-wider ml-1">{label}</label>
       <input type={type} value={val} onChange={e=>onChange(e.target.value)} placeholder={placeholder} className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3.5 text-sm outline-none focus:border-amber-400 focus:bg-white transition-all text-slate-800 shadow-sm placeholder:text-slate-300" />
    </div>
  );
}

function FormGrid({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <div>
      <h4 className="text-[10px] uppercase font-bold text-amber-600 mb-5 tracking-[3px] flex items-center gap-3">
         {title} <div className="flex-1 h-px bg-slate-100" />
      </h4>
      <div className="grid grid-cols-2 gap-6">{children}</div>
    </div>
  );
}

function DetailRow({ label, val }: { label: string; val: string }) {
  return (
    <div className="flex justify-between items-center py-1">
      <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">{label}</span>
      <span className="text-xs font-bold text-slate-700">{val}</span>
    </div>
  );
}

function StatusBadge({ status }: { status: string }) {
  const cfg = STATUS_BADGE[status] || STATUS_BADGE["Draft"];
  return (
    <div style={{ backgroundColor: cfg.bg, color: cfg.color, borderColor: cfg.border }} className="inline-flex items-center px-2.5 py-1 rounded-full border text-[9px] font-bold uppercase tracking-widest">
      {status === "Diajukan" ? <CheckCircle2 size={10} className="mr-1.5" /> : <Clock size={10} className="mr-1.5" />}
      {status}
    </div>
  );
}

function EventRow({ e, onClick, canEdit, handleEdit, setDelConfirm }: { e: EventData; onClick: () => void; canEdit: boolean; handleEdit: (e: EventData) => void; setDelConfirm: (id: number) => void }) {
  return (
    <tr className="group hover:bg-slate-50 transition-colors cursor-pointer" onClick={onClick}>
      <td className="px-8 py-5">
        <div className="font-bold text-sm text-slate-700 group-hover:text-amber-700 transition-colors capitalize">{e.namaEvent}</div>
        <div className="text-[10px] text-slate-400 mt-1 flex items-center gap-1.5 font-medium"><MapPin size={10} /> {e.kabupatenKota}</div>
      </td>
      <td className="px-6 py-5">
        <div className="text-xs text-slate-500 font-medium">{e.kategori}</div>
      </td>
      <td className="px-6 py-5">
        <div className="text-xs font-sans text-slate-600">{e.tanggalMulai}</div>
      </td>
      <td className="px-8 py-5 text-right" onClick={ev=>ev.stopPropagation()}>
        <div className="flex justify-end gap-2 text-slate-400">
            {canEdit && (
              <>
                <RowBtn key={`edit-${e.id}`} icon={Edit2} color="#B8860B" onClick={(ev)=>{ ev.stopPropagation(); handleEdit(e); }} />
                <RowBtn key={`del-${e.id}`} icon={Trash2} color="#E11D48" onClick={(ev)=>{ ev.stopPropagation(); setDelConfirm(e.id); }} />
              </>
            )}
            <RowBtn key={`view-${e.id}`} icon={ExternalLink} color="#64748B" onClick={(ev)=>{ ev.stopPropagation(); onClick(); }} />
          </div>
      </td>
    </tr>
  );
}

function RowBtn({ icon: Icon, color, onClick }: { icon: any; color: string; onClick: (e: any)=>void }) {
  return (
    <button onClick={onClick} className="w-9 h-9 flex items-center justify-center rounded-xl bg-slate-50 border border-slate-200 transition-all hover:scale-110 active:scale-95 shadow-sm" style={{ color }}>
      <Icon size={14} />
    </button>
  );
}
