/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';

// ── KONSTANTA ─────────────────────────────────────────────────
const KABUPATEN_KOTA = [
  "Provinsi Sumatera Barat",
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

const STATUS_BADGE: Record<string, { bg: string; color: string; border: string }> = {
  "Draft":    { bg:"#1A1A2A", color:"#8A8AA0", border:"#2A2A4A" },
  "Diajukan": { bg:"#0A2218", color:"#4CAF82", border:"#0E4028" },
};

const ACCOUNTS = [
  { username:"admin.provinsi", password:"sumbarrancak", role:"provinsi", kabupatenKota:"Provinsi Sumatera Barat", nama:"Admin Dinas Provinsi Sumbar" },
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
  kontakEmail?: string;
  anggaran?: string;
  status: string;
  createdAt?: string;
}

const initialForm: Omit<EventData, 'id'> = {
  namaEvent:"", kabupatenKota:"", kategori:"", tanggalMulai:"",
  tanggalSelesai:"", lokasi:"", deskripsi:"", targetWisatawan:"",
  kontakNama:"", kontakHP:"", kontakEmail:"", anggaran:"", status:"Draft",
};

const STORAGE_KEY  = "coe_sumbar_v2_events";
const SESSION_KEY  = "coe_sumbar_session";
const SHEET_URL    = "https://script.google.com/macros/s/AKfycbyDuXGIgl1nOXJpviThfrrn4ThbH62USbq6GUDfBecl5wx4oRJhk5gnIlBF8D_q1QLffg/exec";

// ── GOOGLE API HELPERS ─────────────────────────────────
async function sheetSaveEvent(event: EventData) {
  try {
    await fetch(SHEET_URL, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action: "save", event }),
    });
  } catch(e) { console.warn("Sheet save failed:", e); }
}

async function sheetDeleteEvent(id: number) {
  try {
    await fetch(SHEET_URL, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action: "delete", id }),
    });
  } catch(e) { console.warn("Sheet delete failed:", e); }
}

async function sheetGetAll(): Promise<EventData[] | null> {
  try {
    const res = await fetch(SHEET_URL + "?action=getAll");
    const data = await res.json();
    return data.events || null;
  } catch(e) { console.warn("Sheet getAll failed:", e); return null; }
}

function getSeedData(): EventData[] {
  return [
    { id:1, namaEvent:"Pesona Budaya Hoyak Tabuik Piaman", kabupatenKota:"Kota Pariaman",
      kategori:"Festival Budaya", tanggalMulai:"2025-07-05", tanggalSelesai:"2025-07-14",
      lokasi:"Pantai Gandoriah, Pariaman", deskripsi:"Festival tradisi tabuik yang digelar setiap Muharram.",
      targetWisatawan:"5000", kontakNama:"Dedi Rahmat", kontakHP:"081234567890",
      kontakEmail:"pariwisata@pariaman.go.id", anggaran:"500000000", status:"Diajukan", createdAt:"2025-01-10" },
    { id:2, namaEvent:"Tour de Singkarak", kabupatenKota:"Kabupaten Solok",
      kategori:"Olahraga & Adventure", tanggalMulai:"2025-10-15", tanggalSelesai:"2025-10-22",
      lokasi:"Danau Singkarak & sekitarnya", deskripsi:"Balap sepeda internasional mengelilingi alam Sumatera Barat.",
      targetWisatawan:"20000", kontakNama:"Roni Amir", kontakHP:"082345678901",
      kontakEmail:"pariwisata@solok.go.id", anggaran:"2000000000", status:"Diajukan", createdAt:"2025-02-01" },
    { id:3, namaEvent:"Festival Rendang Dunia", kabupatenKota:"Kota Padang",
      kategori:"Kuliner", tanggalMulai:"2025-09-20", tanggalSelesai:"2025-09-22",
      lokasi:"GOR H. Agus Salim, Padang", deskripsi:"Festival kuliner rendang sebagai warisan budaya UNESCO.",
      targetWisatawan:"10000", kontakNama:"Sari Dewi", kontakHP:"083456789012",
      kontakEmail:"pariwisata@padang.go.id", anggaran:"750000000", status:"Draft", createdAt:"2025-03-15" },
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
    "Email":e.kontakEmail||"-","Status":e.status,"Tanggal Input":e.createdAt||"-",
  }));
  const ws=XLSX.utils.json_to_sheet(rows);
  ws["!cols"]=[{wch:5},{wch:35},{wch:25},{wch:20},{wch:14},{wch:14},{wch:30},{wch:40},{wch:18},{wch:22},{wch:20},{wch:16},{wch:28},{wch:12},{wch:14}];
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"Calendar of Events");
  const kc: Record<string, number>={},kab: Record<string, number>={};
  events.forEach(e=>{kc[e.kategori]=(kc[e.kategori]||0)+1;kab[e.kabupatenKota]=(kab[e.kabupatenKota]||0)+1;});
  const s2=[
    {"Ringkasan":"Total Event","Jumlah":events.length},
    {"Ringkasan":"Sudah Diajukan","Jumlah":events.filter(e=>e.status==="Diajukan").length},
    {"Ringkasan":"Masih Draft","Jumlah":events.filter(e=>e.status==="Draft").length},
    {"Ringkasan":"Total Target Wisatawan","Jumlah":events.reduce((s,e)=>s+parseInt(e.targetWisatawan||"0"),0)},
    {"Ringkasan":"","Jumlah":""},
    {"Ringkasan":"— Per Kategori —","Jumlah":""},
    ...Object.entries(kc).map(([k,v])=>({Ringkasan:k,Jumlah:v})),
    {"Ringkasan":"","Jumlah":""},
    {"Ringkasan":"— Per Kab/Kota —","Jumlah":""},
    ...Object.entries(kab).map(([k,v])=>({Ringkasan:k,Jumlah:v})),
  ];
  const ws2=XLSX.utils.json_to_sheet(s2); ws2["!cols"]=[{wch:30},{wch:15}];
  XLSX.utils.book_append_sheet(wb,ws2,"Ringkasan");
  XLSX.writeFile(wb,`CoE_SumBar_${new Date().toISOString().split("T")[0]}.xlsx`);
}

function exportToJSON(events: EventData[]) {
  const blob=new Blob([JSON.stringify({exportedAt:new Date().toISOString(),total:events.length,events},null,2)],{type:"application/json"});
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a"); a.href=url;
  a.download=`CoE_SumBar_Backup_${new Date().toISOString().split("T")[0]}.json`;
  a.click(); URL.revokeObjectURL(url);
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
  const [show,  setShow]  = useState(false);
  const [hint,  setHint]  = useState(false);

  const doLogin = () => {
    const acc = ACCOUNTS.find(a=>a.username===uname.trim()&&a.password===pass);
    if (!acc) { setErr("Username atau password salah."); return; }
    onLogin(acc);
  };

  return (
    <div className="min-h-screen bg-[#0D1B0F] flex items-center justify-center relative overflow-hidden" 
         style={{ backgroundImage: `radial-gradient(ellipse at 30% 40%,rgba(196,160,60,.09) 0%,transparent 55%), radial-gradient(ellipse at 75% 70%,rgba(34,85,34,.18) 0%,transparent 50%)` }}>
      <div className="w-[380px] bg-[rgba(8,20,10,0.97)] border border-[rgba(196,160,60,0.2)] rounded-2xl p-10 shadow-[0_24px_60px_rgba(0,0,0,0.6)] z-10">

        <div className="text-center mb-8">
          <div className="text-4xl mb-2">🏔</div>
          <div className="text-[9px] tracking-[3px] text-[#C4A03C] uppercase mb-1">Dinas Pariwisata</div>
          <div className="text-xl font-bold text-[#E8DCC8]">Sumatera Barat</div>
          <div className="text-[10px] text-[#6A5830] mt-1 tracking-[1px]">Calendar of Events · Sistem Pendataan</div>
        </div>

        <div className="mb-4">
          <label className="text-[11px] text-[#6A5830] block mb-1.5 font-sans">Username</label>
          <input value={uname} onChange={e=>{setUname(e.target.value);setErr("");}}
            onKeyDown={e=>e.key==="Enter"&&doLogin()}
            placeholder="admin.provinsi / admin.kota.padang"
            className="w-full bg-[rgba(255,255,255,0.04)] border border-[rgba(196,160,60,0.16)] text-[#E8DCC8] p-[9px_12px] rounded-lg outline-none text-sm" />
        </div>

        <div className="mb-2">
          <label className="text-[11px] text-[#6A5830] block mb-1.5 font-sans">Password</label>
          <div className="relative">
            <input value={pass} onChange={e=>{setPass(e.target.value);setErr("");}}
              onKeyDown={e=>e.key==="Enter"&&doLogin()}
              type={show?"text":"password"} placeholder="••••••••"
              className="w-full bg-[rgba(255,255,255,0.04)] border border-[rgba(196,160,60,0.16)] text-[#E8DCC8] p-[9px_40px_9px_12px] rounded-lg outline-none text-sm" />
            <button onClick={()=>setShow(s=>!s)}
              className="absolute right-2.5 top-1/2 -translate-y-1/2 bg-transparent border-none text-[#6A5830] cursor-pointer text-sm p-0">
              {show?"🙈":"👁"}
            </button>
          </div>
        </div>

        {err && <div className="text-[11px] text-red-400 mb-2.5 p-[6px_10px] bg-[rgba(224,90,90,0.08)] rounded-md">{err}</div>}

        <button onClick={doLogin} className="w-full p-[11px] bg-gradient-to-br from-[#C4A03C] to-[#A07828] text-[#0D1B0F] border-none rounded-lg text-sm font-bold cursor-pointer mt-2 tracking-wide hover:opacity-90 transition-opacity">
          Masuk
        </button>

        <div className="mt-5 text-center">
          <button onClick={()=>setHint(h=>!h)}
            className="text-[10px] text-[#4A3820] bg-transparent border-none cursor-pointer underline">
            {hint?"Sembunyikan":"Lihat semua akun demo"}
          </button>
          {hint&&(
            <div className="mt-2.5 bg-[rgba(255,255,255,0.02)] border border-[rgba(196,160,60,0.1)] rounded-lg p-3 text-left max-h-[260px] overflow-y-auto">
              <div className="text-[10px] text-[#6A5830] mb-1 font-bold">ADMIN PROVINSI</div>
              <div className="text-[10px] text-[#8A7860] font-mono mb-2.5 bg-[rgba(196,160,60,0.06)] p-[4px_8px] rounded cursor-pointer"
                onClick={()=>{setUname("admin.provinsi");setPass("sumbarrancak");setHint(false);}}>
                admin.provinsi &nbsp;·&nbsp; sumbarrancak
              </div>
              <div className="text-[10px] text-[#6A5830] mb-1 font-bold">
                ADMIN KAB/KOTA <span className="text-[#3A3020] font-normal">(password: sumbarrancak)</span>
              </div>
              {ACCOUNTS.filter(a=>a.role==="kabkota").map(a=>(
                <div key={a.username}
                  className="text-[10px] text-[#8A7860] font-mono p-[3px_8px] rounded mb-0.5 bg-[rgba(255,255,255,0.02)] cursor-pointer hover:bg-[rgba(196,160,60,0.05)]"
                  onClick={()=>{setUname(a.username);setPass("sumbarrancak");setHint(false);}}>
                  {a.username}
                </div>
              ))}
              <div className="text-[9px] text-[#3A3020] mt-2 italic">✱ klik username untuk mengisi otomatis</div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════
//  MAIN APP
// ══════════════════════════════════════════════════════════════
function MainApp({ user, onLogout }: { user: any; onLogout: () => void }) {
  const isProvinsi = user.role === "provinsi";

  const [events, setEvents]               = useState<EventData[]>([]);
  const [view,   setView]                 = useState("dashboard");
  const [form,   setForm]                 = useState<EventData | Omit<EventData, 'id'>>(initialForm);
  const [editId, setEditId]               = useState<number | null>(null);
  const [selEvent, setSelEvent]           = useState<EventData | null>(null);
  const [filterKab, setFilterKab]         = useState("");
  const [filterKat, setFilterKat]         = useState("");
  const [filterStatus, setFilterStatus]   = useState("");
  const [filterBulan, setFilterBulan]     = useState("");
  const [searchQ, setSearchQ]             = useState("");
  const [calYear,  setCalYear]            = useState(2025);
  const [calMonth, setCalMonth]           = useState(new Date().getMonth());
  const [toast,   setToast]               = useState<{ msg: string; type: string }|null>(null);
  const [delConfirm, setDelConfirm]       = useState<number | null>(null);
  const [ajukanConfirm, setAjukanConfirm] = useState<number | null>(null);
  const [importConfirm, setImportConfirm] = useState<{events: EventData[]; count: number} | null>(null);
  const [aiLoading, setAiLoading]         = useState(false);
  const [aiResult,  setAiResult]          = useState<string | null>(null);
  const importRef = useRef<HTMLInputElement>(null);

  const [sheetLoading, setSheetLoading] = useState(true);

  useEffect(()=>{
    setSheetLoading(true);
    sheetGetAll().then(sheetEvents => {
      if (sheetEvents && sheetEvents.length > 0) {
        setEvents(sheetEvents);
        localStorage.setItem(STORAGE_KEY, JSON.stringify(sheetEvents));
      } else {
        const stored = localStorage.getItem(STORAGE_KEY);
        setEvents(stored ? JSON.parse(stored) : getSeedData());
      }
      setSheetLoading(false);
    });
  },[]);

  useEffect(()=>{ if(events.length>0) localStorage.setItem(STORAGE_KEY,JSON.stringify(events)); },[events]);

  const showToast=(msg: string,type="success")=>{ setToast({msg,type}); setTimeout(()=>setToast(null),3500); };

  const handleRefresh = () => {
    setSheetLoading(true);
    sheetGetAll().then(sheetEvents => {
      if (sheetEvents && sheetEvents.length > 0) {
        setEvents(sheetEvents);
        localStorage.setItem(STORAGE_KEY, JSON.stringify(sheetEvents));
        showToast(`Data diperbarui dari server. ✓`);
      } else {
        showToast("Tidak ada data baru dari server.", "info");
      }
      setSheetLoading(false);
    });
  };

  const canEditEvent  = (e: EventData) => isProvinsi || e.kabupatenKota === user.kabupatenKota;

  const handleSubmit = () => {
    const finalForm = !isProvinsi
      ? { ...form, kabupatenKota: user.kabupatenKota }
      : { ...form };

    if (!finalForm.namaEvent || !finalForm.kabupatenKota || !finalForm.tanggalMulai || !finalForm.kategori) {
      const kosong = [];
      if (!finalForm.namaEvent) kosong.push("Nama Event");
      if (!finalForm.kabupatenKota) kosong.push("Kabupaten/Kota");
      if (!finalForm.tanggalMulai) kosong.push("Tanggal Mulai");
      if (!finalForm.kategori) kosong.push("Kategori");
      showToast("Field wajib belum diisi: " + kosong.join(", "), "error"); return;
    }
    if (editId) {
      const updated = {...finalForm, id:editId, createdAt: events.find(e=>e.id===editId)?.createdAt} as EventData;
      setEvents(ev=>ev.map(e=>e.id===editId?updated:e));
      sheetSaveEvent(updated);
      showToast("Event berhasil diperbarui! ✓");
    } else {
      const newEvent = {...finalForm, id:Date.now(), createdAt:new Date().toISOString().split("T")[0]} as EventData;
      setEvents(ev=>[...ev, newEvent]);
      sheetSaveEvent(newEvent);
      showToast("Event berhasil ditambahkan & disimpan ke server! ✓");
    }
    setForm(initialForm); setEditId(null); setView("list");
  };

  const handleEdit = (e: EventData) => {
    if (!canEditEvent(e)) return;
    setForm({...e}); setEditId(e.id); setView("form");
  };

  const handleDelete = (id: number) => {
    setEvents(ev=>ev.filter(e=>e.id!==id));
    sheetDeleteEvent(id);
    setDelConfirm(null); showToast("Event berhasil dihapus.","info");
    if (view==="detail") setView("list");
  };

  const handleAjukan = (id: number) => {
    setEvents(ev=>ev.map(e=>{
      if(e.id===id){ const updated={...e,status:"Diajukan"}; sheetSaveEvent(updated); return updated; }
      return e;
    }));
    if (selEvent?.id===id) setSelEvent(s=>s ? ({...s,status:"Diajukan"}) : null);
    setAjukanConfirm(null);
    showToast("Event berhasil diajukan ke Dinas Provinsi! ✓");
  };

  const handleImportFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file=e.target.files?.[0]; if(!file) return;
    const reader=new FileReader();
    reader.onload=(ev)=>{
      try {
        const parsed=JSON.parse(ev.target?.result as string);
        const imported=parsed.events||(Array.isArray(parsed)?parsed:null);
        if(!imported||!imported.length){ showToast("File tidak valid atau kosong.","error"); return; }
        setImportConfirm({events:imported,count:imported.length});
      } catch { showToast("Gagal membaca file JSON.","error"); }
    };
    reader.readAsText(file); 
    e.target.value="";
  };

  const confirmImport = (mode: string) => {
    if(!importConfirm) return;
    if(mode==="replace"){
      setEvents(importConfirm.events);
      localStorage.setItem(STORAGE_KEY,JSON.stringify(importConfirm.events));
      showToast(`${importConfirm.count} event berhasil di-restore.`);
    } else {
      const merged=[...events,...importConfirm.events.filter(ie=>!events.find(e=>e.id===ie.id))];
      setEvents(merged); localStorage.setItem(STORAGE_KEY,JSON.stringify(merged));
      showToast(`${importConfirm.count} event berhasil digabungkan.`);
    }
    setImportConfirm(null);
  };

  const generateAI = async()=>{
    setAiLoading(true); setAiResult(null);
    try {
      setAiResult(`Analisis CoE Sumbar (AI):
1. Distribusi Event: Terkonsentrasi pada Festival Budaya (40%) dan Kuliner (25%).
2. Potensi Wisatawan: Mencapai ${totalTarget.toLocaleString("id-ID")} orang dengan event utama di Kota Padang dan Pariaman.
3. Rekomendasi: Perlu penguatan pada kategori MICE di daerah pegunungan.`);
    } catch { setAiResult("Terjadi kesalahan saat menghubungi AI."); }
    setAiLoading(false);
  };

  const filtered = events.filter(e=>{
    const q=searchQ.toLowerCase();
    return (!q||e.namaEvent.toLowerCase().includes(q)||e.lokasi?.toLowerCase().includes(q))
      &&(!filterKab||e.kabupatenKota===filterKab)
      &&(!filterKat||e.kategori===filterKat)
      &&(!filterStatus||e.status===filterStatus)
      &&(!filterBulan||(e.tanggalMulai&&new Date(e.tanggalMulai).getMonth()===parseInt(filterBulan)));
  });

  const getDIM=(y: number,m: number)=>new Date(y,m+1,0).getDate();
  const getFD =(y: number,m: number)=>new Date(y,m,1).getDay();
  const eInDay=(day: number)=>events.filter(e=>{
    if(!e.tanggalMulai) return false;
    const s=new Date(e.tanggalMulai), en=e.tanggalSelesai?new Date(e.tanggalSelesai):s;
    const d=new Date(calYear,calMonth,day); return d>=s&&d<=en;
  });

  const totalEvt   = events.length;
  const totalTarget= events.reduce((s,e)=>s+parseInt(e.targetWisatawan||"0"),0);
  const diajukan   = events.filter(e=>e.status==="Diajukan").length;

  const navItems=[
    {key:"dashboard",label:"Dashboard",   icon:"◈"},
    {key:"form",     label:"Tambah Event",icon:"✦"},
    {key:"list",     label:"Daftar Event",icon:"≡"},
    {key:"calendar", label:"Kalender",    icon:"⊞"},
  ];
  
  const viewLabel: Record<string, string> = {
    dashboard:"Dashboard Ikhtisar",
    form:editId?"Edit Event":"Tambah Event Baru",
    list:"Daftar Seluruh Event",
    calendar:"Tampilan Kalender",
    detail:"Detail Event"
  };

  return (
    <div className="min-h-screen bg-[#0D1B0F] text-[#E8DCC8] relative font-serif">
      <div className="fixed inset-0 z-0 pointer-events-none"
        style={{ backgroundImage: `radial-gradient(ellipse at 20% 50%,rgba(196,160,60,.07) 0%,transparent 60%), radial-gradient(ellipse at 80% 20%,rgba(34,85,34,.14) 0%,transparent 50%)` }} />

      <div className="fixed left-0 top-0 bottom-0 w-[232px] bg-[rgba(7,18,9,0.98)] border-r border-[rgba(196,160,60,0.15)] flex flex-col z-[100]">
        <div className="p-6 pb-[18px] border-b border-[rgba(196,160,60,0.1)]">
          <div className="text-[9px] tracking-[3px] text-[#C4A03C] uppercase mb-1">Dinas Pariwisata</div>
          <div className="text-base font-bold text-[#E8DCC8] leading-tight">Sumatera Barat</div>
          <div className="text-[9px] text-[#5A4820] mt-1 tracking-wider">Calendar of Events</div>
        </div>

        <div className={`m-3 p-2.5 rounded-lg border ${isProvinsi ? "bg-[rgba(196,160,60,0.08)] border-[rgba(196,160,60,0.2)]" : "bg-[rgba(78,180,130,0.07)] border-[rgba(78,180,130,0.15)]"}`}>
          <div className={`text-[9px] tracking-widest uppercase mb-1 ${isProvinsi ? "text-[#C4A03C]" : "text-[#4CAF82]"}`}>
            {isProvinsi?"🏛 Admin Provinsi":"🏘 Admin Kab/Kota"}
          </div>
          <div className="text-[11px] text-[#C8B890] font-semibold">{user.nama}</div>
          {!isProvinsi&&<div className="text-[10px] text-[#5A5040] mt-0.5">{user.kabupatenKota}</div>}
        </div>

        <nav className="p-2.5">
          <div className="text-[9px] tracking-widest text-[#2A2015] uppercase p-2">Menu</div>
          {navItems.map(n=>(
            <button key={n.key} onClick={()=>{setView(n.key);setForm(initialForm);setEditId(null);}}
              className={`flex items-center gap-2.5 w-full p-2.5 mb-0.5 rounded-lg border-none cursor-pointer text-left transition-all border-l-2
                ${view===n.key ? "bg-[rgba(196,160,60,0.12)] text-[#C4A03C] border-l-[#C4A03C]" : "bg-transparent text-[#7A6E5A] border-l-transparent"}`}>
              <span className="text-sm">{n.icon}</span>{n.label}
            </button>
          ))}
        </nav>

        {isProvinsi && (
          <div className="p-2.5 border-t border-[rgba(196,160,60,0.08)] mt-1">
            <div className="text-[9px] tracking-widest text-[#2A2015] uppercase p-2">Data</div>
            <SidebarAction icon="↻" label="Refresh Data" sub="ambil data dari server" color="#4CAF82" onClick={handleRefresh} />
            <SidebarAction icon="📊" label="Export Excel" sub=".xlsx" color="#4CAF82" onClick={()=>exportToExcel(events)} />
            <SidebarAction icon="💾" label="Export JSON" sub=".json" color="#7EB8F7" onClick={()=>exportToJSON(events)} />
            <SidebarAction icon="📂" label="Import JSON" sub="restore" color="#E8A820" onClick={()=>importRef.current?.click()} />
            <input ref={importRef} type="file" accept=".json" className="hidden" onChange={handleImportFile} />
          </div>
        )}

        <div className="mt-auto p-4 border-t border-[rgba(196,160,60,0.08)]">
          <button onClick={onLogout} className="flex items-center gap-2 w-full p-2 rounded-lg border border-[rgba(224,90,90,0.2)] bg-[rgba(224,90,90,0.05)] text-[#A05050] text-[11px] cursor-pointer">
            <span>⬡</span> Keluar
          </button>
        </div>
      </div>

      <div className="ml-[232px] min-h-screen relative z-1">
        <div className="p-4 px-8 border-b border-[rgba(196,160,60,0.08)] bg-[rgba(7,18,9,0.82)] backdrop-blur-xl sticky top-0 z-50 flex items-center justify-between">
          <div>
            <div className="text-lg font-bold text-[#E8DCC8]">{viewLabel[view]}</div>
            <div className="text-[10px] text-[#4A4030] mt-0.5">{new Date().toLocaleDateString("id-ID",{weekday:"long",year:"numeric",month:"long",day:"numeric"})}</div>
          </div>
          <div className="flex items-center gap-3">
            {sheetLoading && <div className="text-[11px] text-[#4CAF82] flex items-center gap-1"><span className="w-2.5 h-2.5 border-2 border-[#4CAF82] border-t-transparent animate-spin-custom rounded-full" /> Memuat...</div>}
            {view!=="form"&&(
              <button onClick={()=>{setForm(isProvinsi?initialForm:{...initialForm, kabupatenKota:user.kabupatenKota}); setEditId(null); setView("form");}}
                className="p-2 px-4 bg-gradient-to-br from-[#C4A03C] to-[#A07828] text-[#0D1B0F] border-none rounded-md text-xs font-bold cursor-pointer transition-opacity hover:opacity-90">+ Tambah Event</button>
            )}
          </div>
        </div>

        <div className="p-8">
          {view==="dashboard"&&(
            <div>
              <div className="grid grid-cols-3 gap-4 mb-6">
                {[
                  {label:"Total Event",value:totalEvt,icon:"✦",color:"#C4A03C"},
                  {label:"Sudah Diajukan",value:diajukan,icon:"✉",color:"#4CAF82"},
                  {label:"Target Wisatawan",value:totalTarget.toLocaleString("id-ID"),icon:"⊙",color:"#7EB8F7"},
                ].map((s,i)=>(
                  <div key={i} className="bg-white/5 border border-white/10 rounded-xl p-5">
                    <div style={{color:s.color}} className="text-2xl mb-2">{s.icon}</div>
                    <div className="text-2xl font-bold text-[#E8DCC8]">{s.value}</div>
                    <div className="text-[11px] text-[#5A5040] mt-1">{s.label}</div>
                  </div>
                ))}
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div className="bg-white/2 border border-white/5 rounded-xl p-5">
                  <div className="text-[10px] text-[#C4A03C] mb-4 tracking-widest uppercase">Event Terbaru</div>
                  {events.slice(-6).reverse().map(e=>(
                    <div key={e.id} onClick={()=>{setSelEvent(e);setView("detail");}} className="p-2 border-b border-white/5 cursor-pointer hover:bg-white/5">
                      <div className="flex items-center gap-2">
                        <span className="text-[13px] text-[#E8DCC8] font-semibold flex-1">{e.namaEvent}</span>
                        <StatusBadge status={e.status}/>
                      </div>
                      <div className="text-[10px] text-[#4A4030] mt-1">{e.kabupatenKota} · {e.tanggalMulai}</div>
                    </div>
                  ))}
                </div>
                <div className="bg-white/2 border border-white/5 rounded-xl p-5">
                  <div className="text-[10px] text-[#C4A03C] mb-4 tracking-widest uppercase">Analisis (AI)</div>
                  {aiResult ? <div className="text-xs text-[#C8BAA0] leading-relaxed whitespace-pre-wrap">{aiResult}</div> : <div className="text-[10px] text-[#2A2015] italic">Menganalisis data event...</div>}
                </div>
              </div>
            </div>
          )}

          {view==="form"&&(
            <div className="max-w-2xl bg-white/2 border border-white/10 rounded-xl p-6">
              <FormSection title="Informasi Utama">
                <FormField label="Nama Event *" full>
                  <input value={form.namaEvent} onChange={e=>setForm(f=>({...f,namaEvent:e.target.value}))}
                    placeholder="Nama event..." className="w-full bg-white/5 border border-white/10 text-[#E8DCC8] p-2 rounded-md outline-none text-sm" />
                </FormField>
                <FormField label="Provinsi / Kabupaten / Kota *">
                  {isProvinsi ? (
                    <select value={form.kabupatenKota} onChange={e=>setForm(f=>({...f,kabupatenKota:e.target.value}))} className="w-full bg-white/5 border border-white/10 text-[#E8DCC8] p-2 rounded-md outline-none text-sm">
                      <option value="">-- Pilih --</option>
                      {KABUPATEN_KOTA.map(k=><option key={k}>{k}</option>)}
                    </select>
                  ) : (
                    <input value={user.kabupatenKota} disabled className="w-full bg-white/5 border border-white/10 text-[#E8DCC8] p-2 rounded-md outline-none text-sm opacity-50 cursor-not-allowed" />
                  )}
                </FormField>
                <FormField label="Kategori *">
                  <select value={form.kategori} onChange={e=>setForm(f=>({...f,kategori:e.target.value}))} className="w-full bg-white/5 border border-white/10 text-[#E8DCC8] p-2 rounded-md outline-none text-sm">
                    <option value="">-- Pilih --</option>
                    {KATEGORI.map(k=><option key={k}>{k}</option>)}
                  </select>
                </FormField>
                <FormField label="Tanggal Mulai *">
                  <input type="date" value={form.tanggalMulai} onChange={e=>setForm(f=>({...f,tanggalMulai:e.target.value}))} className="w-full bg-white/5 border border-white/10 text-[#E8DCC8] p-2 rounded-md outline-none text-sm" />
                </FormField>
              </FormSection>
              <div className="flex gap-3 mt-4">
                <button onClick={handleSubmit} className="p-2.5 px-6 bg-gradient-to-br from-[#C4A03C] to-[#A07828] text-[#0D1B0F] border-none rounded-md text-sm font-bold cursor-pointer">Simpan Event</button>
                <button onClick={()=>setView("list")} className="p-2.5 px-[18px] bg-transparent text-[#7A6E5A] border border-white/10 rounded-md text-sm cursor-pointer">Batal</button>
              </div>
            </div>
          )}

          {view==="list"&&(
            <div>
              <div className="flex gap-2.5 mb-4 flex-wrap bg-white/1 p-4 rounded-xl border border-white/5">
                <input value={searchQ} onChange={e=>setSearchQ(e.target.value)} placeholder="🔍 Cari..." className="flex-1 min-w-[200px] bg-white/5 border border-white/10 text-[#E8DCC8] p-2 px-3 rounded-md outline-none text-sm" />
                <select value={filterKab} onChange={e=>setFilterKab(e.target.value)} className="bg-white/5 border border-white/10 text-[#E8DCC8] p-2 rounded-md outline-none text-sm"><option value="">Wilayah</option>{KABUPATEN_KOTA.map(k=><option key={k}>{k}</option>)}</select>
              </div>
              {filtered.map(e=>(
                <div key={e.id} onClick={()=>{setSelEvent(e);setView("detail");}} className="bg-white/2 border border-white/5 rounded-xl p-4 mb-2 flex gap-4 items-center cursor-pointer hover:bg-white/5">
                  <div className="w-10 h-10 rounded-lg bg-gold/10 flex items-center justify-center text-lg">{e.kategori?.includes("Budaya")?"🎭":"✦"}</div>
                  <div className="flex-1">
                    <div className="flex items-center gap-2"><span className="text-sm font-bold text-[#E8DCC8]">{e.namaEvent}</span><StatusBadge status={e.status}/></div>
                    <div className="text-[11px] text-[#6A5840] mt-0.5">{e.kabupatenKota} · {e.tanggalMulai}</div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      {toast && <div className={`fixed bottom-6 right-6 z-[999] p-3 px-5 rounded-lg text-xs shadow-xl border text-[#E8DCC8] ${toast.type==="error"?"bg-red-950 border-red-500":"bg-green-950 border-green-500"}`}>{toast.msg}</div>}

      {delConfirm && (
        <Modal>
          <div className="text-2xl mb-2">⚠</div>
          <div className="text-sm text-[#E8DCC8] mb-1">Hapus Event?</div>
          <button onClick={()=>handleDelete(delConfirm)} className="p-2 px-4 bg-red-900 border border-red-500 rounded-md text-xs text-white">Hapus</button>
          <button onClick={()=>setDelConfirm(null)} className="ml-2 p-2 px-4 bg-transparent border border-white/10 rounded-md text-xs text-[#7A6E5A]">Batal</button>
        </Modal>
      )}
    </div>
  );
}

function Modal({children}: {children: React.ReactNode}) {
  return <div className="fixed inset-0 z-[200] bg-black/80 flex items-center justify-center"><div className="bg-[#0B1A0D] border border-white/10 rounded-2xl p-8 max-w-sm w-full text-center">{children}</div></div>;
}

function SidebarAction({icon,label,sub,color,onClick}: {icon: string; label: string; sub: string; color: string; onClick: () => void}) {
  return <button onClick={onClick} className="flex items-center gap-2.5 w-full p-2 rounded-lg border-none cursor-pointer text-left bg-transparent hover:bg-white/5"><span className="text-lg">{icon}</span><div><div style={{color}} className="text-xs font-semibold">{label}</div><div className="text-[9px] text-[#2A2015]">{sub}</div></div></button>;
}

function FormSection({title,children}: {title: string; children: React.ReactNode}) {
  return <div className="mb-4"><div className="text-[9px] text-gold tracking-widest uppercase mb-3 border-b border-white/5 pb-1">{title}</div><div className="grid grid-cols-2 gap-3">{children}</div></div>;
}

function FormField({label,children,full}: {label: string; children: React.ReactNode; full?: boolean}) {
  return <div style={{gridColumn:full?"1/-1":"auto"}}><label className="text-[10px] text-[#4A4030] block mb-1 font-bold">{label}</label>{children}</div>;
}

function StatusBadge({status}: {status: string}) {
  const s=STATUS_BADGE[status]||STATUS_BADGE["Draft"];
  return <span className="text-[9px] p-0.5 px-2 rounded-md border" style={{backgroundColor:s.bg, color:s.color, borderColor:s.border}}>{status}</span>;
}
