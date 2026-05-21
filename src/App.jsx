import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

const LOKASI = ["Sky House BSD","Treepark City Cikokol"];
const SHIFT  = ["Shift 1","Shift 2"];
const KAT_IN  = ["Omset Penjualan","Denda Customer"];
const KAT_OUT = ["Operasional","Gaji Karyawan","Belanja","Tagihan","Transportasi","Peralatan","Lainnya"];
const ROLES = {
  owner:{label:"Owner",color:"#c084fc",bg:"rgba(192,132,252,.15)",canEdit:true,canDelete:true,canSetting:true},
  admin:{label:"Admin",color:"#fb923c",bg:"rgba(251,146,60,.15)",canEdit:true,canDelete:true,canSetting:true},
  user: {label:"User", color:"#60a5fa",bg:"rgba(96,165,250,.15)",canEdit:false,canDelete:false,canSetting:false},
  kasir:{label:"Kasir",color:"#34d399",bg:"rgba(52,211,153,.15)",canEdit:false,canDelete:false,canSetting:false},
};
const UK="sg_u5",TK="sg_t5";
const DEF=[{id:1,username:"owner",password:"owner123",role:"owner"},{id:2,username:"admin",password:"admin123",role:"admin"},{id:3,username:"kasir",password:"kasir123",role:"kasir"}];
const rp=n=>new Intl.NumberFormat("id-ID",{style:"currency",currency:"IDR",minimumFractionDigits:0}).format(n||0);
const fmt=v=>{const r=v.replace(/\D/g,"");return r?parseInt(r).toLocaleString("id-ID"):""};
const nowStr=()=>new Date().toLocaleString("id-ID");
const G="#f5c842",BG0="#07080f",BG1="#0f1018",BG2="#181924",BG3="#20222e",BD="#252736";
const GR="#34d399",RD="#f87171",TX="#f0f0f8",TM="#8888aa",TQ="#44455a";
const IS={width:"100%",background:BG3,border:"1.5px solid "+BD,borderRadius:14,padding:"13px 16px",color:TX,fontFamily:"inherit",fontWeight:600,fontSize:14,boxSizing:"border-box",outline:"none"};
const LS={display:"block",fontSize:10,fontWeight:800,color:TM,marginBottom:8,textTransform:"uppercase",letterSpacing:1.2};
const card={background:BG2,borderRadius:20,border:"1px solid "+BD};

function Chips({opts,val,set,color}){
  return <div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom:16}}>{opts.map(o=><button key={o} onClick={()=>set(o)} style={{padding:"8px 16px",borderRadius:30,cursor:"pointer",fontWeight:700,fontSize:12,border:val===o?"none":"1.5px solid "+BD,background:val===o?color:"transparent",color:val===o?"#000":TM,boxShadow:val===o?"0 2px 12px "+color+"55":"none"}}>{o}</button>)}</div>;
}

function Bd({children,onClose,bottom}){
  return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.85)",display:"flex",alignItems:bottom?"flex-end":"center",justifyContent:"center",zIndex:999,backdropFilter:"blur(6px)"}} onClick={onClose}><div onClick={e=>e.stopPropagation()} style={{width:"100%",maxWidth:430}}>{children}</div></div>;
}

function TxRow({t,role,onDel,onEdit}){
  const [open,setOpen]=useState(false);
  const p=ROLES[role]||ROLES.kasir;
  const inc=t.type==="income";
  return <div style={{borderBottom:"1px solid "+BD}}>
    <div onClick={()=>setOpen(o=>!o)} style={{display:"flex",alignItems:"center",gap:12,padding:"13px 0",cursor:"pointer"}}>
      <div style={{width:44,height:44,borderRadius:14,flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,fontWeight:900,background:inc?"rgba(52,211,153,.1)":"rgba(248,113,113,.1)",color:inc?GR:RD,border:"1px solid "+(inc?"rgba(52,211,153,.2)":"rgba(248,113,113,.2)")}}>
        {inc?"+":"-"}
      </div>
      <div style={{flex:1,minWidth:0}}>
        <div style={{fontWeight:800,fontSize:13,color:TX}}>{t.category}{t.lokasi?" | "+t.lokasi:""}</div>
        <div style={{fontSize:11,color:TQ,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",marginTop:2}}>{t.shift?t.shift+" | ":""}{t.note||t.datetime}</div>
      </div>
      <div style={{textAlign:"right",flexShrink:0,marginRight:6}}>
        <div style={{fontWeight:900,fontSize:13,color:inc?GR:RD}}>{inc?"+":"-"}{rp(t.amount)}</div>
        <div style={{fontSize:10,color:TQ,marginTop:1}}>{(t.datetime||"").slice(0,8)}</div>
      </div>
      <span style={{color:TQ,fontSize:12}}>{open?"^":"v"}</span>
    </div>
    {open&&<div style={{background:BG0,borderRadius:14,padding:14,marginBottom:12,border:"1px solid "+BD}}>
      {t.penerima&&<div style={{marginBottom:6,fontSize:12,color:TM}}>Penerima: <b style={{color:TX}}>{t.penerima}</b></div>}
      {t.user&&<div style={{marginBottom:6,fontSize:12,color:TM}}>Dicatat: <b style={{color:TX}}>{t.user}</b></div>}
      <div style={{fontSize:12,color:TM,marginBottom:t.photo?10:0}}>Waktu: <b style={{color:TX}}>{t.datetime}</b></div>
      {t.photo&&<img src={t.photo} alt="bukti" style={{width:"100%",borderRadius:12,maxHeight:200,objectFit:"cover",marginTop:8}}/>}
      {(p.canEdit||p.canDelete)&&<div style={{display:"flex",gap:8,marginTop:12}}>
        {p.canEdit&&<button onClick={onEdit} style={{flex:1,padding:9,borderRadius:11,border:"none",background:"rgba(251,146,60,.12)",color:"#fb923c",fontWeight:700,fontSize:12,cursor:"pointer"}}>Edit Nominal</button>}
        {p.canDelete&&<button onClick={onDel} style={{flex:1,padding:9,borderRadius:11,border:"none",background:"rgba(248,113,113,.12)",color:RD,fontWeight:700,fontSize:12,cursor:"pointer"}}>Hapus</button>}
      </div>}
    </div>}
  </div>;
}

function exportXLS(txs){
  const wb=XLSX.utils.book_new();
  const ms=[...new Set(txs.map(t=>t.datetime.slice(0,7)))].sort().reverse();
  const sd=[["Bulan","Pemasukan","Pengeluaran","Saldo"]];
  ms.forEach(m=>{const mx=txs.filter(t=>t.datetime.startsWith(m));const i=mx.filter(t=>t.type==="income").reduce((s,t)=>s+t.amount,0);const e=mx.filter(t=>t.type==="expense").reduce((s,t)=>s+t.amount,0);sd.push([m,i,e,i-e]);});
  const w0=XLSX.utils.aoa_to_sheet(sd);w0["!cols"]=[{wch:14},{wch:18},{wch:18},{wch:18}];XLSX.utils.book_append_sheet(wb,w0,"Ringkasan");
  const ar=[["Waktu","Jenis","Lokasi","Shift","Kategori","Jumlah","Catatan","Penerima","User"]];
  [...txs].sort((a,b)=>b.datetime.localeCompare(a.datetime)).forEach(t=>ar.push([t.datetime,t.type==="income"?"Pemasukan":"Pengeluaran",t.lokasi||"-",t.shift||"-",t.category,t.amount,t.note||"-",t.penerima||"-",t.user||"-"]));
  const w1=XLSX.utils.aoa_to_sheet(ar);w1["!cols"]=[{wch:18},{wch:14},{wch:20},{wch:10},{wch:20},{wch:16},{wch:28},{wch:20},{wch:12}];XLSX.utils.book_append_sheet(wb,w1,"Transaksi");
  XLSX.writeFile(wb,"OmsetSayyGroup_"+new Date().toISOString().slice(0,10)+".xlsx");
}
function exportPDF(txs){
  const ti=txs.filter(t=>t.type==="income").reduce((s,t)=>s+t.amount,0);
  const to=txs.filter(t=>t.type==="expense").reduce((s,t)=>s+t.amount,0);
  const rows=[...txs].sort((a,b)=>b.datetime.localeCompare(a.datetime)).map(t=>"<tr><td>"+t.datetime+"</td><td style='color:"+(t.type==="income"?"green":"red")+"'>"+(t.type==="income"?"Pemasukan":"Pengeluaran")+"</td><td>"+t.category+"</td><td>"+rp(t.amount)+"</td><td>"+(t.note||"-")+"</td></tr>").join("");
  const w=window.open("","_blank");
  w.document.write("<html><head><title>Laporan SayyGroup</title><style>body{font-family:Arial;font-size:12px;padding:20px}h2{text-align:center;color:#b45309}table{width:100%;border-collapse:collapse}th{background:#1e293b;color:#f59e0b;padding:8px}td{padding:6px 8px;border-bottom:1px solid #eee}</style></head><body><h2>Laporan Omset SayyGroup</h2><p style='text-align:center'>Dicetak: "+nowStr()+"</p><p><b style='color:green'>Masuk: "+rp(ti)+"</b> | <b style='color:red'>Keluar: "+rp(to)+"</b> | <b>Saldo: "+rp(ti-to)+"</b></p><table><thead><tr><th>Waktu</th><th>Jenis</th><th>Kategori</th><th>Jumlah</th><th>Catatan</th></tr></thead><tbody>"+rows+"</tbody></table></body></html>");
  w.document.close();w.print();
}

export default function App(){
  const [users,setUsers]=useState([]);
  const [sess,setSess]=useState(null);
  const [lf,setLf]=useState({u:"",p:""});
  const [le,setLe]=useState("");
  const [txs,setTxs]=useState([]);
  const [view,setView]=useState("dash");
  const [type,setType]=useState("income");
  const [month,setMonth]=useState(new Date().toISOString().slice(0,7));
  const [toast,setToast]=useState(null);
  const [delId,setDelId]=useState(null);
  const [editT,setEditT]=useState(null);
  const [expM,setExpM]=useState(false);
  const [stab,setStab]=useState("members");
  const [newU,setNewU]=useState({username:"",password:"",role:"kasir"});
  const [editU,setEditU]=useState(null);
  const [delU,setDelU]=useState(null);
  const [clock,setClock]=useState(new Date());
  const [form,setForm]=useState({amt:"",lok:"",sh:"",cat:"",note:"",recv:"",photo:null});
  const fRef=useRef();

  useEffect(()=>{try{const u=localStorage.getItem(UK);setUsers(u?JSON.parse(u):DEF);const t=localStorage.getItem(TK);if(t)setTxs(JSON.parse(t));}catch{setUsers(DEF);}},[]);
  useEffect(()=>{try{localStorage.setItem(UK,JSON.stringify(users));}catch{}},[users]);
  useEffect(()=>{try{localStorage.setItem(TK,JSON.stringify(txs));}catch{}},[txs]);
  useEffect(()=>{const id=setInterval(()=>setClock(new Date()),1000);return()=>clearInterval(id);},[]);

  const toast2=(m,c=GR)=>{setToast({m,c});setTimeout(()=>setToast(null),2400);};
  const login=()=>{
    const u=users.find(x=>x.username.toLowerCase()===lf.u.toLowerCase());
    if(!u||u.password!==lf.p)return setLe("Username atau password salah");
    setSess({id:u.id,u:u.username,role:u.role});setLe("");
  };

  const role=sess?.role||"kasir";
  const perm=ROLES[role]||ROLES.kasir;
  const allIn=txs.filter(t=>t.type==="income").reduce((s,t)=>s+t.amount,0);
  const allOut=txs.filter(t=>t.type==="expense").reduce((s,t)=>s+t.amount,0);
  const saldo=allIn-allOut;
  const dk=clock.getDate().toString().padStart(2,"0")+"/"+(clock.getMonth()+1).toString().padStart(2,"0")+"/"+clock.getFullYear();
  const todayTxs=txs.filter(t=>(t.datetime||"").startsWith(dk));
  const todayIn=todayTxs.filter(t=>t.type==="income").reduce((s,t)=>s+t.amount,0);
  const todayOut=todayTxs.filter(t=>t.type==="expense").reduce((s,t)=>s+t.amount,0);
  const fil=txs.filter(t=>(t.datetime||"").includes("/"+month.slice(5)+"/"+month.slice(0,4)));
  const totI=fil.filter(t=>t.type==="income").reduce((s,t)=>s+t.amount,0);
  const totO=fil.filter(t=>t.type==="expense").reduce((s,t)=>s+t.amount,0);
  const timeStr=clock.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"});
  const dateStr=clock.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});

  const saveTx=()=>{
    const a=parseFloat(form.amt.replace(/\D/g,""));
    if(!a||a<=0)return toast2("Masukkan jumlah yang valid",RD);
    if(type==="income"&&!form.lok)return toast2("Pilih lokasi",RD);
    if(type==="income"&&!form.sh)return toast2("Pilih shift",RD);
    if(!form.cat)return toast2("Pilih kategori",RD);
    setTxs(p=>[{id:Date.now(),type,amount:a,lokasi:form.lok,shift:form.sh,category:form.cat,note:form.note,penerima:form.recv,photo:form.photo,datetime:nowStr(),user:sess.u},...p]);
    setForm({amt:"",lok:"",sh:"",cat:"",note:"",recv:"",photo:null});
    setView("dash");
    toast2(type==="income"?"Pemasukan dicatat!":"Pengeluaran dicatat!");
  };
  const delTx=id=>{setTxs(p=>p.filter(t=>t.id!==id));setDelId(null);toast2("Transaksi dihapus","#fb923c");};
  const editSave=()=>{const a=parseFloat(editT.v.replace(/\D/g,""));if(!a)return;setTxs(p=>p.map(t=>t.id===editT.id?{...t,amount:a}:t));setEditT(null);toast2("Diperbarui");};
  const doPhoto=e=>{const f=e.target.files[0];if(!f)return;const r=new FileReader();r.onload=ev=>setForm(f=>({...f,photo:ev.target.result}));r.readAsDataURL(f);};
  const addUser=()=>{if(!newU.username.trim())return toast2("Masukkan username",RD);if(newU.password.length<4)return toast2("Password min 4 karakter",RD);if(users.find(u=>u.username.toLowerCase()===newU.username.toLowerCase()))return toast2("Username sudah ada",RD);setUsers(p=>[...p,{id:Date.now(),...newU,username:newU.username.toLowerCase()}]);setNewU({username:"",password:"",role:"kasir"});setStab("members");toast2("Anggota ditambahkan!");};
  const saveEditU=()=>{if(!editU.username.trim()||editU.password.length<4)return toast2("Data tidak valid",RD);setUsers(p=>p.map(u=>u.id===editU.id?{...u,...editU}:u));setEditU(null);setStab("members");toast2("Diperbarui");};
  const deleteUser=id=>{if(sess?.id===id)return toast2("Tidak bisa hapus akun sendiri",RD);setUsers(p=>p.filter(u=>u.id!==id));setDelU(null);toast2("Anggota dihapus","#fb923c");};

  const navItems=[{id:"dash",label:"Beranda",sym:"[H]"},{id:"form",label:"Catat",sym:"[+]"},{id:"hist",label:"Riwayat",sym:"[=]"},...(perm.canSetting?[{id:"setting",label:"Setting",sym:"[S]"}]:[])];

  if(!sess) return(
    <div style={{fontFamily:"'Nunito',sans-serif",background:BG0,minHeight:"100vh",maxWidth:430,margin:"0 auto",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:24,color:TX,position:"relative",overflow:"hidden"}}>
      <link href="https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800;900&display=swap" rel="stylesheet"/>
      <div style={{position:"absolute",top:"15%",left:"50%",transform:"translateX(-50%)",width:280,height:280,background:"radial-gradient(circle,rgba(245,200,66,.1) 0%,transparent 70%)",pointerEvents:"none"}}/>
      <div style={{width:76,height:76,borderRadius:24,background:"linear-gradient(135deg,"+G+",#c9a52a)",display:"flex",alignItems:"center",justifyContent:"center",marginBottom:16,boxShadow:"0 8px 40px rgba(245,200,66,.35)",fontSize:32,fontWeight:900,color:"#000"}}>$</div>
      <div style={{fontWeight:900,fontSize:22,color:G,letterSpacing:.5,marginBottom:4}}>OMSET SAYY GROUP</div>
      <div style={{fontSize:12,color:TM,marginBottom:30}}>Sistem Pencatatan Keuangan</div>
      <div style={{...card,width:"100%",padding:26}}>
        <label style={LS}>Username</label>
        <input value={lf.u} onChange={e=>setLf(f=>({...f,u:e.target.value}))} placeholder="username..." style={{...IS,marginBottom:14}}/>
        <label style={LS}>Password</label>
        <input type="password" value={lf.p} onChange={e=>setLf(f=>({...f,p:e.target.value}))} onKeyDown={e=>e.key==="Enter"&&login()} placeholder="min 4 karakter" style={{...IS,marginBottom:8}}/>
        {le&&<div style={{color:RD,fontSize:12,marginBottom:10}}>{le}</div>}
        <button onClick={login} style={{width:"100%",padding:14,marginTop:8,borderRadius:14,border:"none",background:"linear-gradient(135deg,"+G+",#c9a52a)",color:"#000",fontWeight:900,fontSize:15,cursor:"pointer",boxShadow:"0 4px 24px rgba(245,200,66,.35)"}}>Masuk</button>
      </div>
      <div style={{marginTop:18,display:"flex",gap:8,flexWrap:"wrap",justifyContent:"center"}}>
        {users.slice(0,3).map(u=>{const p2=ROLES[u.role]||ROLES.kasir;return <span key={u.id} style={{padding:"3px 10px",borderRadius:8,background:p2.bg,color:p2.color,fontWeight:700,fontSize:11}}>[{p2.label}] {u.username}</span>;})}
      </div>
    </div>
  );

  return(
    <div style={{fontFamily:"'Nunito',sans-serif",background:BG0,minHeight:"100vh",maxWidth:430,margin:"0 auto",color:TX,paddingBottom:86}}>
      <link href="https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800;900&display=swap" rel="stylesheet"/>

      {toast&&<div style={{position:"fixed",top:18,left:"50%",transform:"translateX(-50%)",background:toast.c,color:"#000",borderRadius:30,padding:"10px 22px",fontWeight:800,zIndex:1000,fontSize:13,whiteSpace:"nowrap",boxShadow:"0 4px 24px rgba(0,0,0,.5)"}}>{toast.m}</div>}

      {delId&&<Bd onClose={()=>setDelId(null)}>
        <div style={{...card,padding:26,margin:"0 16px",textAlign:"center"}}>
          <div style={{width:52,height:52,borderRadius:16,background:"rgba(248,113,113,.15)",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 14px",fontSize:24,color:RD}}>X</div>
          <div style={{fontWeight:900,fontSize:16,marginBottom:6}}>Hapus Transaksi?</div>
          <div style={{color:TM,fontSize:13,marginBottom:20}}>Data tidak bisa dikembalikan.</div>
          <div style={{display:"flex",gap:10}}>
            <button onClick={()=>setDelId(null)} style={{flex:1,padding:12,borderRadius:12,border:"1px solid "+BD,background:"transparent",color:TM,fontWeight:700,cursor:"pointer"}}>Batal</button>
            <button onClick={()=>delTx(delId)} style={{flex:1,padding:12,borderRadius:12,border:"none",background:"linear-gradient(135deg,#ef4444,#dc2626)",color:"#fff",fontWeight:800,cursor:"pointer"}}>Hapus</button>
          </div>
        </div>
      </Bd>}

      {editT&&<Bd onClose={()=>setEditT(null)}>
        <div style={{...card,padding:26,margin:"0 16px"}}>
          <div style={{fontWeight:900,fontSize:16,marginBottom:14}}>Edit Nominal</div>
          <input value={editT.v} inputMode="numeric" onChange={e=>setEditT(t=>({...t,v:fmt(e.target.value)}))} style={{...IS,marginBottom:16,fontSize:22,fontWeight:900,textAlign:"center"}}/>
          <div style={{display:"flex",gap:10}}>
            <button onClick={()=>setEditT(null)} style={{flex:1,padding:12,borderRadius:12,border:"1px solid "+BD,background:"transparent",color:TM,fontWeight:700,cursor:"pointer"}}>Batal</button>
            <button onClick={editSave} style={{flex:1,padding:12,borderRadius:12,border:"none",background:"linear-gradient(135deg,#fb923c,#ea580c)",color:"#fff",fontWeight:800,cursor:"pointer"}}>Simpan</button>
          </div>
        </div>
      </Bd>}

      {expM&&<Bd bottom onClose={()=>setExpM(false)}>
        <div style={{...card,borderRadius:"22px 22px 0 0",padding:26}}>
          <div style={{fontWeight:900,fontSize:16,marginBottom:20,textAlign:"center",color:G}}>Export Laporan</div>
          <button onClick={()=>{exportXLS(txs);setExpM(false);toast2("File Excel diunduh!");}} style={{width:"100%",padding:14,marginBottom:10,borderRadius:14,border:"none",background:"linear-gradient(135deg,#22c55e,#15803d)",color:"#fff",fontWeight:800,fontSize:14,cursor:"pointer"}}>Export Excel (.xlsx)</button>
          <button onClick={()=>{exportPDF(txs);setExpM(false);}} style={{width:"100%",padding:14,marginBottom:14,borderRadius:14,border:"none",background:"linear-gradient(135deg,#ef4444,#b91c1c)",color:"#fff",fontWeight:800,fontSize:14,cursor:"pointer"}}>Export PDF / Print</button>
          <button onClick={()=>setExpM(false)} style={{width:"100%",padding:12,borderRadius:14,border:"1px solid "+BD,background:"transparent",color:TM,fontWeight:700,cursor:"pointer"}}>Batal</button>
        </div>
      </Bd>}

      {delU&&<Bd onClose={()=>setDelU(null)}>
        <div style={{...card,padding:26,margin:"0 16px",textAlign:"center"}}>
          <div style={{fontWeight:900,fontSize:16,marginBottom:6}}>Hapus Anggota?</div>
          <div style={{color:"#fb923c",fontSize:14,fontWeight:800,marginBottom:4}}>{delU.username}</div>
          <div style={{color:TM,fontSize:12,marginBottom:20}}>Akun ini tidak bisa login lagi.</div>
          <div style={{display:"flex",gap:10}}>
            <button onClick={()=>setDelU(null)} style={{flex:1,padding:12,borderRadius:12,border:"1px solid "+BD,background:"transparent",color:TM,fontWeight:700,cursor:"pointer"}}>Batal</button>
            <button onClick={()=>deleteUser(delU.id)} style={{flex:1,padding:12,borderRadius:12,border:"none",background:"linear-gradient(135deg,#ef4444,#dc2626)",color:"#fff",fontWeight:800,cursor:"pointer"}}>Hapus</button>
          </div>
        </div>
      </Bd>}

      {editU&&<Bd onClose={()=>setEditU(null)}>
        <div style={{...card,padding:24,margin:"0 16px"}}>
          <div style={{fontWeight:900,fontSize:16,marginBottom:16}}>Edit Anggota</div>
          <label style={LS}>Username</label>
          <input value={editU.username} onChange={e=>setEditU(u=>({...u,username:e.target.value}))} style={{...IS,marginBottom:12}}/>
          <label style={LS}>Password</label>
          <input type="password" value={editU.password} onChange={e=>setEditU(u=>({...u,password:e.target.value}))} style={{...IS,marginBottom:12}}/>
          <label style={LS}>Role</label>
          <Chips opts={Object.keys(ROLES)} val={editU.role} set={v=>setEditU(u=>({...u,role:v}))} color={ROLES[editU.role]?.color||G}/>
          <div style={{display:"flex",gap:10,marginTop:4}}>
            <button onClick