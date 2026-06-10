import React, { useState, useEffect, useMemo, useRef } from 'react';
import { 
  LayoutDashboard, Users, UserSquare, FileSignature, 
  Wallet, Receipt, ArrowRightLeft, CreditCard, 
  Banknote, TrendingUp, TrendingDown, FileText, 
  Plus, Trash2, Menu, X, ChevronDown, ChevronUp, Edit,
  UserCog, LogOut, KeyRound, MessageCircle, Upload, Landmark, BookOpen
} from 'lucide-react';

// --- MOCK DATA UNTUK PREVIEW JIKA URL API BELUM DIISI ---
const initialUsers = [{ id: 1, username: 'admin', password: '123', nama: 'Admin Utama', role: 'Super Admin' }, { id: 2, username: 'bendahara', password: '123', nama: 'Ibu Bendahara', role: 'Bendahara' }];
const initialSiswa = [{ id: 1, nis: '2023001', nisn: '0051234567', password: '123', nama: 'Budi Santoso', kelas: '10-A', wali: 'Agus', noWa: '6281234567890', tanggalLahir: '2008-01-15' }];
const initialTanggungan = [{ id: 1, nis: '2023001', jenis: 'SPP Juli 2026', nilai_tagihan: 250000, telah_dibayar: 0, sisa_tagihan: 250000, status: 'Belum Lunas' }];
const initialKategoriPendapatan = [{ id: 1, nama: 'Dana BOS' }];
const initialKategoriPengeluaran = [{ id: 1, nama: 'Listrik & Air' }];
const initialAkunKas = [{ id: 1, kode: '1-1001', nama: 'Kas Tunai / Brankas' }];
const initialTransaksi = [];

export default function App() {
  // ==========================================
  // 1. KONFIGURASI API DATABASE
  // ==========================================
  const API_URL = "https://script.google.com/macros/s/AKfycbxeux66KJPE2kATIt49eDZSF8UeArFbNjJzniaqhAds7T1xVs2cVHVmFq3ZAuYxzISO/exec";

  // ==========================================
  // 2. STATE MANAGEMENT
  // ==========================================
  const [isLoading, setIsLoading] = useState(true);
  const [currentUser, setCurrentUser] = useState(null); 
  const [activeMenu, setActiveMenu] = useState('dashboard');
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);

  // Data States
  const [users, setUsers] = useState([]);
  const [siswa, setSiswa] = useState([]);
  const [tanggungan, setTanggungan] = useState([]);
  const [kategoriPendapatan, setKategoriPendapatan] = useState([]);
  const [kategoriPengeluaran, setKategoriPengeluaran] = useState([]);
  const [akunKas, setAkunKas] = useState([]);
  const [transaksi, setTransaksi] = useState([]);

  // UI States
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [modalConfig, setModalConfig] = useState(null);
  const [expandedRow, setExpandedRow] = useState(null);
  const [searchTagihan, setSearchTagihan] = useState('');
  const [mutasiAkunId, setMutasiAkunId] = useState('');

  const fileInputSiswaRef = useRef(null);
  const fileInputAkunSiswaRef = useRef(null);

  // ==========================================
  // 3. FUNGSI SINKRONISASI DATABASE (API CALL)
  // ==========================================
  const syncDatabase = async (sheetName, method, id = null, dataArray = null) => {
    if (!API_URL.startsWith('http')) {
      console.warn(`Simulasi API Lokal (${method}):`, sheetName, dataArray);
      return;
    }
    try {
      await fetch(API_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ sheet: sheetName, method: method, id: id, data: dataArray })
      });
    } catch (err) {
      console.error(`Gagal sinkronisasi ${method} ke tabel ${sheetName}:`, err);
    }
  };

  // ==========================================
  // 4. MEMUAT DATA AWAL (READ)
  // ==========================================
  useEffect(() => {
    const fetchInitialData = async () => {
      if (!API_URL.startsWith('http')) {
        setUsers(initialUsers);
        setSiswa(initialSiswa);
        setTanggungan(initialTanggungan);
        setKategoriPendapatan(initialKategoriPendapatan);
        setKategoriPengeluaran(initialKategoriPengeluaran);
        setAkunKas(initialAkunKas);
        setTransaksi(initialTransaksi);
        setIsLoading(false);
        return;
      }
      try {
        const response = await fetch(`${API_URL}?action=all_data`);
        const result = await response.json();
        
        setUsers(result.users || []);
        setSiswa(result.siswa || []);
        setTanggungan(result.tanggungan || []);
        setKategoriPendapatan(result.kategoriPendapatan || []);
        setKategoriPengeluaran(result.kategoriPengeluaran || []);
        setAkunKas(result.akunKas || []);
        setTransaksi(result.transaksi || []);
        
        setIsLoading(false);
      } catch (error) {
        console.error("Gagal menarik data:", error);
        alert("Gagal terhubung ke database. Pastikan URL API benar dan koneksi internet stabil.");
        setIsLoading(false);
      }
    };
    fetchInitialData();
  }, []);

  // ==========================================
  // 5. KALKULASI GLOBAL
  // ==========================================
  const totalKasMasuk = useMemo(() => transaksi.filter(t => t.tipe && t.tipe.includes('MASUK')).reduce((sum, t) => sum + Number(t.nominal), 0), [transaksi]);
  const totalKasKeluar = useMemo(() => transaksi.filter(t => t.tipe === 'KELUAR').reduce((sum, t) => sum + Number(t.nominal), 0), [transaksi]);
  const saldoAkhir = totalKasMasuk - totalKasKeluar;

  const formatRupiah = (angka) => new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', minimumFractionDigits: 0 }).format(angka || 0);

  // ==========================================
  // 6. HANDLER IMPORT EXCEL
  // ==========================================
  const handleImportSiswa = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
      const XLSX = await import('https://cdn.sheetjs.com/xlsx-0.20.0/package/xlsx.mjs');
      const reader = new FileReader();
      reader.onload = async (event) => {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        const newSiswa = [];
        for (let i = 1; i < jsonData.length; i++) { 
          const row = jsonData[i];
          if (row && row.length >= 4 && row[0]) {
            newSiswa.push({
              id: Date.now() + i,
              nis: String(row[0] || '').trim(),
              nisn: String(row[1] || '-').trim(),
              password: String(row[2] || '123').trim(),
              nama: String(row[3] || '').trim(),
              kelas: String(row[4] || '-').trim(),
              wali: String(row[5] || '-').trim(),
              noWa: String(row[6] || '').trim(),
              tanggalLahir: String(row[7] || '').trim()
            });
          }
        }
        
        if (newSiswa.length > 0) {
            setSiswa(prev => [...prev, ...newSiswa]);
            alert(`${newSiswa.length} data siswa diimpor. Sedang menyimpan ke database (Background)...`);
            
            for(const s of newSiswa) {
              await syncDatabase('Siswa', 'POST', null, [s.id, s.nis, s.nisn, s.password, s.nama, s.kelas, s.wali, s.noWa, s.tanggalLahir]);
            }
        } else {
            alert('Gagal mengimpor. Format harus: [NIS, NISN, Password, Nama, Kelas, Wali, No WA, Tgl Lahir]');
        }
      };
      reader.readAsArrayBuffer(file);
    } catch (err) { alert('Gagal memproses Excel.'); }
    e.target.value = null; 
  };

  const handleImportAkunSiswa = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
      const XLSX = await import('https://cdn.sheetjs.com/xlsx-0.20.0/package/xlsx.mjs');
      const reader = new FileReader();
      reader.onload = async (event) => {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        let updatedCount = 0;
        let addedCount = 0;
        const newSiswaList = [...siswa];

        for (let i = 1; i < jsonData.length; i++) { 
          const row = jsonData[i];
          if (row && row.length >= 2 && row[0]) {
            const nisToFind = String(row[0] || '').trim();
            const newPassword = String(row[1] || '').trim() || '123';
            const newNama = String(row[2] || '').trim() || 'Siswa Baru'; 

            const existingIndex = newSiswaList.findIndex(s => s.nis === nisToFind);
            if (existingIndex !== -1) {
              newSiswaList[existingIndex].password = newPassword;
              const s = newSiswaList[existingIndex];
              await syncDatabase('Siswa', 'PUT', s.id, [s.id, s.nis, s.nisn, newPassword, s.nama, s.kelas, s.wali, s.noWa, s.tanggalLahir]);
              updatedCount++;
            } else {
              const newId = Date.now() + i;
              newSiswaList.push({ id: newId, nis: nisToFind, nisn: '-', password: newPassword, nama: newNama, kelas: '-', wali: '-', noWa: '', tanggalLahir: '' });
              await syncDatabase('Siswa', 'POST', null, [newId, nisToFind, '-', newPassword, newNama, '-', '-', '', '']);
              addedCount++;
            }
          }
        }
        
        if (updatedCount > 0 || addedCount > 0) {
            setSiswa(newSiswaList);
            alert(`Berhasil! ${updatedCount} akun diupdate, ${addedCount} akun ditambahkan.`);
        } else {
            alert('Format Excel salah. Pastikan kolom: [NIS, Password, Nama]');
        }
      };
      reader.readAsArrayBuffer(file);
    } catch (error) { alert('Gagal memuat library Excel.'); }
    e.target.value = null;
  };

  const handleKirimWA = (studentData) => {
    if (!studentData.noWa) return alert('Nomor WA ortu belum didaftarkan.');
    if (studentData.totalKekurangan <= 0) return alert('Tidak ada tunggakan.');
    const rincian = studentData.studentBills.filter(b => b.sisa_tagihan > 0).map(b => `- ${b.jenis}: ${formatRupiah(b.sisa_tagihan)}`).join('\n');
    const msg = `Halo Bapak/Ibu ${studentData.wali}, wali dari siswa *${studentData.nama}*.\nBerikut rincian tunggakan sekolah:\n${rincian}\n*Total Kekurangan: ${formatRupiah(studentData.totalKekurangan)}*\nMohon segera diselesaikan. Terima kasih.`;
    window.open(`https://wa.me/${studentData.noWa}?text=${encodeURIComponent(msg)}`, '_blank');
  };

  // ==========================================
  // 7. AUTENTIKASI (LOGIN / LOGOUT)
  // ==========================================
  const handleLogin = (e) => {
    e.preventDefault();
    const formData = new FormData(e.target);
    const uname = formData.get('username');
    const pass = formData.get('password');

    const staf = users.find(u => String(u.username) === String(uname) && String(u.password) === String(pass));
    if (staf) { setCurrentUser(staf); setActiveMenu('dashboard'); return; }

    const dataSiswa = siswa.find(s => String(s.nis) === String(uname) && String(s.password) === String(pass));
    if (dataSiswa) { setCurrentUser({ ...dataSiswa, username: dataSiswa.nis, role: 'Siswa' }); setActiveMenu('tagihan_pribadi'); return; }

    alert('Username/NIS atau Password salah!');
  };

  const menuGroups = (() => {
    if (!currentUser) return [];
    const role = currentUser.role;
    const g = [];
    if (['Super Admin', 'Kepala Sekolah', 'Bendahara'].includes(role)) g.push({ title: 'Pusat Kendali', items: [{ id: 'dashboard', label: 'Dashboard', icon: <LayoutDashboard size={20} /> }] });
    if (['Super Admin', 'Bendahara'].includes(role)) {
      g.push({ title: 'Master Data', items: [ { id: 'akun_kas', label: 'Akun Kas', icon: <Landmark size={20} /> }, { id: 'siswa', label: 'Data Siswa', icon: <Users size={20} /> }, { id: 'tanggungan', label: 'Tagihan Siswa', icon: <FileSignature size={20} /> }, { id: 'kat_pendapatan', label: 'Pendapatan', icon: <TrendingUp size={20} /> }, { id: 'kat_pengeluaran', label: 'Pengeluaran', icon: <TrendingDown size={20} /> } ] });
      g.push({ title: 'Transaksi', items: [ { id: 'bayar_siswa', label: 'Bayar Siswa', icon: <Receipt size={20} /> }, { id: 'pendapatan_lain', label: 'Masuk Lain', icon: <Wallet size={20} /> }, { id: 'pengeluaran', label: 'Pengeluaran', icon: <CreditCard size={20} /> } ] });
    }
    if (['Super Admin', 'Kepala Sekolah'].includes(role)) g.push({ title: 'Reporting', items: [ { id: 'mutasi_akun', label: 'Buku Besar', icon: <BookOpen size={20} /> }, { id: 'lap_arus_kas', label: 'Arus Kas', icon: <ArrowRightLeft size={20} /> }, { id: 'lap_laba_rugi', label: 'Laba/Rugi', icon: <FileText size={20} /> } ] });
    if (role === 'Super Admin') g.push({ title: 'Sistem', items: [ { id: 'manajemen_user', label: 'Staf Admin', icon: <UserCog size={20} /> }, { id: 'akun_siswa', label: 'Akun Siswa', icon: <KeyRound size={20} /> } ] });
    if (role === 'Siswa') g.push({ title: 'Area Siswa', items: [ { id: 'tagihan_pribadi', label: 'Riwayat Tagihan', icon: <FileSignature size={20} /> } ] });
    return g;
  })();

  // ==========================================
  // 8. KOMPONEN UI GLOBAL
  // ==========================================
  const handleOpenModal = (title, fields, onSubmit, initialData = null) => {
    setModalConfig({ title, fields, onSubmit, initialData });
    setIsModalOpen(true);
  };

  const DynamicModal = () => {
    if (!isModalOpen || !modalConfig) return null;
    return (
      <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
        <div className="bg-white rounded-xl shadow-xl w-full max-w-md p-6 max-h-[90vh] overflow-y-auto">
          <div className="flex justify-between items-center mb-4"><h3 className="text-lg font-bold">{modalConfig.title}</h3><button onClick={() => setIsModalOpen(false)}><X size={20}/></button></div>
          <form onSubmit={(e) => { e.preventDefault(); const d = Object.fromEntries(new FormData(e.target)); modalConfig.onSubmit(d); setIsModalOpen(false); }} className="space-y-4">
            {modalConfig.fields.map((f, i) => (
              <div key={i}><label className="block text-sm font-medium mb-1">{f.label}</label>
                {f.type === 'select' ? 
                  <select name={f.name} required defaultValue={modalConfig.initialData?.[f.name] || ''} className="w-full p-2 border rounded-lg focus:ring-2 focus:ring-blue-500">
                    <option value="">Pilih...</option>{f.options.map((o, j) => <option key={j} value={o.value || o}>{o.label || o}</option>)}
                  </select> : 
                  <input type={f.type || 'text'} name={f.name} required={f.required !== false} defaultValue={modalConfig.initialData?.[f.name] || ''} className="w-full p-2 border rounded-lg focus:ring-2 focus:ring-blue-500" />
                }
              </div>
            ))}
            <div className="pt-4 flex justify-end gap-2"><button type="button" onClick={() => setIsModalOpen(false)} className="px-4 py-2 bg-gray-100 rounded-lg">Batal</button><button type="submit" className="px-4 py-2 bg-blue-600 text-white rounded-lg">Simpan</button></div>
          </form>
        </div>
      </div>
    );
  };

  const DataTable = ({ title, columns, data, onAdd, onEdit, onDelete, customHeader }) => (
    <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
      <div className="p-4 border-b flex justify-between items-center bg-gray-50/50">
        <h2 className="text-lg font-semibold flex items-center gap-2"><FileText size={20} className="text-blue-600"/> {title}</h2>
        <div className="flex gap-2">{customHeader} {onAdd && <button onClick={onAdd} className="flex gap-2 bg-blue-600 text-white px-4 py-2 rounded-lg text-sm"><Plus size={16} /> Tambah Data</button>}</div>
      </div>
      <div className="overflow-x-auto">
        <table className="w-full text-left text-sm text-gray-600"><thead className="bg-gray-50 text-xs uppercase"><tr>{columns.map(c => <th key={c.key} className="px-6 py-4">{c.label}</th>)}{(onEdit||onDelete) && <th className="px-6 py-4 text-right">Aksi</th>}</tr></thead>
          <tbody className="divide-y divide-gray-100">
            {data.length === 0 ? <tr><td colSpan="100%" className="px-6 py-8 text-center text-gray-400">Belum ada data</td></tr> : 
            data.map((row, i) => (
              <tr key={i} className="hover:bg-blue-50/30">
                {columns.map(c => <td key={c.key} className="px-6 py-4">{c.render ? c.render(row[c.key], row) : row[c.key]}</td>)}
                {(onEdit||onDelete) && <td className="px-6 py-4 text-right flex justify-end gap-2">
                  {onEdit && <button onClick={()=>onEdit(row)} className="text-blue-500 p-1 bg-blue-50 rounded"><Edit size={16}/></button>}
                  {onDelete && <button onClick={()=>{ if(window.confirm('Hapus data ini?')) onDelete(row.id) }} className="text-red-500 p-1 bg-red-50 rounded"><Trash2 size={16}/></button>}
                </td>}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );

  // ==========================================
  // 9. LAYAR UTAMA (RENDER CONTENT)
  // ==========================================
  if (isLoading) {
    return <div className="min-h-screen bg-slate-900 flex flex-col items-center justify-center"><div className="w-12 h-12 border-4 border-blue-500 border-t-transparent rounded-full animate-spin mb-4"></div><p className="text-white text-lg font-medium">Menghubungkan ke Database...</p></div>;
  }

  if (!currentUser) {
    return (
      <div className="min-h-screen bg-slate-900 flex items-center justify-center p-4">
        <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md overflow-hidden">
          <div className="bg-blue-600 p-8 text-center"><Banknote className="text-white mx-auto mb-4" size={32} /><h1 className="text-2xl font-bold text-white">EduFinance</h1></div>
          <div className="p-8"><form onSubmit={handleLogin} className="space-y-5">
            <div><label className="block text-sm font-medium mb-1">Username / NIS</label><input type="text" name="username" required className="w-full p-2.5 border rounded-lg" /></div>
            <div><label className="block text-sm font-medium mb-1">Password</label><input type="password" name="password" required className="w-full p-2.5 border rounded-lg" /></div>
            <button type="submit" className="w-full bg-blue-600 text-white font-bold py-2.5 rounded-lg">Masuk Sistem</button>
          </form></div>
        </div>
      </div>
    );
  }

  const renderContent = () => {
    switch (activeMenu) {
      case 'dashboard': return (
        <div className="space-y-6">
          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            <div className="bg-white p-6 rounded-xl flex gap-4"><div className="p-4 bg-green-100 text-green-600 rounded-lg"><TrendingUp size={28}/></div><div><p className="text-sm text-gray-500">Kas Masuk</p><h3 className="text-2xl font-bold">{formatRupiah(totalKasMasuk)}</h3></div></div>
            <div className="bg-white p-6 rounded-xl flex gap-4"><div className="p-4 bg-red-100 text-red-600 rounded-lg"><TrendingDown size={28}/></div><div><p className="text-sm text-gray-500">Kas Keluar</p><h3 className="text-2xl font-bold">{formatRupiah(totalKasKeluar)}</h3></div></div>
            <div className="bg-white p-6 rounded-xl flex gap-4"><div className="p-4 bg-blue-100 text-blue-600 rounded-lg"><Wallet size={28}/></div><div><p className="text-sm text-gray-500">Saldo Akhir</p><h3 className="text-2xl font-bold text-blue-600">{formatRupiah(saldoAkhir)}</h3></div></div>
          </div>
        </div>
      );

      case 'manajemen_user': return <DataTable title="Manajemen Staf" columns={[{key:'username', label:'Username'}, {key:'nama', label:'Nama Lengkap'}, {key:'role', label:'Role'}]} data={users} 
        onAdd={() => handleOpenModal('Tambah Staf', [{name:'username', label:'Username'}, {name:'password', label:'Password'}, {name:'nama', label:'Nama Lengkap'}, {name:'role', label:'Role', type:'select', options:['Super Admin', 'Kepala Sekolah', 'Bendahara']}], 
          (d) => { const newId = Date.now(); setUsers([...users, {id: newId, ...d}]); syncDatabase('Users', 'POST', null, [newId, d.username, d.password, d.nama, d.role]); })}
        onEdit={(row) => handleOpenModal('Edit Staf', [{name:'username', label:'Username'}, {name:'password', label:'Password'}, {name:'nama', label:'Nama Lengkap'}, {name:'role', label:'Role', type:'select', options:['Super Admin', 'Kepala Sekolah', 'Bendahara']}], 
          (d) => { setUsers(users.map(u => u.id === row.id ? {...u, ...d} : u)); syncDatabase('Users', 'PUT', row.id, [row.id, d.username, d.password, d.nama, d.role]); }, row)}
        onDelete={(id) => { setUsers(users.filter(u => u.id !== id)); syncDatabase('Users', 'DELETE', id); }}
      />;

      case 'akun_kas': return <DataTable title="Data Akun Kas & Bank" columns={[{key:'kode', label:'Kode Akun'}, {key:'nama', label:'Nama Akun'}]} data={akunKas}
        onAdd={() => handleOpenModal('Tambah Akun Kas', [{name:'kode', label:'Kode'}, {name:'nama', label:'Nama Akun'}], 
          (d) => { const newId = Date.now(); setAkunKas([...akunKas, {id: newId, ...d}]); syncDatabase('AkunKas', 'POST', null, [newId, d.kode, d.nama]); })}
        onEdit={(row) => handleOpenModal('Edit Akun Kas', [{name:'kode', label:'Kode'}, {name:'nama', label:'Nama Akun'}], 
          (d) => { setAkunKas(akunKas.map(a => a.id === row.id ? {...a, ...d} : a)); syncDatabase('AkunKas', 'PUT', row.id, [row.id, d.kode, d.nama]); }, row)}
        onDelete={(id) => { setAkunKas(akunKas.filter(a => a.id !== id)); syncDatabase('AkunKas', 'DELETE', id); }}
      />;

      case 'siswa': return (
        <>
          <input type="file" accept=".xls,.xlsx" className="hidden" ref={fileInputSiswaRef} onChange={handleImportSiswa} />
          <DataTable title="Master Data Siswa" 
            customHeader={<button onClick={() => fileInputSiswaRef.current.click()} className="flex gap-2 bg-emerald-600 text-white px-4 py-2 rounded-lg text-sm"><Upload size={16} /> Impor Excel</button>}
            columns={[{key:'nis', label:'NIS'}, {key:'nisn', label:'NISN'}, {key:'nama', label:'Nama Lengkap'}, {key:'kelas', label:'Kelas'}, {key:'tanggalLahir', label:'Tgl Lahir'}, {key:'wali', label:'Wali Murid'}, {key:'noWa', label:'No. WA Ortu'}]} data={siswa}
            onAdd={() => handleOpenModal('Tambah Siswa', [{name:'nis', label:'NIS'}, {name:'nisn', label:'NISN'}, {name:'password', label:'Password'}, {name:'nama', label:'Nama Lengkap'}, {name:'kelas', label:'Kelas'}, {name:'tanggalLahir', label:'Tanggal Lahir', type:'date'}, {name:'wali', label:'Nama Wali'}, {name:'noWa', label:'No. WA Ortu'}], 
              (d) => { const newId = Date.now(); setSiswa([...siswa, {id: newId, ...d}]); syncDatabase('Siswa', 'POST', null, [newId, d.nis, d.nisn, d.password, d.nama, d.kelas, d.wali, d.noWa, d.tanggalLahir]); })}
            onEdit={(row) => handleOpenModal('Edit Data Siswa', [{name:'nis', label:'NIS'}, {name:'nisn', label:'NISN'}, {name:'password', label:'Password'}, {name:'nama', label:'Nama Lengkap'}, {name:'kelas', label:'Kelas'}, {name:'tanggalLahir', label:'Tanggal Lahir', type:'date'}, {name:'wali', label:'Nama Wali'}, {name:'noWa', label:'No. WA Ortu'}], 
              (d) => { setSiswa(siswa.map(s => s.id === row.id ? {...s, ...d} : s)); syncDatabase('Siswa', 'PUT', row.id, [row.id, d.nis, d.nisn, d.password, d.nama, d.kelas, d.wali, d.noWa, d.tanggalLahir]); }, row)}
            onDelete={(id) => { setSiswa(siswa.filter(s => s.id !== id)); syncDatabase('Siswa', 'DELETE', id); }}
          />
        </>
      );

      case 'tanggungan': 
        const groupedTanggungan = siswa.map(s => {
          const studentBills = tanggungan.filter(t => t.nis === s.nis);
          const totalTagihan = studentBills.reduce((sum, t) => sum + Number(t.nilai_tagihan), 0);
          const totalKekurangan = studentBills.reduce((sum, t) => sum + Number(t.sisa_tagihan), 0);
          return { ...s, studentBills, totalTagihan, totalKekurangan };
        });
        const filteredGroupedTanggungan = groupedTanggungan.filter(s => s.nama.toLowerCase().includes(searchTagihan.toLowerCase()) || s.nis.includes(searchTagihan));

        return (
          <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
            <div className="p-4 border-b flex flex-col md:flex-row justify-between items-start md:items-center gap-4 bg-gray-50/50">
              <h2 className="text-lg font-semibold flex items-center gap-2"><FileText size={20} className="text-blue-600"/> Data Tagihan Siswa</h2>
              <div className="flex flex-wrap gap-2 w-full md:w-auto">
                <input type="text" value={searchTagihan} onChange={(e) => setSearchTagihan(e.target.value)} placeholder="Cari Siswa/NIS..." className="border rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 outline-none flex-1 md:flex-none" />
                <button onClick={() => handleOpenModal('Generate Tagihan Massal', [{name:'jenis', label:'Jenis Tagihan (Contoh: SPP Agt 2026)'}, {name:'nilai_tagihan', label:'Nilai Tagihan Per Siswa (Rp)', type:'number'}], 
                  (d) => {
                    const nilai = Number(d.nilai_tagihan);
                    const newBills = siswa.map((s, i) => ({ id: Date.now() + i, nis: s.nis, jenis: d.jenis, nilai_tagihan: nilai, telah_dibayar: 0, sisa_tagihan: nilai, status: 'Belum Lunas' }));
                    setTanggungan([...tanggungan, ...newBills]);
                    newBills.forEach(b => syncDatabase('Tanggungan', 'POST', null, [b.id, b.nis, b.jenis, b.nilai_tagihan, 0, b.nilai_tagihan, 'Belum Lunas']));
                })} className="flex gap-2 bg-blue-600 text-white px-4 py-2 rounded-lg text-sm">% Generate</button>
                <button onClick={() => handleOpenModal('Tambah Tagihan Manual', [{name:'nis', label:'Siswa', type:'select', options: siswa.map(s => ({value: s.nis, label: `${s.nis} - ${s.nama}`}))}, {name:'jenis', label:'Jenis Tagihan'}, {name:'nilai_tagihan', label:'Nilai Tagihan (Rp)', type:'number'}], 
                  (d) => {
                    const nilai = Number(d.nilai_tagihan);
                    const newId = Date.now();
                    setTanggungan([...tanggungan, {id: newId, ...d, nilai_tagihan: nilai, telah_dibayar: 0, sisa_tagihan: nilai, status: 'Belum Lunas'}]);
                    syncDatabase('Tanggungan', 'POST', null, [newId, d.nis, d.jenis, nilai, 0, nilai, 'Belum Lunas']);
                })} className="flex gap-2 bg-green-600 text-white px-4 py-2 rounded-lg text-sm"><Plus size={16} /> Tambah</button>
              </div>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-left text-sm text-gray-600"><thead className="bg-gray-50 text-xs uppercase"><tr><th className="px-6 py-4">Nama Siswa</th><th className="px-6 py-4">Total Tagihan</th><th className="px-6 py-4">Total Kekurangan</th><th className="px-6 py-4 text-right">Aksi</th></tr></thead>
                <tbody className="divide-y divide-gray-100">
                  {filteredGroupedTanggungan.map((row) => (
                    <React.Fragment key={row.id}>
                      <tr className="hover:bg-blue-50/30 cursor-pointer" onClick={() => setExpandedRow(expandedRow === row.nis ? null : row.nis)}>
                        <td className="px-6 py-4 font-medium flex items-center gap-2">{expandedRow === row.nis ? <ChevronUp size={16}/> : <ChevronDown size={16}/>}{row.nama}</td>
                        <td className="px-6 py-4">{formatRupiah(row.totalTagihan)}</td>
                        <td className="px-6 py-4 font-bold text-red-600">{formatRupiah(row.totalKekurangan)}</td>
                        <td className="px-6 py-4 text-right flex justify-end gap-2">
                          <button onClick={(e) => { e.stopPropagation(); handleKirimWA(row); }} className="flex gap-1 bg-green-500 text-white px-3 py-1 rounded text-sm"><MessageCircle size={16} /> WA</button>
                        </td>
                      </tr>
                      {expandedRow === row.nis && (
                        <tr className="bg-gray-50"><td colSpan="4" className="px-6 py-4">
                          <div className="bg-white rounded-lg border overflow-hidden"><table className="w-full text-left text-sm"><thead className="bg-gray-100 text-xs uppercase"><tr><th className="px-4 py-3">Rincian Pembayaran</th><th className="px-4 py-3">Tagihan</th><th className="px-4 py-3">Dibayar</th><th className="px-4 py-3">Sisa</th><th className="px-4 py-3 text-right">Aksi</th></tr></thead>
                            <tbody>
                              {row.studentBills.map((bill) => (
                                <tr key={bill.id} className="border-t hover:bg-gray-50">
                                  <td className="px-4 py-3 uppercase">{bill.jenis}</td><td className="px-4 py-3">{formatRupiah(bill.nilai_tagihan)}</td><td className="px-4 py-3 text-green-600">{formatRupiah(bill.telah_dibayar)}</td><td className="px-4 py-3 font-bold text-red-600">{formatRupiah(bill.sisa_tagihan)}</td>
                                  <td className="px-4 py-3 text-right flex justify-end gap-2">
                                    <button onClick={(e) => { e.stopPropagation(); handleOpenModal('Edit Rincian Tagihan', [{name: 'jenis', label: 'Jenis Tagihan'}, {name: 'nilai_tagihan', label: 'Nilai Tagihan (Rp)', type: 'number'}], (d) => {
                                      const newNilai = Number(d.nilai_tagihan);
                                      const newSisa = newNilai - Number(bill.telah_dibayar);
                                      const newStatus = newSisa <= 0 ? 'Lunas' : 'Belum Lunas';
                                      setTanggungan(tanggungan.map(t => t.id === bill.id ? { ...t, jenis: d.jenis, nilai_tagihan: newNilai, sisa_tagihan: newSisa, status: newStatus } : t));
                                      syncDatabase('Tanggungan', 'PUT', bill.id, [bill.id, bill.nis, d.jenis, newNilai, bill.telah_dibayar, newSisa, newStatus]);
                                    }, bill); }} className="p-1 bg-blue-500 text-white rounded"><Edit size={14}/></button>
                                    <button onClick={(e) => { e.stopPropagation(); if(window.confirm('Hapus tagihan?')) { setTanggungan(tanggungan.filter(t => t.id !== bill.id)); syncDatabase('Tanggungan', 'DELETE', bill.id); } }} className="p-1 bg-red-500 text-white rounded"><Trash2 size={14} /></button>
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table></div>
                        </td></tr>
                      )}
                    </React.Fragment>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        );

      case 'kat_pendapatan': return <DataTable title="Kategori Jenis Pendapatan" columns={[{key:'nama', label:'Nama Kategori'}]} data={kategoriPendapatan}
        onAdd={() => handleOpenModal('Tambah Kategori Pendapatan', [{name:'nama', label:'Nama Kategori'}], (d) => { const newId = Date.now(); setKategoriPendapatan([...kategoriPendapatan, {id: newId, ...d}]); syncDatabase('KatPendapatan', 'POST', null, [newId, d.nama]); })}
        onEdit={(row) => handleOpenModal('Edit Kategori Pendapatan', [{name:'nama', label:'Nama Kategori'}], (d) => { setKategoriPendapatan(kategoriPendapatan.map(k => k.id === row.id ? {...k, ...d} : k)); syncDatabase('KatPendapatan', 'PUT', row.id, [row.id, d.nama]); }, row)}
        onDelete={(id) => { setKategoriPendapatan(kategoriPendapatan.filter(k => k.id !== id)); syncDatabase('KatPendapatan', 'DELETE', id); }}
      />;

      case 'kat_pengeluaran': return <DataTable title="Kategori Pengeluaran" columns={[{key:'nama', label:'Nama Kategori'}]} data={kategoriPengeluaran}
        onAdd={() => handleOpenModal('Tambah Kategori Pengeluaran', [{name:'nama', label:'Nama Kategori'}], (d) => { const newId = Date.now(); setKategoriPengeluaran([...kategoriPengeluaran, {id: newId, ...d}]); syncDatabase('KatPengeluaran', 'POST', null, [newId, d.nama]); })}
        onEdit={(row) => handleOpenModal('Edit Kategori Pengeluaran', [{name:'nama', label:'Nama Kategori'}], (d) => { setKategoriPengeluaran(kategoriPengeluaran.map(k => k.id === row.id ? {...k, ...d} : k)); syncDatabase('KatPengeluaran', 'PUT', row.id, [row.id, d.nama]); }, row)}
        onDelete={(id) => { setKategoriPengeluaran(kategoriPengeluaran.filter(k => k.id !== id)); syncDatabase('KatPengeluaran', 'DELETE', id); }}
      />;

      case 'bayar_siswa': return <DataTable title="Riwayat Pembayaran Siswa"
        columns={[{key:'tanggal', label:'Tanggal'}, {key:'keterangan', label:'Keterangan'}, {key:'nominal', label:'Nominal', render: (val) => <span className="text-green-600 font-medium">+{formatRupiah(val)}</span>}]} data={transaksi.filter(t => t.tipe === 'MASUK_SISWA')}
        onAdd={() => handleOpenModal('Input Pembayaran Siswa', [
          {name:'tanggal', label:'Tanggal Bayar', type:'date'},
          {name:'akunKasId', label:'Terima Uang Ke (Kas/Bank)', type:'select', options: akunKas.map(a => ({value: a.id, label: a.nama}))},
          {name:'tanggunganId', label:'Pilih Tagihan (Yg Nunggak)', type:'select', options: tanggungan.filter(t=>t.sisa_tagihan > 0).map(t => ({value: t.id, label: `${siswa.find(s=>s.nis===t.nis)?.nama} - ${t.jenis} (Sisa: ${formatRupiah(t.sisa_tagihan)})`}))},
          {name:'nominal_bayar', label:'Nominal Bayar (Rp)', type:'number'},
        ], (d) => {
          const tgn = tanggungan.find(t => t.id === Number(d.tanggunganId));
          const ssw = siswa.find(s => s.nis === tgn.nis);
          const bayar = Number(d.nominal_bayar);
          if (bayar > tgn.sisa_tagihan) return alert('Nominal lebih besar dari tagihan!');

          const newSisa = tgn.nilai_tagihan - (Number(tgn.telah_dibayar) + bayar);
          const newStatus = newSisa <= 0 ? 'Lunas' : 'Belum Lunas';
          
          setTanggungan(tanggungan.map(t => t.id === tgn.id ? {...t, telah_dibayar: Number(tgn.telah_dibayar) + bayar, sisa_tagihan: newSisa, status: newStatus} : t));
          syncDatabase('Tanggungan', 'PUT', tgn.id, [tgn.id, tgn.nis, tgn.jenis, tgn.nilai_tagihan, Number(tgn.telah_dibayar) + bayar, newSisa, newStatus]);

          const newId = Date.now();
          const ket = `Pembayaran ${tgn.jenis} an. ${ssw.nama}`;
          setTransaksi([...transaksi, { id: newId, tanggal: d.tanggal, tipe: 'MASUK_SISWA', nominal: bayar, keterangan: ket, refId: tgn.id, akunKasId: Number(d.akunKasId) }]);
          syncDatabase('Transaksi', 'POST', null, [newId, d.tanggal, 'MASUK_SISWA', bayar, ket, '', tgn.id, Number(d.akunKasId)]);
        })}
        onDelete={(id) => { setTransaksi(transaksi.filter(t => t.id !== id)); syncDatabase('Transaksi', 'DELETE', id); }}
      />;

      case 'pendapatan_lain': return <DataTable title="Penerimaan Pendapatan Lain"
        columns={[{key:'tanggal', label:'Tanggal'}, {key:'kategori', label:'Sumber Dana'}, {key:'keterangan', label:'Keterangan'}, {key:'nominal', label:'Nominal', render: (val) => <span className="text-green-600 font-medium">+{formatRupiah(val)}</span>}]} data={transaksi.filter(t => t.tipe === 'MASUK_LAIN')}
        onAdd={() => handleOpenModal('Input Pendapatan Lain', [
          {name:'tanggal', label:'Tanggal', type:'date'},
          {name:'kategori', label:'Sumber Dana', type:'select', options: kategoriPendapatan.map(k=>k.nama)},
          {name:'akunKasId', label:'Terima Uang Ke', type:'select', options: akunKas.map(a => ({value: a.id, label: a.nama}))},
          {name:'nominal', label:'Nominal (Rp)', type:'number'},
          {name:'keterangan', label:'Keterangan Tambahan'}
        ], (d) => {
          const newId = Date.now();
          setTransaksi([...transaksi, {id: newId, tipe: 'MASUK_LAIN', ...d, nominal: Number(d.nominal), akunKasId: Number(d.akunKasId)}]);
          syncDatabase('Transaksi', 'POST', null, [newId, d.tanggal, 'MASUK_LAIN', Number(d.nominal), d.keterangan, d.kategori, '', Number(d.akunKasId)]);
        })}
        onDelete={(id) => { setTransaksi(transaksi.filter(t => t.id !== id)); syncDatabase('Transaksi', 'DELETE', id); }}
      />;

      case 'pengeluaran': return <DataTable title="Pencatatan Pengeluaran Kas"
        columns={[{key:'tanggal', label:'Tanggal'}, {key:'kategori', label:'Kategori'}, {key:'keterangan', label:'Keterangan'}, {key:'nominal', label:'Nominal', render: (val) => <span className="text-red-600 font-medium">-{formatRupiah(val)}</span>}]} data={transaksi.filter(t => t.tipe === 'KELUAR')}
        onAdd={() => handleOpenModal('Input Pengeluaran Kas', [
          {name:'tanggal', label:'Tanggal', type:'date'},
          {name:'kategori', label:'Kategori Biaya', type:'select', options: kategoriPengeluaran.map(k=>k.nama)},
          {name:'akunKasId', label:'Keluar Dari', type:'select', options: akunKas.map(a => ({value: a.id, label: a.nama}))},
          {name:'nominal', label:'Nominal Keluar (Rp)', type:'number'},
          {name:'keterangan', label:'Penerima/Keterangan'}
        ], (d) => {
          const newId = Date.now();
          setTransaksi([...transaksi, {id: newId, tipe: 'KELUAR', ...d, nominal: Number(d.nominal), akunKasId: Number(d.akunKasId)}]);
          syncDatabase('Transaksi', 'POST', null, [newId, d.tanggal, 'KELUAR', Number(d.nominal), d.keterangan, d.kategori, '', Number(d.akunKasId)]);
        })}
        onDelete={(id) => { setTransaksi(transaksi.filter(t => t.id !== id)); syncDatabase('Transaksi', 'DELETE', id); }}
      />;

      case 'mutasi_akun':
        const mutasiTransaksi = mutasiAkunId ? transaksi.filter(t => Number(t.akunKasId) === Number(mutasiAkunId)).sort((a,b) => new Date(a.tanggal) - new Date(b.tanggal)) : [];
        let saldoBerjalan = 0;
        return (
          <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
            <div className="p-6 border-b flex justify-between items-center bg-gray-50/50">
              <h2 className="text-xl font-bold flex items-center gap-2"><BookOpen className="text-blue-600"/> Buku Besar / Mutasi Kas</h2>
              <select className="p-2.5 border rounded-lg focus:ring-2 focus:ring-blue-500 font-medium" value={mutasiAkunId} onChange={(e) => setMutasiAkunId(e.target.value)}>
                <option value="">-- Pilih Akun Bank/Kas --</option>{akunKas.map(a => <option key={a.id} value={a.id}>{a.kode} - {a.nama}</option>)}
              </select>
            </div>
            {mutasiAkunId === '' ? <div className="p-12 text-center text-gray-400">Silakan pilih akun kas di atas.</div> : (
              <div className="overflow-x-auto"><table className="w-full text-left text-sm"><thead className="bg-gray-800 text-white"><tr><th className="px-6 py-4">Tanggal</th><th className="px-6 py-4">Keterangan</th><th className="px-6 py-4 text-right">Debit (Masuk)</th><th className="px-6 py-4 text-right">Kredit (Keluar)</th><th className="px-6 py-4 text-right">Saldo</th></tr></thead>
                <tbody>
                  {mutasiTransaksi.length === 0 ? <tr><td colSpan="5" className="px-6 py-8 text-center text-gray-400">Belum ada transaksi.</td></tr> : 
                  mutasiTransaksi.map((t, i) => {
                    const isMasuk = t.tipe.includes('MASUK');
                    saldoBerjalan = isMasuk ? saldoBerjalan + Number(t.nominal) : saldoBerjalan - Number(t.nominal);
                    return (<tr key={i} className="border-b hover:bg-gray-50"><td className="px-6 py-4">{t.tanggal}</td><td className="px-6 py-4 font-medium">{t.keterangan}</td><td className="px-6 py-4 text-right text-green-600">{isMasuk ? formatRupiah(t.nominal) : '-'}</td><td className="px-6 py-4 text-right text-red-600">{!isMasuk ? formatRupiah(t.nominal) : '-'}</td><td className="px-6 py-4 text-right font-bold">{formatRupiah(saldoBerjalan)}</td></tr>)
                  })}
                </tbody>
              </table></div>
            )}
          </div>
        );

      case 'lap_arus_kas': return (
        <div className="bg-white rounded-xl shadow-sm border p-6">
          <h2 className="text-xl font-bold mb-6 flex items-center gap-2"><ArrowRightLeft/> Laporan Arus Kas</h2>
          <table className="w-full text-left text-sm"><thead className="bg-gray-800 text-white"><tr><th className="px-4 py-3">Tanggal</th><th className="px-4 py-3">Keterangan</th><th className="px-4 py-3 text-right">Kas Masuk</th><th className="px-4 py-3 text-right">Kas Keluar</th></tr></thead>
            <tbody>
              {transaksi.sort((a,b) => new Date(a.tanggal) - new Date(b.tanggal)).map((t, i) => (
                <tr key={i} className="border-b hover:bg-gray-50"><td className="px-4 py-3">{t.tanggal}</td><td className="px-4 py-3">{t.keterangan}</td><td className="px-4 py-3 text-right text-green-600">{t.tipe.includes('MASUK') ? formatRupiah(t.nominal) : '-'}</td><td className="px-4 py-3 text-right text-red-600">{t.tipe === 'KELUAR' ? formatRupiah(t.nominal) : '-'}</td></tr>
              ))}
              <tr className="bg-gray-100 font-bold"><td colSpan="2" className="px-4 py-4 text-right">TOTAL</td><td className="px-4 py-4 text-right text-green-600">{formatRupiah(totalKasMasuk)}</td><td className="px-4 py-4 text-right text-red-600">{formatRupiah(totalKasKeluar)}</td></tr>
            </tbody>
          </table>
        </div>
      );

      case 'lap_laba_rugi': return (
        <div className="max-w-3xl mx-auto bg-white rounded-xl shadow-sm border p-8">
           <div className="text-center mb-8 border-b pb-4"><h2 className="text-2xl font-bold">Laporan Laba Rugi</h2></div>
           <div className="space-y-6">
             <div><h3 className="font-bold text-lg text-blue-800 bg-blue-50 p-2 rounded">PENDAPATAN</h3>
               <div className="flex justify-between p-2 border-b"><span>Penerimaan Siswa</span><span>{formatRupiah(transaksi.filter(t=>t.tipe==='MASUK_SISWA').reduce((s,t)=>s+Number(t.nominal),0))}</span></div>
               <div className="flex justify-between p-2 border-b"><span>Pendapatan Lain</span><span>{formatRupiah(transaksi.filter(t=>t.tipe==='MASUK_LAIN').reduce((s,t)=>s+Number(t.nominal),0))}</span></div>
               <div className="flex justify-between p-3 font-bold bg-gray-50"><span>Total Pendapatan</span><span className="text-green-600">{formatRupiah(totalKasMasuk)}</span></div>
             </div>
             <div><h3 className="font-bold text-lg text-red-800 bg-red-50 p-2 rounded">PENGELUARAN</h3>
               {kategoriPengeluaran.map(k => {
                 const totalByKat = transaksi.filter(t => t.tipe === 'KELUAR' && t.kategori === k.nama).reduce((s,t)=>s+Number(t.nominal), 0);
                 if(totalByKat === 0) return null;
                 return <div key={k.id} className="flex justify-between p-2 border-b"><span>{k.nama}</span><span>{formatRupiah(totalByKat)}</span></div>
               })}
               <div className="flex justify-between p-3 font-bold bg-gray-50"><span>Total Pengeluaran</span><span className="text-red-600">{formatRupiah(totalKasKeluar)}</span></div>
             </div>
             <div className={`flex justify-between p-4 rounded-lg text-xl font-bold text-white ${saldoAkhir >= 0 ? 'bg-green-600' : 'bg-red-600'}`}>
                <span>{saldoAkhir >= 0 ? 'SURPLUS (Laba)' : 'DEFISIT (Rugi)'}</span><span>{formatRupiah(saldoAkhir)}</span>
             </div>
           </div>
        </div>
      );

      case 'akun_siswa': return (
        <>
          <input type="file" accept=".xls,.xlsx" className="hidden" ref={fileInputAkunSiswaRef} onChange={handleImportAkunSiswa} />
          <DataTable title="Manajemen Akses Login Siswa" customHeader={<button onClick={() => fileInputAkunSiswaRef.current.click()} className="flex gap-2 bg-emerald-600 text-white px-4 py-2 rounded-lg text-sm"><Upload size={16} /> Impor Excel</button>}
            columns={[{key:'nis', label:'NIS'}, {key:'nama', label:'Nama'}, {key:'kelas', label:'Kelas'}, {key:'password', label:'Password Akses', render: (val) => <span className="font-mono bg-gray-100 border px-2 py-1 rounded">{val}</span>}]} data={siswa}
            onEdit={(row) => handleOpenModal(`Ganti Password: ${row.nama}`, [{name:'password', label:'Password Baru'}], (d) => {
              setSiswa(siswa.map(s => s.id === row.id ? {...s, password: d.password} : s));
              syncDatabase('Siswa', 'PUT', row.id, [row.id, row.nis, row.nisn, d.password, row.nama, row.kelas, row.wali, row.noWa, row.tanggalLahir]);
            }, row)}
          />
        </>
      );

      case 'tagihan_pribadi': 
        const tagihanSaya = tanggungan.filter(t => t.nis === currentUser.nis);
        const totalKekuranganSaya = tagihanSaya.reduce((sum, t) => sum + Number(t.sisa_tagihan), 0);
        return (
          <div className="space-y-6">
            <div className="bg-white rounded-xl shadow-sm border p-6 flex justify-between items-center">
               <div><h2 className="text-2xl font-bold">Halo, {currentUser.nama}!</h2><p className="text-gray-500">NIS: {currentUser.nis} | Kelas: {currentUser.kelas}</p></div>
               <div className="bg-blue-50 p-4 rounded-xl text-right"><p className="text-sm font-medium">Sisa Kekurangan</p><h3 className={`text-3xl font-bold ${totalKekuranganSaya > 0 ? 'text-red-600' : 'text-green-600'}`}>{formatRupiah(totalKekuranganSaya)}</h3></div>
            </div>
            <DataTable title="Rincian Tagihan Anda" columns={[{key:'jenis', label:'Jenis Tagihan'}, {key:'nilai_tagihan', label:'Total', render: (val) => formatRupiah(val)}, {key:'telah_dibayar', label:'Dibayar', render: (val) => <span className="text-green-600">{formatRupiah(val)}</span>}, {key:'sisa_tagihan', label:'Kekurangan', render: (val) => <span className="font-bold text-red-600">{formatRupiah(val)}</span>}, {key:'status', label:'Status'}]} data={tagihanSaya} />
          </div>
        );

      default: return <div className="p-12 text-center text-gray-500">Fitur belum tersedia atau tidak dikenali.</div>;
    }
  };

  return (
    <div className="flex h-screen bg-gray-100 font-sans">
      {/* Overlay background hitam saat sidebar terbuka di HP */}
      {isMobileMenuOpen && (
        <div 
          className="fixed inset-0 bg-black/50 z-30 md:hidden transition-opacity"
          onClick={() => setIsMobileMenuOpen(false)}
        />
      )}

      <div className={`fixed inset-y-0 left-0 z-40 w-64 bg-slate-900 text-gray-300 transform transition-transform duration-300 ease-in-out ${isMobileMenuOpen ? 'translate-x-0' : '-translate-x-full'} md:relative md:translate-x-0 overflow-y-auto`}>
        <div className="p-6 flex items-center justify-between border-b border-slate-800"><div className="flex items-center gap-3"><div className="bg-blue-500 p-2 rounded-lg"><Banknote className="text-white"/></div><h1 className="text-xl font-bold text-white">EduFinance</h1></div></div>
        <div className="px-4 py-6 space-y-6">
          {menuGroups.map((group, gIndex) => (
            <div key={gIndex}><h2 className="text-xs font-semibold text-slate-500 uppercase mb-3 px-3">{group.title}</h2><ul className="space-y-1">
                {group.items.map(item => (<li key={item.id}><button onClick={() => { setActiveMenu(item.id); setIsMobileMenuOpen(false); }} className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg ${activeMenu === item.id ? 'bg-blue-600 text-white' : 'hover:bg-slate-800'}`}>{item.icon}<span className="text-sm">{item.label}</span></button></li>))}
              </ul></div>
          ))}
        </div>
      </div>
      <div className="flex-1 flex flex-col overflow-hidden">
        <header className="bg-white border-b h-16 flex items-center justify-between px-4 sm:px-6 z-10">
          <div className="flex items-center gap-3">
            <button onClick={() => setIsMobileMenuOpen(true)} className="md:hidden p-2 text-gray-600 hover:bg-gray-100 rounded-lg">
              <Menu size={24} />
            </button>
            <h2 className="text-xl font-semibold hidden sm:block">{menuGroups.flatMap(g => g.items).find(i => i.id === activeMenu)?.label || 'Aplikasi'}</h2>
          </div>
          <div className="flex items-center gap-4"><div className="text-right hidden sm:block"><p className="text-sm font-semibold">{currentUser.nama}</p><p className="text-xs text-gray-500">{currentUser.role}</p></div><button onClick={() => setCurrentUser(null)} className="flex items-center gap-2 text-sm text-red-500 bg-red-50 px-3 py-2 rounded-lg"><LogOut size={16}/> Keluar</button></div>
        </header>
        <main className="flex-1 overflow-y-auto p-4 sm:p-6 bg-gray-50/50"><div className="max-w-7xl mx-auto">{renderContent()}</div></main>
      </div>
      <DynamicModal />
    </div>
  );
}
