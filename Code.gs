/**
 * BACKEND FINAL iMersWAstore - SINGLE STORE EDITION - KHUSUS RESELLER RIGHT BACA SYARTANYA DIMEMBER AREA
 * MEGA UPDATE: Kategori Dinamis, Kupon, Ongkir, PPN, Marquee Text & Homepage Sections
 * + THE HOLY TRINITY: Blog, Pop-up Promo, & Tentang Kami (About)
 * ARSITEKTUR: Single Store (1 Web = 1 Toko, Super Cepat & Anti-CORS)
 * RULES BACA DISINI YA WAJIB : https://ai.imersnesia.com/rulesimerswastorecode/
 */

const SCRIPT_PROP = PropertiesService.getScriptProperties();

// =========================================================================
// 1. MENU ADD-ON & SETUP AWAL (SINGLE STORE)
// =========================================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ iMersWAstore Setup')
    .addItem('🚀 Setup Database Single Store', 'setupSingleStore')
    .addToUi();
}

function setupSingleStore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SCRIPT_PROP.setProperty('key', ss.getId());
  const now = new Date();

  // 1. APP SETTINGS TABLE
  let appSheet = ss.getSheetByName('App_Settings') || ss.insertSheet('App_Settings');
  if (appSheet.getLastRow() === 0) {
    appSheet.appendRow(['setting_name', 'setting_value']);
    appSheet.appendRow(['app_name', 'Toko Utama Ku']); 
    appSheet.appendRow(['marquee_text', '🔥 Promo Spesial Bulan Ini! Gratis Ongkir ke Seluruh Kota! 🔥']); 
    appSheet.appendRow(['homepage_sections', '["hero", "categories", "products", "footer"]']); 
    // --- BANNER DINAMIS ---
    appSheet.appendRow(['banner_slider', '["https://placehold.co/800x350/3b82f6/ffffff?text=Promo+Gila", "https://placehold.co/800x350/10b981/ffffff?text=Gratis+Ongkir"]']); 
    appSheet.appendRow(['banner_mid_image', 'https://placehold.co/800x200/f59e0b/ffffff?text=Flash+Sale+Extra+Diskon']); 
    appSheet.appendRow(['show_banner_mid', 'on']); 
    // --- FITUR BARU: POP-UP, ABOUT, BLOG TOGGLE ---
    appSheet.appendRow(['show_popup', 'off']);
    appSheet.appendRow(['popup_image', 'https://placehold.co/500x500/ec4899/ffffff?text=BIG+SALE']);
    appSheet.appendRow(['popup_desc', 'Diskon gila-gilaan akhir tahun. Langsung sikat sebelum kehabisan!']);
    appSheet.appendRow(['popup_btn_text', 'Lihat Promo']);
    appSheet.appendRow(['popup_btn_link', '#produk']);
    appSheet.appendRow(['about_us_text', '<p>Selamat datang di <b>Toko Utama Ku</b>. Kami adalah toko online terpercaya yang menyediakan berbagai macam kebutuhan Anda sejak tahun 2026.</p>']);
    appSheet.appendRow(['show_blog_section', 'on']);
    appSheet.appendRow(['footer_branding', 'Powered by iMersWAstore']); // <-- TAMBAHAN FOOTER DINAMIS
    appSheet.getRange("A1:B1").setFontWeight("bold").setBackground("#ef4444").setFontColor("white");
  }
  
  // 2. USERS TABLE (Hanya untuk Admin dan Staf Kasir - SEKARANG 30 KOLOM)
  let userSheet = ss.getSheetByName('Users') || ss.insertSheet('Users');
  let userHeaders = [
    'id', 'username', 'email', 'password', 'role', 'nama_toko', 
    'status', 'created_at', 'logo', 'wa_number', 'auto_reply', 'qris_url', 
    'bank_info', 'gateway_token', 'gateway_sender', 'gateway_provider', 'banner_url', 
    'theme_color', 'social_media', 'notif_templates', 'legal_settings', 'pixel_fb', 'pixel_tiktok', 'pixel_google', 'tax_settings',
    'secret_login_status', 'secret_tap_count', 'secret_pin', 'announcement_text', 'announcement_link'
  ];
  userSheet.getRange(1, 1, 1, userHeaders.length).setValues([userHeaders]).setFontWeight("bold").setBackground("#10b981").setFontColor("white");
  
  if (userSheet.getLastRow() === 1) {
    userSheet.appendRow([
      'ADM-001', 'admin', 'admin@toko.com', hashPassword('admin123'), 'admin', 'Toko Super', 
      'approved', now, '', '6281234567890', 'Halo Kak, saya mau order!', '', 
      'BCA 123456 a/n Budi', '', '', 'fonnte', '', 
      '#f8fafc', '[]', '{}', '{}', '', '', '', '{}',
      'off', '5', '1234', '🔥 Promo Spesial Bulan Ini! Gratis Ongkir!', ''
    ]);
  }

  // 3. PRODUCTS TABLE (Disederhanakan, tanpa id_user_toko - 11 Kolom)
  let prodSheet = ss.getSheetByName('Products') || ss.insertSheet('Products');
  let prodHeaders = [
    'id_produk', 'nama_produk', 'harga', 'stok', 'gambar', 'waktu', 
    'deskripsi', 'satuan', 'badge', 'variants_data', 'id_kategori'
  ];
  prodSheet.getRange(1, 1, 1, prodHeaders.length).setValues([prodHeaders]).setFontWeight("bold").setBackground("#3b82f6").setFontColor("white");

  if (prodSheet.getLastRow() === 1) {
    prodSheet.appendRow(['PRD-001', 'Kopi Susu Aren', 18000, '', 'https://placehold.co/300x300/10b981/ffffff?text=Kopi', now, 'Kopi susu dengan gula aren asli.', 'Cup', 'Terlaris', '[]', 'CAT-001']);
    prodSheet.appendRow(['PRD-002', 'Rice Bowl Ayam', 28000, 50, 'https://placehold.co/300x300/f59e0b/ffffff?text=Ayam', now, 'Nasi hangat dengan ayam katsu.', 'Porsi', 'Promo', '[]', 'CAT-002']);
  }

  // 4. ORDERS TABLE (17 Kolom)
  let orderSheet = ss.getSheetByName('Orders') || ss.insertSheet('Orders');
  let orderHeaders = [
    'id_order', 'nama_pembeli', 'total_belanja', 'status', 'detail_items', 'waktu', 
    'metode_pembayaran', 'uang_bayar', 'kembalian', 'no_wa_pembeli', 'subtotal', 'ongkir', 'diskon', 'kode_kupon', 'nama_area_ongkir',
    'pajak_ppn', 'biaya_layanan'
  ];
  orderSheet.getRange(1, 1, 1, orderHeaders.length).setValues([orderHeaders]).setFontWeight("bold").setBackground("#f59e0b").setFontColor("white");

  // 5. CATEGORIES TABLE (5 Kolom)
  let catSheet = ss.getSheetByName('Categories') || ss.insertSheet('Categories');
  let catHeaders = ['id_kategori', 'nama_kategori', 'is_active', 'sort_order', 'parent_id'];
  catSheet.getRange(1, 1, 1, catHeaders.length).setValues([catHeaders]).setFontWeight("bold").setBackground("#8b5cf6").setFontColor("white");

  if (catSheet.getLastRow() === 1) {
    catSheet.appendRow(['CAT-001', 'Minuman', 'true', 1, '']);
    catSheet.appendRow(['CAT-002', 'Makanan', 'true', 2, '']);
  }

  // 6. COUPONS TABLE (6 Kolom)
  let couponSheet = ss.getSheetByName('Coupons') || ss.insertSheet('Coupons');
  let couponHeaders = ['id_kupon', 'kode_kupon', 'tipe_diskon', 'nilai_diskon', 'min_belanja', 'status_aktif'];
  couponSheet.getRange(1, 1, 1, couponHeaders.length).setValues([couponHeaders]).setFontWeight("bold").setBackground("#ec4899").setFontColor("white");

  if (couponSheet.getLastRow() === 1) {
    couponSheet.appendRow(['CUP-001', 'WELCOME20', 'percent', 20, 50000, 'true']);
  }

  // 7. SHIPPING RATES TABLE (3 Kolom)
  let shipSheet = ss.getSheetByName('Shipping_Rates') || ss.insertSheet('Shipping_Rates');
  let shipHeaders = ['id_ongkir', 'nama_area', 'tarif_ongkir'];
  shipSheet.getRange(1, 1, 1, shipHeaders.length).setValues([shipHeaders]).setFontWeight("bold").setBackground("#14b8a6").setFontColor("white");

  if (shipSheet.getLastRow() === 1) {
    shipSheet.appendRow(['SHP-001', 'Dalam Kota (Instant)', 15000]);
    shipSheet.appendRow(['SHP-002', 'Luar Kota (Reguler)', 25000]);
  }

  // 8. ARTICLES / BLOG TABLE (6 Kolom) - HOLY TRINITY
  let articleSheet = ss.getSheetByName('Articles') || ss.insertSheet('Articles');
  let articleHeaders = ['id_artikel', 'judul', 'slug', 'gambar', 'isi_artikel', 'tanggal'];
  articleSheet.getRange(1, 1, 1, articleHeaders.length).setValues([articleHeaders]).setFontWeight("bold").setBackground("#6366f1").setFontColor("white");

  if (articleSheet.getLastRow() === 1) {
    articleSheet.appendRow([
      'ART-001', 
      'Selamat Datang di Toko Kami', 
      'selamat-datang-di-toko-kami', 
      'https://placehold.co/800x400/6366f1/ffffff?text=Blog+Baru', 
      '<p>Halo pelanggan setia! Kami sangat senang menyambut Anda di toko online terbaru kami. Dapatkan promo menarik setiap minggunya.</p>', 
      now
    ]);
  }

  // Hapus tab kosong bawaan (Sheet1) kalau masih ada
  let sheet1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('Sheet 1');
  if(sheet1) ss.deleteSheet(sheet1);

  SpreadsheetApp.getUi().alert("✅ Setup Database Single Store Berhasil! Semua Tabel Lurus dan Bersih!");
}

// =========================================================================
// ZERO TRUST SECURITY (Untuk Membedakan Admin Utama & Kasir)
// =========================================================================
function checkZeroTrust(reqId, reqToken, userData) {
  if (!reqId || !reqToken) return { isValid: false, role: 'guest', id: null };
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] === reqId && userData[i][3] === reqToken) {
      return { isValid: true, role: userData[i][4], id: userData[i][0] };
    }
  }
  return { isValid: false, role: 'guest', id: null };
}

// =========================================================================
// 2. MAIN ROUTER API (POST)
// =========================================================================
function doPost(e) {
  const action = e.parameter.action;
  try {
    switch (action) {
      // SETTINGS & AUTH
      case 'get_app_settings': return handleGetAppSettings(e);
      case 'update_app_settings': return handleUpdateAppSettings(e);
      case 'update_store_settings': return handleUpdateStoreSettings(e);
      case 'login': return handleLogin(e);
      case 'change_password': return handleChangePassword(e);
      
      // USER / KASIR MANAGEMENT
      case 'get_users': return handleGetUsers(e);
      case 'create_account': return handleCreateAccount(e);
      case 'edit_user': return handleEditUser(e);
      case 'delete_user': return handleDeleteUser(e);
      
      // STORE FRONT (Etalase Utama - Ambil Semua Data Sekaligus)
      case 'get_store_catalog': return handleGetStoreCatalog(e);
      
      // PRODUK CRUD
      case 'add_product': return handleAddProduct(e);
      case 'get_products': return handleGetProducts(e);
      case 'edit_product': return handleEditProduct(e);
      case 'delete_product': return handleDeleteProduct(e);
      
      // KATEGORI CRUD
      case 'add_category': return handleAddCategory(e);
      case 'get_categories': return handleGetCategories(e);
      case 'edit_category': return handleEditCategory(e);
      case 'delete_category': return handleDeleteCategory(e);

      // KUPON CRUD
      case 'add_coupon': return handleAddCoupon(e);
      case 'get_coupons': return handleGetCoupons(e);
      case 'edit_coupon': return handleEditCoupon(e);
      case 'delete_coupon': return handleDeleteCoupon(e);
      case 'validate_coupon': return handleValidateCoupon(e);

      // ONGKIR CRUD
      case 'add_shipping': return handleAddShipping(e);
      case 'get_shipping': return handleGetShipping(e);
      case 'edit_shipping': return handleEditShipping(e);
      case 'delete_shipping': return handleDeleteShipping(e);

      // ARTIKEL / BLOG CRUD (HOLY TRINITY)
      case 'add_article': return handleAddArticle(e);
      case 'get_articles': return handleGetArticles(e);
      case 'edit_article': return handleEditArticle(e);
      case 'delete_article': return handleDeleteArticle(e);

      // ORDER CRUD & POS (DIARAHKAN KE Orders.gs)
      case 'add_order': return handleAddOrder(e);
      case 'get_orders': return handleGetOrders(e);
      case 'update_order_status': return handleUpdateOrderStatus(e);
      case 'delete_order': return handleDeleteOrder(e);

      default: return responseJSON({ success: false, message: "Action tidak dikenali!" });
    }
  } catch (error) {
    return responseJSON({ success: false, message: "Server Error: " + error.toString() });
  }
}

// =========================================================================
// 3. FUNGSI APP SETTINGS, STOREFRONT & USER
// =========================================================================
function handleGetAppSettings(e) {
  const data = getSheet('App_Settings').getDataRange().getValues();
  let settings = {};
  for (let i = 1; i < data.length; i++) {
    settings[data[i][0]] = data[i][1];
  }
  return responseJSON({ success: true, data: settings });
}

function handleUpdateAppSettings(e) {
  const sheetUsers = getSheet('Users');
  const dataUsers = sheetUsers.getDataRange().getValues();
  const zt = checkZeroTrust(e.parameter.requestor_id, e.parameter.requestor_token, dataUsers);
  
  if (!zt.isValid || zt.role !== 'admin') {
    return responseJSON({ success: false, message: "SECURITY ALERT: Hanya Admin Utama yang bisa mengubah ini!" });
  }

  const sheet = getSheet('App_Settings');
  const data = sheet.getDataRange().getValues();
  
  // Update dinamis semua key yang ada di App_Settings jika dikirim dari frontend
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    if (e.parameter[key] !== undefined) {
      sheet.getRange(i + 1, 2).setValue(e.parameter[key]);
    }
  }
  
  return responseJSON({ success: true, message: "Pengaturan App berhasil disimpan!" });
}

function handleGetStoreCatalog(e) {
  const userSheet = getSheet('Users').getDataRange().getValues();
  const prodSheet = getSheet('Products').getDataRange().getValues();
  
  if (userSheet.length < 2) return responseJSON({success: false, message: "Database toko belum disetup."});
  
  let storeInfo = {
     id: userSheet[1][0], 
     nama_toko: userSheet[1][5], 
     logo: userSheet[1][8], 
     wa_number: userSheet[1][9], 
     auto_reply: userSheet[1][10],
     qris_url: userSheet[1][11], 
     bank_info: userSheet[1][12], 
     banner_url: userSheet[1][16], 
     theme_color: userSheet[1][17] || '#f8fafc',
     social_media: userSheet[1][18] || '[]',
     notif_templates: userSheet[1][19] || '{}',
     legal_settings: userSheet[1][20] || '{}',
     pixel_fb: userSheet[1][21] || '',
     pixel_tiktok: userSheet[1][22] || '',
     pixel_google: userSheet[1][23] || '',
     tax_settings: userSheet[1][24] || '{}',
     secret_login_status: userSheet[1][25] || 'off',
     secret_tap_count: userSheet[1][26] || '5',
     secret_pin: userSheet[1][27] || '1234',
     announcement_text: userSheet[1][28] || '', // Fix Marquee
     announcement_link: userSheet[1][29] || ''  // Fix Marquee
  }; 
  
  let products = [];
  for (let i = 1; i < prodSheet.length; i++) {
     if (prodSheet[i][3] === '' || parseInt(prodSheet[i][3]) > 0 || isNaN(parseInt(prodSheet[i][3]))) { 
        products.push({ 
          id: prodSheet[i][0], 
          name: prodSheet[i][1], 
          price: prodSheet[i][2], 
          stock: prodSheet[i][3], 
          img: prodSheet[i][4], 
          desc: prodSheet[i][6], 
          unit: prodSheet[i][7], 
          badge: prodSheet[i][8] || '', 
          variants: prodSheet[i][9] || '[]', 
          category_id: prodSheet[i][10] || '' 
        });
     }
  }

  let categories = [];
  const catData = getSheet('Categories').getDataRange().getValues();
  for (let i = 1; i < catData.length; i++) {
    if (String(catData[i][2]) === 'true') {
      categories.push({ 
        id: catData[i][0], 
        name: catData[i][1], 
        sort: catData[i][3],
        parent_id: catData[i][4] || '' 
      });
    }
  }
  
  let shipping = [];
  const shipData = getSheet('Shipping_Rates').getDataRange().getValues();
  for (let i = 1; i < shipData.length; i++) {
     shipping.push({ id: shipData[i][0], area: shipData[i][1], rate: shipData[i][2] });
  }

  const appData = getSheet('App_Settings').getDataRange().getValues();
  let appSettings = {};
  for(let x=1; x<appData.length; x++) appSettings[appData[x][0]] = appData[x][1];

  let articles = [];
  try {
    const artData = getSheet('Articles').getDataRange().getValues();
    for (let i = 1; i < artData.length; i++) {
        articles.push({ id: artData[i][0], title: artData[i][1], slug: artData[i][2], img: artData[i][3], content: artData[i][4], date: artData[i][5] });
    }
  } catch (err) { console.log("Tabel Articles belum ada"); }

  // --- LOGIC FOOTER DINAMIS ---
  let footerBranding = appSettings.footer_branding || ("Powered by <b>" + storeInfo.nama_toko + "</b>");
  
  return responseJSON({
    success: true, 
    store: storeInfo, 
    app_settings: appSettings,
    footer_branding: footerBranding, // Kirim footer ke index.html
    products: products.reverse(), 
    categories: categories, 
    shipping_rates: shipping,
    articles: articles.reverse()
  });
}

// =========================================================================
// 4. AUTH / LOGIN & GANTI PASSWORD
// =========================================================================
function handleLogin(e) {
  const data = getSheet('Users').getDataRange().getValues();
  const hashedInput = hashPassword(e.parameter.password);
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === e.parameter.username && data[i][3] === hashedInput && data[i][6] === 'approved') { 
      return responseJSON({ 
        success: true, 
        message: "Login Berhasil", 
        data: { 
          id: data[i][0], 
          username: data[i][1], 
          role: data[i][4], 
          token: data[i][3], 
          nama_toko: data[i][5]
        }
      });
    }
  }
  return responseJSON({ success: false, message: "Username atau Password salah!" });
}

function handleChangePassword(e) {
  const sheet = getSheet('Users');
  const data = sheet.getDataRange().getValues();
  const userId = e.parameter.requestor_id;
  const oldPassHash = hashPassword(e.parameter.old_password);
  const newPassHash = hashPassword(e.parameter.new_password);

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      if (data[i][3] === oldPassHash) {
        sheet.getRange(i + 1, 4).setValue(newPassHash);
        return responseJSON({ success: true, message: "Password berhasil diubah! Silakan pakai sandi baru pada sesi login selanjutnya." });
      } else {
        return responseJSON({ success: false, message: "Password lama salah!" });
      }
    }
  }
  return responseJSON({ success: false, message: "User tidak ditemukan!" });
}

// =========================================================================
// 5. MANAJEMEN KASIR
// =========================================================================
function handleGetUsers(e) {
  const data = getSheet('Users').getDataRange().getValues();
  const zt = checkZeroTrust(e.parameter.requestor_id, e.parameter.requestor_token, data);
  if (!zt.isValid || zt.role !== 'admin') return responseJSON({ success: false, message: "Unauthorized" });

  let result = [];
  for (let i = 1; i < data.length; i++) {
    result.push({
      id: data[i][0],
      username: data[i][1],
      email: data[i][2],
      role: data[i][4],
      nama_toko: data[i][5],
      status: data[i][6],
      expired_at: data[i][7]
    });
  }
  return responseJSON({ success: true, data: result });
}

function handleCreateAccount(e) {
  const sheet = getSheet('Users');
  const data = sheet.getDataRange().getValues();
  const zt = checkZeroTrust(e.parameter.requestor_id, e.parameter.requestor_token, data);
  if (!zt.isValid || zt.role !== 'admin') return responseJSON({ success: false, message: "Unauthorized" });

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === e.parameter.new_username) return responseJSON({ success: false, message: "Username sudah dipakai!" });
  }

  const newId = 'USR-' + Math.floor(Math.random() * 100000);
  sheet.appendRow([
    newId, 
    e.parameter.new_username, 
    e.parameter.email || '', 
    hashPassword(e.parameter.new_password), 
    e.parameter.new_role || 'kasir', 
    e.parameter.nama_toko || '', 
    'approved', 
    new Date(), 
    '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
    'off', '5', '1234', '', '' // Pastikan index kolom sejajar hingga 30
  ]);
  return responseJSON({ success: true, message: "Akun kasir berhasil dibuat!" });
}

function handleEditUser(e) {
  const sheet = getSheet('Users');
  const data = sheet.getDataRange().getValues();
  const zt = checkZeroTrust(e.parameter.requestor_id, e.parameter.requestor_token, data);
  if (!zt.isValid || zt.role !== 'admin') return responseJSON({ success: false, message: "Unauthorized" });

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === e.parameter.target_user_id) {
      if (e.parameter.edit_username !== undefined) sheet.getRange(i+1, 2).setValue(e.parameter.edit_username);
      if (e.parameter.edit_email !== undefined) sheet.getRange(i+1, 3).setValue(e.parameter.edit_email);
      if (e.parameter.edit_nama_toko !== undefined) sheet.getRange(i+1, 6).setValue(e.parameter.edit_nama_toko);
      if (e.parameter.edit_status !== undefined) sheet.getRange(i+1, 7).setValue(e.parameter.edit_status);
      if (e.parameter.edit_expired_at !== undefined) sheet.getRange(i+1, 8).setValue(e.parameter.edit_expired_at);
      return responseJSON({ success: true, message: "Akun berhasil diupdate!" });
    }
  }
  return responseJSON({ success: false, message: "User tidak ditemukan!" });
}

function handleDeleteUser(e) {
  const sheet = getSheet('Users');
  const data = sheet.getDataRange().getValues();
  const zt = checkZeroTrust(e.parameter.requestor_id, e.parameter.requestor_token, data);
  if (!zt.isValid || zt.role !== 'admin') return responseJSON({ success: false, message: "Unauthorized" });

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === e.parameter.target_user_id) {
      if (data[i][4] === 'admin') return responseJSON({ success: false, message: "Akun Admin Utama tidak boleh dihapus!" });
      sheet.deleteRow(i + 1);
      return responseJSON({ success: true, message: "Akun kasir dihapus!" });
    }
  }
  return responseJSON({ success: false, message: "User tidak ditemukan!" });
}

// =========================================================================
// 5. CRUD PRODUK (SINGLE STORE)
// =========================================================================
function handleAddProduct(e) {
  const sheet = getSheet('Products');
  const prodId = 'PRD-' + Math.floor(Math.random() * 100000);
  
  sheet.appendRow([
    prodId, 
    e.parameter.nama_produk, 
    e.parameter.harga, 
    e.parameter.stok, 
    e.parameter.gambar || '', 
    new Date(), 
    e.parameter.deskripsi || '', 
    e.parameter.satuan || '',       
    e.parameter.badge || '',        
    e.parameter.variants_data || '[]', 
    e.parameter.id_kategori || ''   
  ]);
  
  return responseJSON({ success: true, message: "Produk ditambahkan!" });
}

function handleGetProducts(e) {
  const data = getSheet('Products').getDataRange().getValues(); 
  let result = [];
  for (let i = 1; i < data.length; i++) {
    result.push({ 
      id: data[i][0], 
      name: data[i][1], 
      price: data[i][2], 
      stock: data[i][3], 
      img: data[i][4], 
      desc: data[i][6], 
      unit: data[i][7],      
      badge: data[i][8] || '', 
      variants: data[i][9] || '[]', 
      category_id: data[i][10] || '' 
    });
  }
  return responseJSON({ success: true, data: result.reverse() });
}

function handleEditProduct(e) {
  const sheet = getSheet('Products');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === e.parameter.id_produk) {
      if(e.parameter.nama_produk !== undefined) sheet.getRange(i+1, 2).setValue(e.parameter.nama_produk);
      if(e.parameter.harga !== undefined) sheet.getRange(i+1, 3).setValue(e.parameter.harga);
      if(e.parameter.stok !== undefined) sheet.getRange(i+1, 4).setValue(e.parameter.stok);
      if(e.parameter.gambar !== undefined) sheet.getRange(i+1, 5).setValue(e.parameter.gambar);
      if(e.parameter.deskripsi !== undefined) sheet.getRange(i+1, 7).setValue(e.parameter.deskripsi);
      if(e.parameter.satuan !== undefined) sheet.getRange(i+1, 8).setValue(e.parameter.satuan); 
      if(e.parameter.badge !== undefined) sheet.getRange(i+1, 9).setValue(e.parameter.badge);  
      if(e.parameter.variants_data !== undefined) sheet.getRange(i+1, 10).setValue(e.parameter.variants_data);
      if(e.parameter.id_kategori !== undefined) sheet.getRange(i+1, 11).setValue(e.parameter.id_kategori);
      
      return responseJSON({ success: true, message: "Produk diupdate!" });
    }
  }
  return responseJSON({ success: false, message: "Produk tidak ditemukan." });
}

function handleDeleteProduct(e) {
  const sheet = getSheet('Products');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === e.parameter.id_produk) { 
      sheet.deleteRow(i + 1); 
      return responseJSON({ success: true, message: "Produk dihapus!" }); 
    }
  }
  return responseJSON({ success: false, message: "Produk tidak ditemukan." });
}

// =========================================================================
// 6. CRUD KATEGORI, KUPON, ONGKIR (SINGLE STORE)
// =========================================================================
function handleAddCategory(e) {
  getSheet('Categories').appendRow([
    'CAT-' + Math.floor(Math.random() * 100000), e.parameter.nama_kategori, 'true', e.parameter.sort_order || 0, e.parameter.parent_id || ''
  ]);
  return responseJSON({ success: true, message: "Kategori ditambahkan!" });
}

function handleGetCategories(e) {
  const data = getSheet('Categories').getDataRange().getValues(); 
  let result = [];
  for (let i = 1; i < data.length; i++) {
      result.push({ id: data[i][0], name: data[i][1], is_active: data[i][2], sort: data[i][3], parent_id: data[i][4] || '' });
  }
  return responseJSON({ success: true, data: result });
}

function handleEditCategory(e) {
  const sheet = getSheet('Categories');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === e.parameter.id_kategori) {
      if(e.parameter.nama_kategori !== undefined) sheet.getRange(i+1, 2).setValue(e.parameter.nama_kategori);
      if(e.parameter.is_active !== undefined) sheet.getRange(i+1, 3).setValue(e.parameter.is_active);
      if(e.parameter.sort_order !== undefined) sheet.getRange(i+1, 4).setValue(e.parameter.sort_order);
      if(e.parameter.parent_id !== undefined) sheet.getRange(i+1, 5).setValue(e.parameter.parent_id);
      return responseJSON({ success: true, message: "Kategori diupdate!" });
    }
  }
  return responseJSON({ success: false });
}

function handleDeleteCategory(e) {
  const sheet = getSheet('Categories');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === e.parameter.id_kategori) {
      sheet.deleteRow(i+1); 
      return responseJSON({ success: true, message: "Kategori dihapus!" });
    }
  }
  return responseJSON({ success: false });
}

function handleAddCoupon(e) {
  getSheet('Coupons').appendRow([
    'CUP-' + Math.floor(Math.random() * 100000), e.parameter.kode_kupon, e.parameter.tipe_diskon, e.parameter.nilai_diskon, e.parameter.min_belanja || 0, 'true'
  ]);
  return responseJSON({ success: true, message: "Kupon ditambahkan!" });
}

function handleGetCoupons(e) {
  const data = getSheet('Coupons').getDataRange().getValues(); 
  let result = [];
  for (let i = 1; i < data.length; i++) {
      result.push({ id: data[i][0], code: data[i][1], type: data[i][2], value: data[i][3], min_purchase: data[i][4], is_active: data[i][5] });
  }
  return responseJSON({ success: true, data: result });
}

function handleEditCoupon(e) {
  const sheet = getSheet('Coupons');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === e.parameter.id_kupon) {
      if(e.parameter.kode_kupon !== undefined) sheet.getRange(i+1, 2).setValue(e.parameter.kode_kupon);
      if(e.parameter.tipe_diskon !== undefined) sheet.getRange(i+1, 3).setValue(e.parameter.tipe_diskon);
      if(e.parameter.nilai_diskon !== undefined) sheet.getRange(i+1, 4).setValue(e.parameter.nilai_diskon);
      if(e.parameter.status_aktif !== undefined) sheet.getRange(i+1, 6).setValue(e.parameter.status_aktif);
      return responseJSON({ success: true, message: "Kupon diupdate!" });
    }
  }
  return responseJSON({ success: false });
}

function handleDeleteCoupon(e) {
  const sheet = getSheet('Coupons');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === e.parameter.id_kupon) { sheet.deleteRow(i+1); return responseJSON({ success: true }); }
  }
  return responseJSON({ success: false });
}

function handleValidateCoupon(e) {
  const data = getSheet('Coupons').getDataRange().getValues();
  const subtotal = parseInt(e.parameter.subtotal) || 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === e.parameter.kode_kupon && String(data[i][5]) === 'true') {
      if (subtotal < (parseInt(data[i][4]) || 0)) return responseJSON({ success: false, message: "Minimal belanja belum tercapai." });
      let diskon = data[i][2] === 'percent' ? (subtotal * (parseInt(data[i][3]) || 0) / 100) : (parseInt(data[i][3]) || 0);
      return responseJSON({ success: true, message: "Kupon valid!", diskon: diskon });
    }
  }
  return responseJSON({ success: false, message: "Kupon tidak ditemukan." });
}

function handleAddShipping(e) {
  getSheet('Shipping_Rates').appendRow(['SHP-' + Math.floor(Math.random() * 100000), e.parameter.nama_area, e.parameter.tarif_ongkir]);
  return responseJSON({ success: true, message: "Ongkir ditambahkan!" });
}

function handleGetShipping(e) {
  const data = getSheet('Shipping_Rates').getDataRange().getValues(); 
  let result = [];
  for (let i = 1; i < data.length; i++) result.push({ id: data[i][0], area: data[i][1], rate: data[i][2] });
  return responseJSON({ success: true, data: result });
}

function handleEditShipping(e) {
  const sheet = getSheet('Shipping_Rates');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === e.parameter.id_ongkir) {
      if(e.parameter.nama_area !== undefined) sheet.getRange(i+1, 2).setValue(e.parameter.nama_area);
      if(e.parameter.tarif_ongkir !== undefined) sheet.getRange(i+1, 3).setValue(e.parameter.tarif_ongkir);
      return responseJSON({ success: true, message: "Ongkir diupdate!" });
    }
  }
  return responseJSON({ success: false });
}

function handleDeleteShipping(e) {
  const sheet = getSheet('Shipping_Rates');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === e.parameter.id_ongkir) { sheet.deleteRow(i+1); return responseJSON({ success: true }); }
  }
  return responseJSON({ success: false });
}

// =========================================================================
// 6.5 CRUD ARTIKEL / BLOG
// =========================================================================
function generateSlug(text) {
  return text.toString().toLowerCase()
    .replace(/\s+/g, '-')           // Ganti spasi dengan -
    .replace(/[^\w\-]+/g, '')       // Hapus semua karakter non-word
    .replace(/\-\-+/g, '-')         // Ganti multiple - dengan single -
    .replace(/^-+/, '')             // Trim - dari awal teks
    .replace(/-+$/, '');            // Trim - dari akhir teks
}

function handleAddArticle(e) {
  const sheet = getSheet('Articles');
  const id = 'ART-' + Math.floor(Math.random() * 100000);
  const slug = generateSlug(e.parameter.judul);
  
  sheet.appendRow([
    id, e.parameter.judul, slug, e.parameter.gambar || '', e.parameter.isi_artikel || '', new Date()
  ]);
  return responseJSON({ success: true, message: "Artikel ditambahkan!" });
}

function handleGetArticles(e) {
  const data = getSheet('Articles').getDataRange().getValues(); 
  let result = [];
  for (let i = 1; i < data.length; i++) {
      result.push({ id: data[i][0], title: data[i][1], slug: data[i][2], img: data[i][3], content: data[i][4], date: data[i][5] });
  }
  return responseJSON({ success: true, data: result.reverse() });
}

function handleEditArticle(e) {
  const sheet = getSheet('Articles');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === e.parameter.id_artikel) {
      if(e.parameter.judul !== undefined) {
        sheet.getRange(i+1, 2).setValue(e.parameter.judul);
        sheet.getRange(i+1, 3).setValue(generateSlug(e.parameter.judul)); // Update slug juga
      }
      if(e.parameter.gambar !== undefined) sheet.getRange(i+1, 4).setValue(e.parameter.gambar);
      if(e.parameter.isi_artikel !== undefined) sheet.getRange(i+1, 5).setValue(e.parameter.isi_artikel);
      return responseJSON({ success: true, message: "Artikel diupdate!" });
    }
  }
  return responseJSON({ success: false });
}

function handleDeleteArticle(e) {
  const sheet = getSheet('Articles');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === e.parameter.id_artikel) { sheet.deleteRow(i+1); return responseJSON({ success: true }); }
  }
  return responseJSON({ success: false });
}

// =========================================================================
// 8. HELPER FUNCTIONS
// =========================================================================
function getSheet(sheetName) { 
  return SpreadsheetApp.openById(SCRIPT_PROP.getProperty('key')).getSheetByName(sheetName); 
}

function hashPassword(password) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password); 
  let txtHash = '';
  for (let i = 0; i < rawHash.length; i++) { 
    let hashVal = rawHash[i]; 
    if (hashVal < 0) hashVal += 256; 
    if (hashVal.toString(16).length == 1) txtHash += '0'; 
    txtHash += hashVal.toString(16); 
  }
  return txtHash;
}

function responseJSON(data) { 
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); 
}

function applyTemplate(templateStr, dataObj) {
  if (!templateStr) return ""; 
  let result = templateStr;
  for (const key in dataObj) {
    const regex = new RegExp(`\\[${key}\\]`, 'g');
    result = result.replace(regex, dataObj[key]);
  }
  return result;
}

function handleUpdateStoreSettings(e) {
  const sheet = getSheet('Users');
  const data = sheet.getDataRange().getValues();
  const userId = e.parameter.id_user_toko;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      sheet.getRange(i + 1, 6).setValue(e.parameter.nama_toko);      
      sheet.getRange(i + 1, 9).setValue(e.parameter.logo);           
      sheet.getRange(i + 1, 10).setValue(e.parameter.wa_number);     
      sheet.getRange(i + 1, 11).setValue(e.parameter.auto_reply);    
      sheet.getRange(i + 1, 12).setValue(e.parameter.qris_url);      
      sheet.getRange(i + 1, 13).setValue(e.parameter.bank_info);     
      sheet.getRange(i + 1, 14).setValue(e.parameter.gateway_token); 
      sheet.getRange(i + 1, 15).setValue(e.parameter.gateway_sender);
      sheet.getRange(i + 1, 16).setValue(e.parameter.gateway_provider);
      sheet.getRange(i + 1, 17).setValue(e.parameter.banner_url);    
      sheet.getRange(i + 1, 18).setValue(e.parameter.theme_color);   
      sheet.getRange(i + 1, 19).setValue(e.parameter.social_media);  
      sheet.getRange(i + 1, 20).setValue(e.parameter.notif_templates); 
      sheet.getRange(i + 1, 21).setValue(e.parameter.legal_settings); 
      sheet.getRange(i + 1, 22).setValue(e.parameter.pixel_fb);      
      sheet.getRange(i + 1, 23).setValue(e.parameter.pixel_tiktok);  
      sheet.getRange(i + 1, 24).setValue(e.parameter.pixel_google);  
      sheet.getRange(i + 1, 25).setValue(e.parameter.tax_settings);  
      sheet.getRange(i + 1, 26).setValue(e.parameter.secret_login_status);
      sheet.getRange(i + 1, 27).setValue(e.parameter.secret_tap_count);
      sheet.getRange(i + 1, 28).setValue(e.parameter.secret_pin);
      sheet.getRange(i + 1, 29).setValue(e.parameter.announcement_text); // Fix Marquee
      sheet.getRange(i + 1, 30).setValue(e.parameter.announcement_link); // Fix Marquee
      
      return responseJSON({ success: true, message: "Pengaturan berhasil disimpan!" });
    }
  }
  return responseJSON({ success: false, message: "User tidak ditemukan!" });
}
