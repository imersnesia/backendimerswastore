// =========================================================================
// ORDER ORCHESTRATOR & POS MANAGEMENT (SINGLE STORE ULTIMATE)
// =========================================================================

function handleAddOrder(e) {
  const sheet = getSheet('Orders');
  const orderId = 'ORD-' + Math.floor(Math.random() * 100000);
  
  // 1. Kalkulasi Final di Backend
  let subtotal = parseInt(e.parameter.subtotal) || 0;
  let pajak_ppn = parseInt(e.parameter.pajak_ppn) || 0;
  let biaya_layanan = parseInt(e.parameter.biaya_layanan) || 0;
  let shipping_cost = parseInt(e.parameter.ongkir_nominal) || 0;
  let discount_amount = parseInt(e.parameter.diskon_nominal) || 0;
  let grand_total = subtotal + shipping_cost - discount_amount + pajak_ppn + biaya_layanan;

  sheet.appendRow([
    orderId, e.parameter.nama_pembeli || 'Pelanggan Walk-in', grand_total, 
    e.parameter.status || 'Baru', e.parameter.detail_items, new Date(), 
    e.parameter.metode_pembayaran || 'Cash', e.parameter.uang_bayar || 0, e.parameter.kembalian || 0, 
    e.parameter.no_wa_pembeli || '', subtotal, shipping_cost, discount_amount, 
    e.parameter.kode_kupon || '', e.parameter.nama_area_ongkir || '', pajak_ppn, biaya_layanan
  ]);

  // 2. POTONG STOK PRODUK (Termasuk Varian)
  try {
     const prodSheet = getSheet('Products');
     const prodData = prodSheet.getDataRange().getValues();
     let orderItems = [];
     
     try { orderItems = JSON.parse(e.parameter.detail_items_raw || '[]'); } catch(err) {}

     if (orderItems.length > 0) {
        for (let idx = 0; idx < orderItems.length; idx++) {
           const item = orderItems[idx];
           for (let j = 1; j < prodData.length; j++) {
              if (prodData[j][0] === item.id) {
                 // Potong Stok Varian (Index varian di Single Store = 9)
                 if (item.variant_id) {
                    let varsArray = [];
                    try { varsArray = JSON.parse(prodData[j][9] || '[]'); } catch(e) {}
                    let isUpdated = false;
                    for (let v = 0; v < varsArray.length; v++) {
                       if (varsArray[v].id === item.variant_id && varsArray[v].stok !== '') {
                          let s = parseInt(varsArray[v].stok);
                          if (!isNaN(s)) { varsArray[v].stok = Math.max(0, s - item.qty); isUpdated = true; }
                       }
                    }
                    if (isUpdated) prodSheet.getRange(j + 1, 10).setValue(JSON.stringify(varsArray));
                 } else {
                 // Potong Stok Biasa (Index stok di Single Store = 3)
                    let s = prodData[j][3];
                    if (s !== '') { 
                      let sInt = parseInt(s); 
                      if (!isNaN(sInt)) prodSheet.getRange(j + 1, 4).setValue(Math.max(0, sInt - item.qty)); 
                    }
                 }
                 break; 
              }
           }
        }
     }
  } catch(e) { console.error("Gagal potong stok: ", e); }

  // 3. TRIGGER WA GATEWAY (Ambil data admin langsung dari baris ke-2 sheet Users)
  let waSent = false;
  try {
    const userSheet = getSheet('Users').getDataRange().getValues();
    if (userSheet.length > 1 && e.parameter.no_wa_pembeli) {
      let token = userSheet[1][13]; // Kolom gateway_token (index 13)
      let sender = userSheet[1][14]; // Kolom gateway_sender
      let provider = userSheet[1][15] || 'fonnte'; // Kolom gateway_provider
      
      let checkoutTemplate = '';
      try { 
        let tplJSON = JSON.parse(userSheet[1][19] || '{}'); // Kolom notif_templates (index 19)
        checkoutTemplate = tplJSON.checkout_msg || ''; 
      } catch(err){}

      if (token) {
        const defaultCheckout = `Halo *[NAMA_PEMBELI]*,\nTerima kasih telah berbelanja di toko kami.\n\nOrder ID: *[ORDER_ID]*\nTotal Belanja: *[GRAND_TOTAL]*\n\nKami akan segera memproses pesanan Anda.`;
        
        let rincianOrderTxt = '';
        try { 
          let arr = JSON.parse(e.parameter.detail_items || '[]');
          rincianOrderTxt = arr.map(a => '- ' + a).join('\n'); 
        } catch(err) { rincianOrderTxt = e.parameter.detail_items; }

        const tagDataCheckout = {
           'NAMA_PEMBELI': e.parameter.nama_pembeli || 'Kak',
           'ORDER_ID': orderId,
           'SUBTOTAL': 'Rp ' + subtotal,
           'ONGKIR': 'Rp ' + shipping_cost,
           'DISKON': '-Rp ' + discount_amount,
           'PPN': 'Rp ' + pajak_ppn,
           'BIAYA_LAYANAN': 'Rp ' + biaya_layanan,
           'GRAND_TOTAL': 'Rp ' + grand_total,
           'METODE_BAYAR': e.parameter.metode_pembayaran || 'Cash',
           'KURIR': e.parameter.nama_area_ongkir || '-',
           'RINCIAN_ORDER': rincianOrderTxt
        };
        
        let finalMsg = applyTemplate(checkoutTemplate || defaultCheckout, tagDataCheckout);
        
        // Panggil WA Gateway
        if (typeof sendWhatsApp !== 'undefined') {
          waSent = sendWhatsApp(e.parameter.no_wa_pembeli, finalMsg, token, sender, provider);
        }
      }
    }
  } catch(err) { console.error("Gagal kirim WA di Order: ", err); }

  return responseJSON({ success: true, message: "Pesanan disimpan!", order_id: orderId, grand_total: grand_total, wa_sent: waSent });
}

function handleGetOrders(e) {
  const data = getSheet('Orders').getDataRange().getValues(); 
  let result = [];
  for (let i = 1; i < data.length; i++) {
      result.push({ 
        id: data[i][0], buyer: data[i][1], total: data[i][2], status: data[i][3], 
        items: data[i][4], time: data[i][5], method: data[i][6], pay: data[i][7], 
        change: data[i][8], wa: data[i][9], subtotal: data[i][10], ongkir: data[i][11], 
        diskon: data[i][12], kode_kupon: data[i][13], area_ongkir: data[i][14],
        pajak_ppn: data[i][15] || 0, biaya_layanan: data[i][16] || 0 
      });
  }
  return responseJSON({ success: true, data: result.reverse() }); 
}

function handleUpdateOrderStatus(e) {
  const sheet = getSheet('Orders');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === e.parameter.id_order) { 
      sheet.getRange(i + 1, 4).setValue(e.parameter.new_status); 
      return responseJSON({ success: true, message: "Status order diupdate!" }); 
    }
  }
  return responseJSON({ success: false });
}

function handleDeleteOrder(e) {
  const sheet = getSheet('Orders');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === e.parameter.id_order) { 
      sheet.deleteRow(i + 1); return responseJSON({ success: true, message: "Pesanan dihapus!" }); 
    }
  }
  return responseJSON({ success: false });
}
