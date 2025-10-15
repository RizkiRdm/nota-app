from fastapi import FastAPI, Request, Form, Depends
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from pydantic import ValidationError
from typing import Optional
from contextlib import asynccontextmanager # BARU: Untuk Lifespan
from datetime import datetime

import excel_service
from models import JurnalPembelian, MasterStockProduct, HargaJual, JurnalPenjualan 

# --- 1. LIFESPAN HANDLER (Menggantikan @app.on_event) ---
@asynccontextmanager
async def lifespan(app: FastAPI):
    """Jalankan saat server dimulai untuk memastikan file Excel ada."""
    print("🚀 Memulai server. Memeriksa file Excel...")
    try:
        excel_service._ensure_file_and_sheets()
        print("✅ File Excel sudah siap.")
    except Exception as e:
        print(f"❌ ERROR saat inisialisasi Excel: {e}. Lanjut menjalankan server.")
    
    # Yield untuk memberitahu FastAPI bahwa startup selesai, server bisa menerima request
    yield
    
    # Kode setelah yield akan berjalan saat shutdown (opsional)
    print("🛑 Server dimatikan.")

# --- 2. INISIALISASI ---

app = FastAPI(
    title="MyPOS - HTMX App", 
    version="1.0", 
    lifespan=lifespan
)

# Setup Jinja2 Templates (Mengarah ke folder 'templates')
templates = Jinja2Templates(directory="templates")

# --- 3. ROUTES/ENDPOINTS ---

# --- A. HOME / DASHBOARD ---
@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    """Menampilkan halaman utama/dashboard."""
    try:
        products = excel_service.read_master_stock() 
    except Exception:
        products = [] 

    return templates.TemplateResponse(
        "home.html", 
        {"request": request, "products": products, "title": "Dashboard Utama"}
    )

# --- B. MASTER STOK (CRUD) ---

@app.get("/master-stok", response_class=HTMLResponse)
async def list_master_stok(request: Request, error: Optional[str] = None):
    """Menampilkan daftar semua Master Stok."""
    try:
        products = excel_service.read_master_stock()
    except Exception as e:
        return templates.TemplateResponse(
            "master_stok.html", 
            {"request": request, "products": [], "error": str(e), "title": "Master Stok"}
        )
        
    return templates.TemplateResponse(
        "master_stok.html", 
        {"request": request, "products": products, "error": error, "title": "Master Stok"}
    )
    
# --- B.1. Create/Add New Product (POST Endpoint) ---
@app.post("/master-stok", response_class=RedirectResponse, status_code=303)
async def create_product(
    request: Request,
    nama_produk: str = Form(...),
    satuan_beli: str = Form(...),
    isi_per_satuan_beli: int = Form(...),
    kategori: str = Form(...),
    satuan_unit_dasar: str = Form(...),
    harga_jual_bungkus: Optional[float] = Form(None),
    harga_jual_batang: Optional[float] = Form(None),
    harga_jual_mentah: Optional[float] = Form(None),
    harga_jual_seduh: Optional[float] = Form(None),
    harga_jual_rebus: Optional[float] = Form(None),
    harga_jual_rebus_telur: Optional[float] = Form(None),
):
    try:
        harga_jual_data = HargaJual(
            bungkus=harga_jual_bungkus,
            batang=harga_jual_batang,
            mentah=harga_jual_mentah,
            seduh=harga_jual_seduh,
            rebus=harga_jual_rebus,
            rebus_telur=harga_jual_rebus_telur,
        )

        new_product = MasterStockProduct(
            nama_produk=nama_produk,
            satuan_beli=satuan_beli,
            isi_per_satuan_beli=isi_per_satuan_beli,
            kategori=kategori,
            satuan_unit_dasar=satuan_unit_dasar,
            harga_jual=harga_jual_data
        )

        excel_service.create_master_stock(new_product)
        
    except (ValueError, ValidationError) as e:
        import urllib.parse
        error_msg = urllib.parse.quote(f"Error Validasi/Data: {e}")
        return RedirectResponse(
            url=f"/master-stok?error={error_msg}", 
            status_code=303
        )
    except Exception as e:
        import urllib.parse
        error_msg = urllib.parse.quote(f"Error tak terduga: {e}")
        return RedirectResponse(
            url=f"/master-stok?error={error_msg}", 
            status_code=303
        )

    return RedirectResponse(url="/master-stok", status_code=303)

# --- B.2. Delete Product (POST Endpoint menggunakan HTMX) ---
@app.post("/master-stok/delete/{nama_produk}", response_class=HTMLResponse)
async def delete_product(request: Request, nama_produk: str):
    """Menghapus produk, khusus untuk HTMX (kembalikan empty response 200)."""
    try:
        excel_service.delete_master_stock(nama_produk)
        return HTMLResponse(status_code=200)

    except ValueError as e:
        return HTMLResponse(content=f"<div class='text-red-500'>Error Hapus: {e}</div>", status_code=400)

# --- C. JURNAL PENJUALAN ---

@app.get("/input-penjualan", response_class=HTMLResponse)
async def sales_input_page(request: Request, error: Optional[str] = None):
    """Menampilkan halaman input penjualan dengan list produk master."""
    try:
        products = excel_service.read_master_stock()
    except Exception as e:
        products = []
        error = f"Gagal memuat Master Stok: {e}"
        
    return templates.TemplateResponse(
        "sales_input.html", 
        {"request": request, "title": "Input Penjualan", "products": products, "error": error}
    )
    
@app.post("/input-penjualan/add-item", response_class=HTMLResponse)
async def add_sales_item(request: Request, product_name: str = Form(..., alias="nama_produk_select")):
    """Endpoint HTMX untuk mengembalikan baris input penjualan baru."""
    
    product_data_pair = excel_service.get_product_by_name(product_name)
    
    if not product_data_pair:
        return HTMLResponse(content="<div class='text-red-500'>Produk tidak ditemukan.</div>", status_code=404)
    
    product, _ = product_data_pair
    
    default_price = 0.0
    default_unit = ""
    for unit, price in product.harga_jual.model_dump().items():
        if price is not None and price > 0:
            default_price = price
            default_unit = unit.replace('_', ' ').capitalize()
            break
            
    return templates.TemplateResponse(
        "sales_item_row.html", 
        {
            "request": request,
            "product": product, 
            "default_price": default_price,
            "default_unit": default_unit,
            "index": datetime.now().strftime('%Y%m%d%H%M%S%f')
        }
    )
    
@app.post("/input-penjualan", response_class=RedirectResponse, status_code=303)
async def submit_sales_transaction(request: Request):
    """Menerima dan menyimpan semua transaksi penjualan dari form multi-entry."""
    form_data = await request.form()
    
    transaction_items = []
    catatan = form_data.get('catatan')
    
    unique_indices = set()
    for key in form_data.keys():
        if key.startswith('item_'):
            parts = key.split('_')
            if len(parts) >= 3:
                unique_indices.add(parts[1])
                
    if not unique_indices:
        return RedirectResponse(url="/input-penjualan?error=Tidak ada item yang dimasukkan.", status_code=303)

    try:
        transactions_to_write: list[JurnalPenjualan] = []
        
        for index in unique_indices:
            nama_produk = form_data.get(f'item_{index}_nama_produk')
            jumlah_jual = int(form_data.get(f'item_{index}_jumlah_jual'))
            harga_jual_unit = float(form_data.get(f'item_{index}_harga_jual_unit'))
            
            total_harga_jual = jumlah_jual * harga_jual_unit
            
            transaction = JurnalPenjualan(
                nama_produk=nama_produk,
                jumlah_jual=jumlah_jual,
                total_harga_jual=total_harga_jual,
            )
            transactions_to_write.append(transaction)

        if transactions_to_write:
             # Tambahkan catatan ke transaksi pertama (satu catatan untuk satu grup transaksi)
             transactions_to_write[0].catatan = catatan

        excel_service.write_sales_transaction(transactions_to_write)

    except (ValueError, ValidationError) as e:
        import urllib.parse
        error_msg = urllib.parse.quote(f"Error Validasi/Data: {e}")
        return RedirectResponse(url=f"/input-penjualan?error={error_msg}", status_code=303)

    return RedirectResponse(url="/input-penjualan?success=Transaksi Penjualan berhasil dicatat.", status_code=303)

# --- D. JURNAL PEMBELIAN (TO BE IMPLEMENTED) ---
@app.get("/input-pembelian", response_class=HTMLResponse)
async def purchase_input_page(request: Request):
    return templates.TemplateResponse(
        "home.html", # Tempatkan template input pembelian di sini
        {"request": request, "title": "Input Pembelian"}
    )
    
@app.post("/input-pembelian/add-item", response_class=HTMLResponse)
async def add_purchase_item(request: Request, product_name: str = Form(..., alias="nama_produk_select")):
    """Endpoint HTMX untuk mengembalikan baris input pembelian baru."""
    
    product_data_pair = excel_service.get_product_by_name(product_name)
    
    if not product_data_pair:
        return HTMLResponse(content="<div class='text-red-500'>Produk tidak ditemukan.</div>", status_code=404)
    
    product, _ = product_data_pair
    
    # Ambil satuan beli dan isi per satuan beli dari Master Stok
    satuan_beli = product.satuan_beli
    
    return templates.TemplateResponse(
        "purchase_item_row.html", 
        {
            "request": request,
            "product": product, 
            "satuan_beli": satuan_beli,
            "index": datetime.now().strftime('%Y%m%d%H%M%S%f')
        }
    )
    
@app.post("/input-pembelian", response_class=RedirectResponse, status_code=303)
async def submit_purchase_transaction(request: Request):
    """Menerima dan menyimpan semua transaksi pembelian dari form multi-entry."""
    form_data = await request.form()
    
    unique_indices = set()
    for key in form_data.keys():
        if key.startswith('item_'):
            parts = key.split('_')
            if len(parts) >= 3:
                unique_indices.add(parts[1])
                
    if not unique_indices:
        return RedirectResponse(url="/input-pembelian?error=Tidak ada item yang dimasukkan.", status_code=303)

    try:
        transactions_to_write: list[excel_service.JurnalPembelian] = []
        
        for index in unique_indices:
            nama_produk = form_data.get(f'item_{index}_nama_produk')
            jumlah_beli = int(form_data.get(f'item_{index}_jumlah_beli'))
            satuan_beli = form_data.get(f'item_{index}_satuan_beli')
            total_harga_beli = float(form_data.get(f'item_{index}_total_harga_beli'))
            
            # Hitung Harga Modal Baru (Harga Beli Per Satuan)
            # Harga Beli Per Satuan = Total Harga Beli / Jumlah Beli
            harga_modal_baru = round(total_harga_beli / jumlah_beli, 2) 
            
            # 1. Update Harga Modal di Master Stok
            # Kita perlu membuat fungsi baru di excel_service
            excel_service.update_master_stock_cost_price(nama_produk, harga_modal_baru)
            
            # 2. Buat model JurnalPembelian
            transaction = JurnalPembelian(
                nama_produk=nama_produk,
                jumlah_beli=jumlah_beli,
                satuan_beli=satuan_beli,
                total_harga_beli=total_harga_beli,
                # Catatan tidak diimplementasikan di sini, tapi bisa ditambahkan
            )
            transactions_to_write.append(transaction)

        # 3. Simpan semua transaksi ke Jurnal Pembelian
        excel_service.write_purchase_transaction(transactions_to_write)

    except (ValueError, ValidationError) as e:
        import urllib.parse
        error_msg = urllib.parse.quote(f"Error Validasi/Data: {e}")
        return RedirectResponse(url=f"/input-pembelian?error={error_msg}", status_code=303)

    return RedirectResponse(url="/input-pembelian?success=Transaksi Pembelian berhasil dicatat dan modal diperbarui.", status_code=303)
