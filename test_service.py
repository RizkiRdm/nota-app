import os
import sys
from datetime import datetime

# Tambahkan direktori saat ini ke path agar modul bisa diimpor
# Ini penting saat menjalankan file dari root project
sys.path.append(os.path.dirname(os.path.abspath(__file__))) 

from excel_service import (
    read_master_stock, create_master_stock, update_master_stock, 
    delete_master_stock, write_sales_transaction, FILE_PATH, 
    SHEET_MASTER_STOK, SHEET_JURNAL_PENJUALAN, _ensure_file_and_sheets
)
from models import MasterStockProduct, HargaJual, JurnalPenjualan

# --- CONFIG ---
TEST_PRODUCT_NAME = "Kopi Bubuk ABC"
UPDATED_PRODUCT_NAME = "Kopi Bubuk ABC Extra"

def cleanup_and_setup():
    """Membersihkan file lama dan memastikan struktur dasar ada."""
    if os.path.exists(FILE_PATH):
        os.remove(FILE_PATH)
        print(f"🧹 File lama {FILE_PATH} dihapus.")
    
    # Panggil fungsi yang memastikan file dan sheet ada
    _ensure_file_and_sheets()
    print("✅ Setup awal: myPos.xlsx berhasil dibuat dengan headers.")

def test_master_stock_crud():
    """Menguji operasi Create, Read, Update, Delete untuk Master Stok."""
    
    print("\n--- TEST: CRUD Master Stok ---")

    # 1. CREATE (C)
    print("1. Menguji CREATE...")
    try:
        new_product = MasterStockProduct(
            nama_produk=TEST_PRODUCT_NAME,
            satuan_beli="Pack Besar",
            isi_per_satuan_beli=12,
            kategori="Minuman",
            satuan_unit_dasar="Bungkus",
            harga_jual=HargaJual(bungkus=2500.0, seduh=4000.0)
        )
        create_master_stock(new_product)
        print(f"   [SUKSES] Produk '{TEST_PRODUCT_NAME}' dibuat.")
    except Exception as e:
        print(f"   [GAGAL] CREATE: {e}")
        return

    # 2. READ (R)
    print("2. Menguji READ...")
    products = read_master_stock()
    if len(products) == 1 and products[0].nama_produk == TEST_PRODUCT_NAME:
        print(f"   [SUKSES] Produk dibaca: {products[0].nama_produk}. Harga display: {products[0].price_display}")
    else:
        print(f"   [GAGAL] READ: Jumlah produk salah ({len(products)}).")
        return

    # 3. UPDATE (U)
    print("3. Menguji UPDATE...")
    try:
        updated_product = MasterStockProduct(
            nama_produk=UPDATED_PRODUCT_NAME, # Mengubah nama produk
            satuan_beli="Pack Besar",
            isi_per_satuan_beli=12,
            kategori="Minuman Panas", # Mengubah kategori
            satuan_unit_dasar="Bungkus",
            harga_jual=HargaJual(bungkus=3000.0, seduh=5000.0) # Mengubah harga
        )
        # Parameter: Nama lama, Objek baru
        update_master_stock(TEST_PRODUCT_NAME, updated_product)
        
        # Baca kembali untuk verifikasi
        updated_check = read_master_stock()
        if updated_check[0].nama_produk == UPDATED_PRODUCT_NAME and updated_check[0].kategori == "Minuman Panas":
             print(f"   [SUKSES] Produk berhasil diupdate menjadi '{UPDATED_PRODUCT_NAME}'.")
        else:
             print("   [GAGAL] UPDATE: Verifikasi gagal.")
             return
    except Exception as e:
        print(f"   [GAGAL] UPDATE: {e}")
        return

    # 4. DELETE (D)
    print("4. Menguji DELETE...")
    try:
        delete_master_stock(UPDATED_PRODUCT_NAME)
        products_after_delete = read_master_stock()
        if len(products_after_delete) == 0:
            print(f"   [SUKSES] Produk '{UPDATED_PRODUCT_NAME}' berhasil dihapus. Stok kosong.")
        else:
            print("   [GAGAL] DELETE: Produk masih ada.")
    except Exception as e:
        print(f"   [GAGAL] DELETE: {e}")


def test_jurnal_write():
    """Menguji penulisan transaksi Jurnal Penjualan."""
    print("\n--- TEST: Write Jurnal Penjualan ---")
    
    # Pastikan ada data untuk transaksi
    create_master_stock(MasterStockProduct(
        nama_produk="Air Mineral", satuan_beli="Karton", isi_per_satuan_beli=24,
        kategori="Minuman", satuan_unit_dasar="Botol", harga_jual=HargaJual(bungkus=3000.0)
    ))

    # Data transaksi penjualan
    transactions = [
        JurnalPenjualan(nama_produk="Air Mineral", jumlah_jual=5, total_harga_jual=15000.0, catatan="Pelanggan A"),
        JurnalPenjualan(nama_produk="Air Mineral", jumlah_jual=10, total_harga_jual=30000.0, catatan="Pelanggan B"),
    ]

    try:
        write_sales_transaction(transactions)
        print("   [SUKSES] Dua transaksi penjualan berhasil dicatat.")
    except Exception as e:
        print(f"   [GAGAL] Write Jurnal: {e}")
        return

# --- MAIN EXECUTION ---
if __name__ == "__main__":
    cleanup_and_setup()
    test_master_stock_crud()
    test_jurnal_write()
    print("\n==================================")
    print(f"⭐ SCRIPT SELESAI. Cek file {FILE_PATH} untuk verifikasi manual.")
    print("==================================")