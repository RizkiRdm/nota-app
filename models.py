from pydantic import BaseModel, Field
from typing import Optional, List 

# --- MODEL DATA DARI EXCEL (MASTER STOK) ---

class ProductKey(BaseModel):
    """Model ringan untuk mengidentifikasi produk yang akan dihapus atau dicari."""
    nama_produk: str = Field(..., description="Nama produk yang akan dioperasi.")

class HargaJual(BaseModel):
    """Model untuk berbagai jenis harga jual per unit dasar."""
    # Menggunakan float untuk harga, acceptable untuk MVP
    bungkus: Optional[float] = None
    batang: Optional[float] = None
    mentah: Optional[float] = None
    seduh: Optional[float] = None
    rebus: Optional[float] = None
    rebus_telur: Optional[float] = None

class MasterStockProduct(BaseModel):
    """Struktur data untuk Master Stok (Sheet_1_Master_Stok)."""
    nama_produk: str = Field(..., description="Nama unik produk.") 
    satuan_beli: str = Field(..., description="Satuan pembelian utama (misal: Karton, Tray).") 
    isi_per_satuan_beli: int = Field(..., description="Jumlah unit dasar per satuan beli.") 
    kategori: str
    satuan_unit_dasar: str = Field(..., description="Satuan terkecil produk (misal: Bungkus, Sachet).") 
    harga_jual: HargaJual 
    
    @property
    def price_display(self) -> str:
        """Mengambil harga jual pertama yang tersedia untuk ditampilkan di card."""
        # Gunakan self.harga_jual.model_dump() untuk Pydantic V2
        for key, value in self.harga_jual.model_dump().items():
            if value is not None and value > 0:
                # Mengganti '_' dengan spasi untuk tampilan yang lebih baik
                display_key = key.replace('_', ' ').capitalize()
                return f"Rp{value:,.0f} ({display_key})"
        return "N/A"

# --- MODEL DATA UNTUK JURNAL TRANSAKSI ---

class JurnalPenjualan(BaseModel):
    """Struktur data untuk Jurnal Penjualan (Sheet_2_Jurnal_Penjualan)."""
    timestamp: Optional[str] = None
    nama_produk: str
    jumlah_jual: int = Field(..., gt=0)
    total_harga_jual: float = Field(..., gt=0)
    catatan: Optional[str] = None

class JurnalPembelian(BaseModel):
    """Struktur data untuk Jurnal Pembelian (Sheet_3_Jurnal_Pembelian)."""
    timestamp: Optional[str] = None
    nama_produk: str
    jumlah_beli: int = Field(..., gt=0)
    satuan_beli: str
    total_harga_beli: float = Field(..., gt=0)

# --- MODEL UNTUK INPUT FORM (FORM PENJUALAN MULTI-ENTRY) ---

class SalesItemInput(BaseModel):
    """Model untuk satu baris item input di form Penjualan."""
    nama_produk: str
    jumlah_jual: int = Field(gt=0)
    harga_jual_unit: float = Field(ge=0) # Harga yang editable per unit
    
class SalesFormInput(BaseModel):
    """Model untuk seluruh data yang disubmit dari form Penjualan."""
    items: list[SalesItemInput] # Menggunakan list[type] (Python 3.9+)
    catatan: Optional[str] = None