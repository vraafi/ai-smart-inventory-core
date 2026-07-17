function _getInventoryColMap(headers) {
  const map = { code: -1, name: -1, category: -1, price: -1, buyPrice: -1, sellPrice: -1, branch: -1, 
                stockIn: -1, stockOut: -1, stock: -1, minStock: -1, unit: -1, status: -1 };
  if (!headers || headers.length === 0) return map;

  const normalized = headers.map(h => String(h).toLowerCase().trim());

  const rules = [
    ['stockIn',  h => h.includes('stock in') || h.includes('stok masuk') || h === 'masuk' || h === 'in' || h.includes('entrada') || h.includes('entree') || h.includes('zugang')],
    ['stockOut', h => h.includes('stock out') || h.includes('stok keluar') || h === 'keluar' || h === 'out' || h.includes('salida') || h.includes('sortie') || h.includes('abgang')],
    ['minStock', h => (h.includes('min') && (h.includes('stock') || h.includes('stok'))) || h.includes('reorder') || h.includes('safety') || h.includes('minimo')],
    ['buyPrice', h => h.includes('buy') || h.includes('beli') || h.includes('modal') || h.includes('achat') || h.includes('compra')],
    ['sellPrice', h => h.includes('sell') || h.includes('jual') || h.includes('vente') || h.includes('venta')],
    ['price',    h => h.includes('harga') || h.includes('price') || h.includes('cost') || h.includes('precio') || h.includes('prix') || h.includes('preis')],
    ['code',     h => h === 'item code' || h === 'kode barang' || h === 'sku' || h === 'code' || h === 'barcode' || h === 'upc' || h === 'ean'],
    ['name',     h => h.includes('nama') || (h.includes('name') && !h.includes('code')) || h.includes('barang') || h.includes('produk') || h.includes('product') || (h.includes('item') && !h.includes('code')) || h.includes('nombre') || h.includes('nom') || h.includes('artikel')],
    ['category', h => h.includes('kategori') || h.includes('category') || h.includes('jenis') || h.includes('tipe') || h.includes('type') || h.includes('categoria') || h.includes('categoría') || h.includes('catégorie') || h.includes('kategorie')],
    ['branch',   h => h.includes('cabang') || h.includes('branch') || h.includes('toko') || h.includes('lokasi') || h.includes('gudang') || h.includes('warehouse') || h.includes('location') || h.includes('sucursal') || h.includes('succursale') || h.includes('filiale')],
    ['unit',     h => !h.includes('harga') && !h.includes('price') && (h === 'satuan' || h === 'unit' || h === 'uom' || h.includes('unidad') || h.includes('unite') || h.includes('einheit'))],
    ['status',   h => h.includes('status') || h.includes('kondisi') || h.includes('state') || h.includes('estado') || h.includes('etat')],
    ['stock',    h => h.includes('stock') || h.includes('stok') || h.includes('qty') || h.includes('quantity') || h.includes('jumlah') || h.includes('persediaan') || h.includes('saldo') || h.includes('inventario') || h.includes('inventaire') || h.includes('bestand') || h.includes('cantidad')],
  ];

  for (let i = 0; i < normalized.length; i++) {
    const h = normalized[i];
    if (!h) continue;
    for (const [key, test] of rules) {
      if (map[key] === -1 && test(h)) {
        map[key] = i;
        break;
      }
    }
  }

  if (map.code === -1 && map.name !== -1) map.code = Math.max(0, map.name - 1);
  if (map.name === -1 && map.code !== -1) map.name = map.code + 1;
  return map;
}

const headers = ['ID', 'Item Code', 'Item Name', 'Category', 'Branch', 'Initial Stock', 'Total In', 'Total Out', 'Current Stock', 'Min Stock', 'Unit', 'Buy Price', 'Sell Price', 'Status'];
console.log(JSON.stringify(_getInventoryColMap(headers)));
