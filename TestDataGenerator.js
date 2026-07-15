/**
 * TEST DATA GENERATOR - Enterprise Scale
 * Generates realistic product data for stress-testing the AI Inventory System.
 *
 * Usage: Open Google Sheets > Menu > 📦 Inventory System > 🧪 Generate Test Data
 */

function generateTestData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '🧪 Generate Test Data',
    'Berapa jumlah produk yang ingin di-generate?\n(Rekomendasi: 1000 untuk tes besar, max 5000)',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;
  const count = parseInt(response.getResponseText()) || 1000;
  if (count > 10000) {
    ui.alert('⚠️ Warning', 'Maksimum 10,000 produk untuk menjaga performa Google Sheets.', ui.ButtonSet.OK);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Inventory');
  if (!sheet) {
    sheet = ss.getActiveSheet();
    sheet.setName('Inventory');
  }

  // Clear existing data (keep header)
  const headerRow = InventoryCore.findHeaderRow(sheet);
  if (sheet.getLastRow() > headerRow) {
    sheet.getRange(headerRow + 1, 1, sheet.getLastRow() - headerRow, 9).clearContent();
  }

  // If no header exists, create one
  if (headerRow <= 0 || sheet.getRange(1, 1).getValue() === '') {
    sheet.getRange(1, 1, 1, 9).setValues([['Branch', 'Item ID', 'Item Name', 'Category', 'Initial Stock', 'Stock In', 'Stock Out', 'Current Stock', 'Status']]);
    sheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  const startRow = headerRow + 1 || 2;

  // ═══════════════════════════════════════════════════
  // ENTERPRISE PRODUCT DATABASE (Realistic Categories)
  // ═══════════════════════════════════════════════════
  const catalog = {
    'Smartphone': [
      'iPhone 16 Pro Max 256GB', 'iPhone 16 Pro 128GB', 'iPhone 16 Plus 256GB', 'iPhone 16 128GB',
      'iPhone 15 Pro Max 512GB', 'iPhone 15 Pro 256GB', 'iPhone 15 128GB', 'iPhone SE 4 128GB',
      'Samsung Galaxy S25 Ultra 512GB', 'Samsung Galaxy S25+ 256GB', 'Samsung Galaxy S25 128GB',
      'Samsung Galaxy Z Fold 6 512GB', 'Samsung Galaxy Z Flip 6 256GB',
      'Samsung Galaxy A56 5G 128GB', 'Samsung Galaxy A36 5G 128GB', 'Samsung Galaxy A16 5G 64GB',
      'Google Pixel 9 Pro XL 256GB', 'Google Pixel 9 Pro 128GB', 'Google Pixel 9 128GB', 'Google Pixel 9a 128GB',
      'Xiaomi 15 Ultra 512GB', 'Xiaomi 15 Pro 256GB', 'Xiaomi 15 128GB',
      'Xiaomi Redmi Note 14 Pro+ 256GB', 'Xiaomi Redmi Note 14 128GB', 'Xiaomi Redmi 14C 64GB',
      'OPPO Find X8 Pro 256GB', 'OPPO Reno 13 Pro 256GB', 'OPPO Reno 13 128GB', 'OPPO A5 Pro 128GB',
      'Vivo X200 Pro 256GB', 'Vivo X200 128GB', 'Vivo V40 Pro 256GB', 'Vivo Y29 128GB',
      'OnePlus 13 256GB', 'OnePlus 13R 128GB', 'OnePlus Nord 4 128GB',
      'Realme GT 7 Pro 256GB', 'Realme 13 Pro+ 256GB', 'Realme C67 128GB',
      'Sony Xperia 1 VI 256GB', 'Nothing Phone 3 256GB', 'Asus ROG Phone 9 Pro 512GB',
      'Huawei Pura 80 Pro 256GB', 'Honor Magic 7 Pro 256GB', 'Motorola Edge 60 Pro 256GB',
      'Infinix Note 40 Pro+ 256GB', 'Tecno Phantom V2 Fold 256GB', 'ZTE Nubia Z70 Ultra 256GB'
    ],
    'Laptop': [
      'MacBook Air M4 13" 256GB', 'MacBook Air M4 15" 512GB', 'MacBook Pro M4 14" 512GB',
      'MacBook Pro M4 Pro 14" 1TB', 'MacBook Pro M4 Max 16" 1TB',
      'Dell XPS 14 i7 16GB 512GB', 'Dell XPS 16 i9 32GB 1TB', 'Dell Inspiron 14 i5 8GB 256GB',
      'Dell Latitude 7450 i7 16GB 512GB', 'Dell Precision 5690 i9 64GB 2TB',
      'HP Spectre x360 14 i7 16GB 512GB', 'HP EliteBook 860 G11 i7 16GB',
      'HP Pavilion 14 i5 8GB 256GB', 'HP Omen 16 RTX 4070 i7 16GB',
      'Lenovo ThinkPad X1 Carbon Gen 12 i7', 'Lenovo ThinkPad T14s Gen 6 i7',
      'Lenovo IdeaPad Slim 5 i5 8GB', 'Lenovo Legion Pro 5 RTX 4070 i9',
      'Lenovo Yoga 9i 14 i7 16GB OLED', 'Lenovo LOQ 15 RTX 4060 i5',
      'Asus ZenBook 14 OLED UX3405 i7', 'Asus ROG Zephyrus G16 RTX 4090',
      'Asus VivoBook 15 i5 8GB 512GB', 'Asus TUF Gaming A15 RTX 4060 R7',
      'Acer Swift Go 14 i7 16GB OLED', 'Acer Predator Helios 18 RTX 4080',
      'Acer Aspire 5 i5 8GB 256GB', 'Acer Nitro V 15 RTX 4050 i5',
      'MSI Stealth 16 Studio RTX 4070', 'MSI GF63 Thin i5 RTX 3050',
      'Samsung Galaxy Book 4 Pro 16 i7', 'Huawei MateBook X Pro i7 16GB',
      'Microsoft Surface Laptop 7 i7', 'Microsoft Surface Pro 11 i7',
      'Razer Blade 16 RTX 4090 i9', 'Framework Laptop 16 R7 32GB'
    ],
    'Tablet': [
      'iPad Pro M4 11" 256GB WiFi', 'iPad Pro M4 13" 512GB WiFi+Cell',
      'iPad Air M3 11" 128GB', 'iPad Air M3 13" 256GB',
      'iPad 10th Gen 64GB WiFi', 'iPad Mini 7 128GB WiFi',
      'Samsung Galaxy Tab S10 Ultra 256GB', 'Samsung Galaxy Tab S10+ 256GB',
      'Samsung Galaxy Tab S10 FE 128GB', 'Samsung Galaxy Tab A9+ 64GB',
      'Xiaomi Pad 7 Pro 128GB', 'Xiaomi Redmi Pad Pro 128GB',
      'Lenovo Tab P12 Pro 256GB', 'Huawei MatePad Pro 13.2 256GB',
      'OnePlus Pad 2 128GB', 'Google Pixel Tablet 2 256GB',
      'Microsoft Surface Go 4 128GB', 'Amazon Fire HD 10 Plus 64GB',
      'Realme Pad 3 Pro 128GB', 'Honor Pad V9 128GB'
    ],
    'Wearables': [
      'Apple Watch Ultra 3 49mm', 'Apple Watch Series 10 46mm GPS',
      'Apple Watch Series 10 42mm GPS', 'Apple Watch SE 3 44mm',
      'Samsung Galaxy Watch 7 Ultra 47mm', 'Samsung Galaxy Watch 7 44mm',
      'Samsung Galaxy Watch FE 40mm', 'Samsung Galaxy Ring Size 8',
      'Google Pixel Watch 3 45mm LTE', 'Google Pixel Watch 3 41mm WiFi',
      'Garmin Fenix 8 51mm Solar', 'Garmin Forerunner 965',
      'Garmin Venu 3 45mm', 'Garmin Instinct 3 Solar',
      'Fitbit Charge 6', 'Fitbit Versa 5',
      'Xiaomi Watch S4', 'Xiaomi Smart Band 9 Pro',
      'Huawei Watch GT 5 Pro 46mm', 'Amazfit T-Rex 3',
      'OnePlus Watch 3', 'Sony WF-1000XM6 Earbuds',
      'Apple AirPods Pro 3', 'Apple AirPods 4 ANC',
      'Apple AirPods Max USB-C', 'Samsung Galaxy Buds 3 Pro',
      'Google Pixel Buds Pro 2', 'Bose QuietComfort Ultra Earbuds',
      'Sony WH-1000XM6 Headphones', 'JBL Tour One M3 Headphones'
    ],
    'Accessories': [
      'Apple MagSafe Charger 15W', 'Apple 35W Dual USB-C Charger',
      'Apple Magic Keyboard with Touch ID', 'Apple Magic Mouse Black',
      'Apple Magic Trackpad Black', 'Apple Pencil Pro',
      'Apple Pencil USB-C', 'Apple AirTag 4 Pack',
      'Samsung 45W Super Fast Charger', 'Samsung S Pen Fold Edition',
      'Anker 737 GaNPrime 120W Charger', 'Anker PowerCore 26800mAh',
      'Anker Soundcore Space A40 Earbuds', 'Baseus 65W GaN Charger',
      'Ugreen Nexode 100W USB-C Charger', 'Spigen Tough Armor iPhone 16 Case',
      'OtterBox Defender Samsung S25 Case', 'Belkin 3-in-1 MagSafe Stand',
      'Logitech MX Master 3S Mouse', 'Logitech MX Keys S Keyboard',
      'Logitech C920 HD Pro Webcam', 'Razer DeathAdder V3 Mouse',
      'SanDisk 1TB microSD Extreme', 'Samsung 990 Pro 2TB NVMe SSD',
      'WD Black SN850X 1TB NVMe', 'Kingston Fury Beast 32GB DDR5',
      'Corsair Vengeance 64GB DDR5', 'Nikon Z fc 28mm Kit',
      'Canon EOS R50 18-45mm Kit', 'GoPro Hero 13 Black'
    ],
    'Monitor': [
      'LG UltraFine 5K 27" 27MD5KL', 'LG 27GP850 27" QHD 165Hz',
      'LG 34WN80C 34" UltraWide QHD', 'LG 48GQ900 48" OLED 4K 120Hz',
      'Samsung Odyssey G9 49" DQHD', 'Samsung ViewFinity S9 5K 27"',
      'Samsung Smart Monitor M8 32" 4K', 'Samsung Odyssey OLED G8 34"',
      'Dell UltraSharp U2724D 27" QHD', 'Dell UltraSharp U3224KB 32" 6K',
      'Dell S2722DGM 27" QHD 165Hz', 'Dell P2425H 24" FHD',
      'ASUS ProArt PA32UCR 32" 4K HDR', 'ASUS ROG Swift PG42UQ 42" OLED',
      'ASUS TUF Gaming VG28UQL1A 28" 4K', 'ASUS Eye Care VZ24EHF 24"',
      'BenQ PD2725U 27" 4K Designer', 'BenQ MOBIUZ EX2710U 27" 4K 144Hz',
      'Acer Predator X34GS 34" UWQHD', 'Acer Nitro XV282K 28" 4K 144Hz',
      'AOC AGON PRO AG274QZM 27" QHD', 'Gigabyte M32U 32" 4K 144Hz',
      'MSI MPG ARTYMIS 343CQR 34" UWQHD', 'ViewSonic VP3268a 32" 4K'
    ],
    'Networking': [
      'TP-Link Deco XE75 Pro WiFi 7 3-Pack', 'TP-Link Archer AXE300 WiFi 7 Router',
      'TP-Link Deco M5 Mesh WiFi 3-Pack', 'TP-Link TL-SG108 8-Port Switch',
      'ASUS RT-BE96U WiFi 7 Router', 'ASUS ZenWiFi BQ16 Pro Mesh 2-Pack',
      'Netgear Orbi 970 WiFi 7 3-Pack', 'Netgear Nighthawk RS700S WiFi 7',
      'Ubiquiti UniFi U7 Pro AP', 'Ubiquiti Dream Machine Pro Max',
      'Ubiquiti USW-Pro-48-PoE Switch', 'Linksys Velop MX6202 Mesh 2-Pack',
      'MikroTik hAP ax3 WiFi 6 Router', 'Cisco Meraki MR46 Access Point',
      'Synology RT6600ax WiFi 6 Router', 'Google Nest WiFi Pro 6E 3-Pack'
    ],
    'Storage': [
      'Synology DS1624+ 6-Bay NAS', 'Synology DS423+ 4-Bay NAS',
      'Synology DS223j 2-Bay NAS', 'QNAP TS-464 4-Bay NAS',
      'WD Red Plus 8TB NAS HDD', 'WD Red Plus 4TB NAS HDD',
      'WD Red Plus 2TB NAS HDD', 'WD Ultrastar HC580 20TB',
      'Seagate IronWolf 8TB NAS', 'Seagate IronWolf 4TB NAS',
      'Seagate Exos X20 20TB Enterprise', 'Seagate Backup Plus Hub 8TB',
      'Samsung 870 EVO 4TB SATA SSD', 'Samsung 870 EVO 1TB SATA SSD',
      'Samsung T9 Portable SSD 4TB', 'Samsung T7 Shield 2TB',
      'Crucial MX500 2TB SATA SSD', 'Kingston NV2 2TB NVMe SSD',
      'WD Black SN850X 4TB NVMe', 'SanDisk Professional G-DRIVE 12TB'
    ],
    'Printer': [
      'HP LaserJet Pro MFP 4101fdw', 'HP Color LaserJet Pro MFP M283fdw',
      'HP OfficeJet Pro 9125e All-in-One', 'HP DeskJet 4155e All-in-One',
      'HP Smart Tank 7305 All-in-One', 'Brother HL-L2460DW Laser',
      'Brother MFC-L3780CDW Color Laser', 'Brother MFC-J4535DW Inkjet',
      'Canon PIXMA G7020 MegaTank', 'Canon imageCLASS MF465dw Laser',
      'Canon PIXMA TR8620a All-in-One', 'Canon MAXIFY GX7021 MegaTank',
      'Epson EcoTank ET-4850 All-in-One', 'Epson EcoTank ET-2850 All-in-One',
      'Epson WorkForce Pro WF-C4810 Color', 'Xerox B310 Mono Laser',
      'Lexmark MC3426i Color Laser MFP', 'Samsung Xpress M2070W Mono Laser'
    ],
    'Software': [
      'Microsoft 365 Personal 1 Year', 'Microsoft 365 Family 1 Year',
      'Microsoft Office LTSC Pro Plus 2024', 'Microsoft Windows 11 Pro License',
      'Microsoft Windows 11 Home License', 'Microsoft Visio Pro 2024',
      'Microsoft Project Pro 2024', 'Adobe Creative Cloud All Apps 1 Year',
      'Adobe Photoshop 1 Year', 'Adobe Premiere Pro 1 Year',
      'Adobe Illustrator 1 Year', 'Adobe Acrobat Pro 1 Year',
      'AutoCAD 2025 License 1 Year', 'SketchUp Pro 1 Year',
      'Notion Plus Plan Annual', 'Canva Pro 1 Year',
      'Zoom Workplace Pro Annual', 'Slack Pro Plan Annual',
      'Norton 360 Deluxe 1 Year 5 Devices', 'Kaspersky Premium 1 Year 5 Devices',
      'Bitdefender Total Security 1 Year', 'NordVPN 2 Year Plan'
    ]
  };

  // Generate products
  const categories = Object.keys(catalog);
  let allProducts = [];

  // First, use all real products from catalog
  categories.forEach(cat => {
    catalog[cat].forEach(name => {
      allProducts.push({ name: name, category: cat });
    });
  });

  // If user wants more than catalog size, generate variants
  const colors = ['Black', 'White', 'Silver', 'Gold', 'Blue', 'Green', 'Red', 'Purple', 'Gray', 'Pink'];
  const regions = ['US', 'EU', 'APAC', 'Global', 'JP', 'KR'];

  while (allProducts.length < count) {
    const cat = categories[Math.floor(Math.random() * categories.length)];
    const items = catalog[cat];
    const baseName = items[Math.floor(Math.random() * items.length)];
    const color = colors[Math.floor(Math.random() * colors.length)];
    const region = regions[Math.floor(Math.random() * regions.length)];
    allProducts.push({ name: `${baseName} (${color}/${region})`, category: cat });
  }

  // Trim to exact count
  allProducts = allProducts.slice(0, count);

  // Build data rows
  const rows = [];
  const formulas = [];
  const branches = ['Pusat', 'Cabang Jakarta', 'Cabang Bandung', 'Cabang Surabaya', 'Cabang Bali'];

  allProducts.forEach((product, index) => {
    const paddedIndex = String(index + 1).padStart(5, '0');
    const catPrefix = product.category.substring(0, 3).toUpperCase();
    const itemId = `${catPrefix}-${paddedIndex}`;
    const initialStock = Math.floor(Math.random() * 500) + 10;
    const stockIn = Math.floor(Math.random() * 200);
    const stockOut = Math.floor(Math.random() * Math.min(initialStock + stockIn, 300));
    const branch = branches[Math.floor(Math.random() * branches.length)];

    const rowNum = startRow + index;
    rows.push([branch, itemId, product.name, product.category, initialStock, stockIn, stockOut, '', '']);
    formulas.push({ row: rowNum, formula: `=E${rowNum}+F${rowNum}-G${rowNum}` });
  });

  // Batch write (much faster than appendRow)
  sheet.getRange(startRow, 1, rows.length, 9).setValues(rows);

  // Set formulas for Current Stock (Column H)
  formulas.forEach(f => {
    sheet.getRange(f.row, 8).setFormula(f.formula);
  });

  // Update all statuses
  SpreadsheetApp.flush(); // Force calculation
  for (let i = startRow; i < startRow + rows.length; i++) {
    InventoryCore.updateStatusColumn(sheet, i);
  }

  // Auto-resize columns
  for (let col = 1; col <= 9; col++) {
    sheet.autoResizeColumn(col);
  }

  ui.alert('✅ Data Generated!',
    `Berhasil membuat ${count} produk enterprise!\n\n` +
    `Kategori: ${categories.length} (${categories.join(', ')})\n\n` +
    `Sekarang coba tes AI dengan mengetik:\n"Terjual 5 unit iPhone 16 Pro Max dan 3 Samsung Galaxy S25 Ultra"`,
    ui.ButtonSet.OK
  );
}
