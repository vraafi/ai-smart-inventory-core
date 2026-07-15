# Workspace Rules: AI Smart Inventory

## Web App Context Rule (Spreadsheet Access)
**Constraint:** This project runs primarily as a Google Apps Script Web App (`doGet`/`doPost`). In this context, `SpreadsheetApp.getActiveSpreadsheet()` will ALWAYS return `null` and cause crashes.
**Rule:** ALWAYS use the custom wrapper `_getSpreadsheet()` to access the spreadsheet. Never use `getActiveSpreadsheet()`.

## Flexible Column Mapping Rule
**Constraint:** Users may change column headers slightly (e.g., from "Stock" to "Jumlah Stok"). 
**Rule:** When updating or creating column detection logic (like `_getInventoryColMap`), always use `.includes()` (e.g., `h.includes("stok")`) rather than strict equality (`h === "stok"`) to ensure the system is resilient to minor header changes.

## Logic-to-Prompt Parity Rule
**Constraint:** In an LLM-driven architecture, writing backend logic to handle a specific state (e.g., `if (parsed.item_code === "AMBIGUOUS")`) is useless if the LLM prompt is not instructed to produce that state.
**Rule:** Whenever adding a new handling condition for the AI's output in the codebase, you MUST simultaneously add an explicit instruction (Rule) to the `prompt` string in `Agent.js` telling the AI exactly when and how to output that condition.

## Library Wrapper Exposure Rule
**Constraint:** The system uses a Google Apps Script Library architecture. The core logic is in the `SmartInventoryCore` library, but the user interacts with a bound script (`Stub.gs`). Functions in the library are invisible to the user's Apps Script Editor dropdown and Google Sheets Macro menus.
**Rule:** Whenever you create a new function in the core library that the user needs to execute manually (such as a test function `runAllTests` or a new UI menu action), you MUST explicitly instruct the user to copy-paste a wrapper function into their `Stub.gs` file (e.g., `function runAllTests() { SmartInventoryCore.runAllTests(); }`).

## Resilient API Parsing Rule
**Constraint:** External APIs (especially free or high-traffic LLM providers) frequently return HTTP 500/524 errors with plain-text or HTML bodies instead of JSON.
**Rule:** NEVER use `JSON.parse(response)` directly on HTTP responses. You MUST wrap it in a `try/catch` block (e.g., `_safeParseJSON`) that catches `SyntaxError` and throws a descriptive `Error` stating that the API returned an invalid JSON response. This ensures the error bubbles up correctly to the fallback logic (e.g., `_callAIAuto`) instead of crashing the parser.

## Full Automation Rule
**Constraint:** The AI has full local file system access and the ability to discover and modify other related Apps Script projects (like the Bound Script / `Src`).
**Rule:** Do NOT ask the user to manually copy-paste code or wrappers. If the code exists in a local directory (e.g., `AI_Smart_Inventory_System`), use your tools to navigate there, modify the code automatically, and push the changes via `clasp`.

## UX Accessibility Rule
**Constraint:** The primary interface for the user is the Google Sheets UI, not the Apps Script IDE. Users may not know how to run functions from the code editor dropdown.
**Rule:** Whenever creating a new utility, test, or debugging function that the user is expected to run, you MUST add it to the Google Sheets Custom Menu (in `Setup.js` `_buildFullMenu`) so it is accessible via the `Inventory System` menu directly inside their spreadsheet.

## Apps Script Library Versioning Rule
**Constraint:** When debugging or actively developing a Google Apps Script Library (`AI_Smart_Inventory_Core`) that is used by a bound script (`AI_Smart_Inventory_System`), the bound script will ignore all new code pushes unless its `appsscript.json` dependency is configured correctly.
**Rule:** Before testing any changes made to a library, you MUST ensure the bound script's `appsscript.json` has `"developmentMode": true` and `"version": "HEAD"` for that library. Without this, you will be debugging a ghost/cached version of the code and your fixes will appear to fail.

## Apps Script UI Logging Rule
**Constraint:** `Logger.log()` and `console.log()` outputs are invisible to users who execute scripts via the Google Sheets Custom Menu. The script will simply show "Running script..." and "Finished script" without providing any feedback or results.
**Rule:** When writing a function intended to be executed from the Google Sheets Custom Menu (e.g., tests, debuggers, or reports), DO NOT rely solely on `Logger.log()`. You MUST accumulate the output text into a variable and display it to the user at the end of the execution using `SpreadsheetApp.getUi().alert("Title", outputString, SpreadsheetApp.getUi().ButtonSet.OK)` so the user can read the results.

## Background Trigger Exposure Rule
**Constraint:** Google Apps Script Time-driven triggers created via `ScriptApp.newTrigger("functionName")` always look for `functionName` in the bound script (`Stub.js`), NOT in the library (`SmartInventoryCore`). 
**Rule:** When creating or relying on ANY background triggers (e.g., `pollEmails`, `pollTelegram`), you MUST ensure that their corresponding wrapper functions exist in `Stub.js`. If the wrapper is missing, the trigger will fail silently in the background indefinitely.

## Email Thread Subject Matching Rule
**Constraint:** When a user replies to an automated email sent by the system, the subject line changes (e.g., adding `Re:` and inheriting the system's outbound subject). If the incoming email parser only searches for a specific inbound keyword, it will fail to read the user's replies to its own messages.
**Rule:** When writing or modifying email polling logic (`GmailApp.search`), you MUST ensure the search query includes BOTH the expected inbound subject keyword AND the outbound subject used by the system itself (e.g., `is:unread (subject:"Inbound Keyword" OR subject:"Outbound Subject")`). This guarantees that conversational replies are caught by the parser.

## Stateful vs Stateless Interaction Rule
**Constraint:** In Google Apps Script, maintaining full chat history for LLM context is expensive and slow.
**Rule:** Use a **Stateless** approach for all standard inventory commands. If an instruction is incomplete (`AMBIGUOUS`), the system must force the user to provide a single, fully-contextualized command in their next message. 
**Exception:** Use a **Stateful** approach (via `CacheService` saving `STATE_CHATID`) ONLY for highly destructive or complex multi-step wizards like `/wipe`, where maintaining strict, temporary memory (e.g., options a/b/c/d) is safer than relying on the LLM to re-evaluate the context.

## JSON Generation Safeguard Rule
**Constraint:** LLMs (especially smaller, faster models like Gemini Flash or Llama3) will break out of forced JSON constraints and return conversational error messages if the prompt's instructions conflict with the required JSON schema (e.g., asking to parse a command that doesn't fit the required `type` enum).
**Rule:** When designing prompts that require strict JSON output, you MUST provide explicit fallback values or hardcoded mappings for every required field in the JSON schema for every possible edge case (e.g., explicitly stating `type="IN", quantity=0` for `/onboarding` commands) to prevent the model from breaking character and crashing the JSON parser.

## Prompt Rule Conflict (Hierarchy)
**Constraint:** LLMs read prompts holistically. If an early rule (e.g., Rule 4: "Set type to IN for /onboarding") conflicts with a later definitional rule (e.g., Rule 6: "type IN is strictly for receiving goods"), the LLM will often prioritize the broader definition and ignore the specific override, leading to silent failures like outputting `UNKNOWN`.
**Rule:** Whenever you instruct an LLM to force a specific enum value (like `type="IN"`), you MUST ensure that the dictionary/definition of that enum explicitly permits your use case. Do not rely on isolated rule overrides; update the core definitions to prevent internal logical conflicts within the prompt.

## Gemini Payload Strictness Rule
**Constraint:** The Gemini REST API (v1/v1beta) is extremely strict about the structure of the `contents` array. Providing multiple `{text: "..."}` objects inside a single `parts` array without explicitly defined roles or multi-turn structures will result in a `400 Bad Request` that can crash the webhook/polling loop.
**Rule:** When building payloads for `generativelanguage.googleapis.com` (Gemini API), if you need to pass both a System Prompt and a User Prompt without using the new root-level `systemInstruction` field, you MUST concatenate them into a single string (`System: [sys_prompt]\n\n[user_prompt]`) and pass it as a SINGLE object in the `parts` array (e.g., `parts: [{ text: combinedText }]`). Do not push multiple text objects into the same part array.

## JSON Object Enforcement Rule (Groq/OpenAI)
**Constraint:** APIs like OpenAI and Groq support `response_format: { type: "json_object" }` to guarantee valid JSON. However, if this flag is enabled, the API strictly requires the output to be a JSON Object (starting with `{`). If your prompt instructs the model to output a JSON Array (`[`), it will cause a fatal HTTP 400 Bad Request error.
**Rule:** When designing prompts for systems that use `json_object` response formats, you MUST explicitly instruct the LLM to output a JSON Object. If you need an array of data, wrap it inside a root object (e.g., `{"transactions": [...]}`). Never instruct the AI to output a raw JSON Array.

## V8 Engine Const Enforcement Rule
**Constraint:** Google Apps Script uses the V8 engine and strictly enforces ES6 standards. Reassigning a `const` variable will immediately crash the execution with `TypeError: Assignment to constant variable`.
**Rule:** When writing or modifying logic where a variable might be assigned a default fallback or a newly created object (e.g., searching for an existing database row, and creating a new one if it doesn't exist), you MUST declare the variable using `let`, NOT `const`. Always double-check variable declarations before reassigning them later in the block.

## Edge-Case Validation Immunity Rule
**Constraint:** Standard backend validation logic (e.g., checking for `quantity > 0` or preventing duplicate names) will indiscriminately block LLM-driven edge cases (e.g., `/onboarding` needing `quantity: 0`, or `/wipe` deleting everything).
**Rule:** When designing backend guardrails for an LLM-driven system, you MUST explicitly carve out exception paths for special LLM commands. Never assume that standard validation applies to all LLM intents. Before throwing a validation error, check if the LLM intentionally triggered an edge-case flag (like `item_new: true`) that justifies bypassing the guardrail.

## Gmail Automation Tracking Rule
**Constraint:** NEVER rely on `is:unread` or `msg.isUnread()` as the sole mechanism to determine if an email has been processed. Gmail filters or client behaviors frequently mark emails as read automatically, causing the script to miss them.
**Solution:** Always use `PropertiesService` to store a rotating list of processed `msg.getId()` strings. 
**Implementation Pattern:**
1. Search using `in:anywhere` without `is:unread` constraints.
2. Check `if (processedIds.includes(msg.getId())) return;`
3. After processing, append `msg.getId()` to the `PropertiesService` string (keep the last 1000 chars to avoid quota limits).

## Aturan Modifikasi Google Apps Script (clasp)
- **Konteks:** Saat mengedit file kode (.js, .gs, .html) yang merupakan bagian dari proyek Google Apps Script lokal (biasanya terdapat file `.clasp.json`).
- **Tindakan Wajib:** SETELAH Anda selesai melakukan pengeditan menggunakan tool seperti `replace_file_content` atau `write_to_file`, Anda **WAJIB** mengeksekusi perintah `clasp push` di dalam direktori tersebut menggunakan `run_command`.
- **Alasan:** Perubahan pada disk lokal tidak akan berdampak apapun pada sistem *cloud* Google Apps Script milik pengguna sampai kode tersebut di-*push*. Jangan pernah meminta pengguna menguji fitur atau perbaikan bug sebelum Anda memastikan `clasp push` telah sukses dieksekusi.

## Aturan Fail-Safe Komunikasi (Email Quota Fallback)
- **Konteks:** Google Apps Script memiliki limit pengiriman 100 email per hari. Jika limit ini tercapai, `_sendEmailNotification` akan gagal secara *silent* dari sudut pandang pengguna.
- **Tindakan Wajib:** Setiap blok `catch` pada fungsi pengiriman pesan utama (seperti `_sendEmailNotification`) **WAJIB** memiliki mekanisme *fallback* (cadangan) yang meneruskan pemberitahuan *error* tersebut ke Admin melalui Telegram (menggunakan `ADMIN_CHAT_ID`).
- **Alasan:** Pengguna tidak boleh dibiarkan dalam kebingungan (tidak tahu apakah perintah destruktif mereka tereksekusi atau tidak) hanya karena kuota email habis. Sistem harus memiliki jalur komunikasi cadangan untuk memperingatkan pengguna tentang status eksekusi terakhir.

## Aturan Penyimpanan State Multi-Turn (Anti-Amnesia)
- **Konteks:** Saat membangun fitur percakapan multi-tahap (seperti `/wipe` yang meminta opsi A/B/C/D), sistem harus mengingat status pengguna (*state*).
- **Tindakan Wajib:** JANGAN PERNAH menggunakan `CacheService` untuk menyimpan *state* percakapan pengguna yang kritis atau membutuhkan waktu tunggu (*human delay*). Anda **WAJIB** menggunakan `PropertiesService.getScriptProperties()` untuk menyimpan *state*.
- **Alasan:** `CacheService` bersifat *volatile* dan bisa dihapus sepihak oleh server Google kapan saja sebelum kedaluwarsa (eviction). `PropertiesService` adalah penyimpanan persisten yang menjamin *state* tetap aman sampai kita menghapusnya secara manual lewat kode.

## Aturan Resolusi Kredensial Proaktif (Anti-Placeholder)
- **Konteks:** Saat memberikan instruksi, perintah (*command*), atau URL yang membutuhkan API Key, Token, atau kredensial pengguna lainnya (misalnya Telegram Token, GitHub Token).
- **Tindakan Wajib:** JANGAN menggunakan *placeholder* (seperti `<TOKEN_ANDA>`) jika kredensial tersebut bisa ditemukan secara lokal. Agen **WAJIB** terlebih dahulu mencari kredensial tersebut di dalam file lingkungan pengguna (seperti `.env` di *workspace* saat ini atau `.env` di *workspace* utama seperti `Nexus-DualBrain-AI`).
- **Alasan:** Meminimalisir beban kerja pengguna dan mencegah *human error*. Pengguna harus bisa langsung menyalin (*copy-paste*) instruksi tanpa harus mengeditnya lagi secara manual.

## Aturan Registrasi Webhook Anti-Gagal (Zero Human Error)
- **Konteks:** Saat menginstruksikan pengguna untuk mendaftarkan Webhook Telegram (atau platform lain) yang membutuhkan query parameter (misalnya `?token=XYZ`).
- **Tindakan Wajib:** JANGAN PERNAH meminta pengguna untuk merakit URL secara manual di browser (misalnya dengan memberikan instruksi untuk membuka `https://api.telegram.org/bot.../setWebhook?url=...`).
- **Alasan:** Rentan terhadap kesalahan *URL encoding* dan *copy-paste* (parameter menyatu/hilang) yang menyebabkan server menolak akses ("Forbidden") secara diam-diam.
- **Solusi Wajib:** Anda WAJIB membuat fungsi khusus (misalnya `forceRegisterWebhook()`) di dalam kode Apps Script (menggunakan `UrlFetchApp`) dan meminta pengguna mengeksekusi fungsi tersebut langsung dari antarmuka pengguna (seperti Custom Menu di Google Sheets) atau dari editor Google Apps Script.

## Aturan Webhook Timeout & Deduplikasi (Anti-Spam Telegram)
- **Konteks:** Saat menangani *webhook* dari Telegram (`doPost`) yang melibatkan proses asinkron yang lambat seperti pemanggilan AI/LLM.
- **Tindakan Wajib:** JANGAN PERNAH membiarkan *webhook* memproses pesan tanpa sistem deduplikasi. Anda **WAJIB** mengekstrak `update_id` dari *payload* Telegram dan menyimpannya ke dalam `CacheService` (dengan *TTL* 1-2 jam) tepat di awal blok eksekusi `doPost`.
- **Alasan:** Proses pemanggilan LLM seringkali melebihi batas waktu 5 detik milik Telegram. Telegram akan menganggap respons lambat sebagai kegagalan dan secara otomatis melakukan *retry* (mengirim ulang *webhook*). Tanpa deduplikasi, agen AI akan memproses perintah yang sama berkali-kali secara beruntun (spam).
- **Contoh Implementasi Wajib:**
  ```javascript
  const updateId = body.update_id.toString();
  const cache = CacheService.getScriptCache();
  if (cache.get("TG_UPDATE_" + updateId)) return ContentService.createTextOutput("OK");
  cache.put("TG_UPDATE_" + updateId, "1", 3600);
  ```

## Aturan Injeksi Konteks Dinamis (Anti-Halusinasi Sheet)
- **Konteks:** Saat membangun *prompt* untuk LLM/AI yang bertugas menghasilkan perintah modifikasi Google Sheets (seperti `DELETE_ROWS`, `CLEAR_CONTENTS`).
- **Tindakan Wajib:** JANGAN PERNAH menuliskan daftar nama *sheet* secara manual (*hardcoded*) di dalam *prompt*. Anda **WAJIB** mengambil daftar nama *sheet* secara dinamis menggunakan API dan menyuntikkannya ke dalam teks *prompt*.
- **Alasan:** Nama *sheet* yang ditulis manual sangat rentan terhadap kesalahan ketik atau perbedaan jamak/tunggal (misalnya `Transaction` vs `Transactions`). LLM akan mematuhi *prompt* yang salah tersebut, sehingga eksekusi akhir akan selalu gagal (fungsi `getSheetByName` mereturn `null`).
- **Contoh Implementasi Wajib:**
  ```javascript
  // BENAR
  const sheetNames = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName()).join(", ");
  const prompt = \`We have the following sheets: \${sheetNames}. Use exact names!\`;
  ```

## Aturan Mindset Produk Komersial (SaaS)
- **Konteks:** Saat merancang fitur, menulis kode logika, atau memberikan instruksi kepada AI di dalam *Smart Inventory System*.
- **Tindakan Wajib:** Selalu asumsikan bahwa sistem ini adalah **Produk Jual (SaaS) untuk Pelanggan (B2B/B2C)**, bukan skrip pribadi. 
- **Larangan Keras:** DILARANG melakukan *hardcode* terhadap variabel lingkungan yang bisa berubah antar pelanggan (seperti jumlah *tab* Spreadsheet, nama *tab*, struktur kolom, URL *webhook*, atau ID tertentu).
- **Solusi:** Semua parameter yang bergantung pada lingkungan pengguna (pelanggan) WAJIB diekstrak secara dinamis melalui API (contoh: `getSheets()`, `getMaxRows()`) atau disimpan di `PropertiesService` / Database. Sistem harus bisa beradaptasi otomatis saat pelanggan menambahkan *tab* baru atau mengubah struktur bisnis mereka.

## Aturan Mengusulkan Fitur Baru (Anti-Misskomunikasi)
- **Konteks:** Saat memberikan saran, ide, atau *brainstorming* mengenai fitur canggih apa yang bisa ditambahkan ke sistem di masa depan.
- **Tindakan Wajib:** AI WAJIB secara eksplisit menegaskan bahwa fitur yang disebutkan **BELUM DIBUAT** dan sekadar ide.
- **Larangan Keras:** DILARANG menggunakan kalimat ambigu seperti *"Mungkin fitur X, atau fitur Y?"* yang bisa membuat pengguna mengira fitur tersebut sudah aktif.
- **Contoh Implementasi Wajib:**
  "Sebagai ide untuk tahap selanjutnya (jika Anda berminat dan fitur ini belum ada sekarang), kita bisa membangun fitur: X, Y, Z."

## Aturan Lokalisasi Bahasa (SaaS Multilingual)
- **Konteks:** Saat merancang *prompt* sistem untuk Groq/Gemini, atau saat membalas pesan pengguna.
- **Tindakan Wajib:** Sistem **WAJIB** merespons dan menghasilkan teks antarmuka (termasuk menu opsi, *warning*, atau notifikasi) dalam bahasa yang SAMA persis dengan bahasa input pengguna. 
- **Solusi:** Di dalam sistem *prompt* internal, jangan melakukan *hardcode* teks ke bahasa Inggris jika itu akan dibaca oleh pengguna akhir. Tambahkan instruksi eksplisit ke dalam prompt seperti: *"Respond ONLY in the same language the user used."*

## Aturan Arsitektur Google Apps Script Library (Stub.gs)
- **Konteks:** Menambah fungsi UI atau fungsi pemicu (menu) baru ke dalam sistem AI Smart Inventory yang berjalan sebagai Library Google Apps Script (`SmartInventoryCore`).
- **Akar Masalah (Kenapa tidak ditaruh di Library saja?):** Secara arsitektur, Google Sheets **hanya** bisa mengeksekusi fungsi yang berada di dalam *Bound Script* miliknya sendiri (`Stub.gs`). Fungsi yang berada murni di dalam *Library* eksternal tidak akan dikenali oleh antarmuka (UI) Google Sheets. Oleh karena itu, kita membuat kode utama di Library (`SmartInventoryCore.createSecurityAuditTab()`), tetapi kita **tetap membutuhkan** fungsi perantara (*wrapper*) di `Stub.gs` agar UI bisa memanggilnya.
- **Aturan Wajib:** Setiap penambahan fungsi menu baru di kode sumber (Library), Anda WAJIB membuatkan fungsi *wrapper/stub* di sisi klien. 
- **Tindakan Otomatisasi (Full Automation):** JANGAN meminta pengguna untuk *copy-paste* kode secara manual. Mengingat Anda memiliki akses ke seluruh *workspace*, Anda WAJIB menggunakan perintah terminal (seperti PowerShell atau `multi_replace_file_content`) untuk mencari folder *client script* (misalnya `AI_Smart_Inventory_System`), mengedit file `Stub.js` secara langsung, dan mengeksekusi `clasp push`!

## Aturan Google Apps Script: Simple Triggers UI
- **Konteks:** Menggunakan onOpen(e) atau onEdit(e) untuk memanipulasi antarmuka (UI).
- **Akar Masalah:** Google Apps Script secara tegas melarang *simple triggers* untuk menampilkan *Sidebar* atau *Modal Dialog* secara otomatis. Melakukan hal ini akan menyebabkan *script crash* dan membatalkan semua perubahan UI (termasuk pembentukan menu).
- **Solusi Wajib:** JANGAN PERNAH memanggil SpreadsheetApp.getUi().showSidebar() atau showModalDialog() langsung dari dalam onOpen(). Sidebar atau Dialog hanya boleh dibuka dari fungsi yang diikat ke tombol menu yang diklik secara eksplisit oleh pengguna.

## Aturan Lisensi SaaS (One-to-One Binding)
- **Konteks:** Validasi lisensi komersial di Google Sheets.
- **Akar Masalah:** Jika kunci lisensi hanya divalidasi ketersediaannya, pelanggan bisa melakukan pembajakan *sharing* lisensi massal.
- **Solusi Wajib:** Sistem lisensi HARUS menerapkan *Account Binding* atau *Sheet Binding*. Saat diaktivasi, Kunci Lisensi harus mengikat `Session.getEffectiveUser().getEmail()` atau `SpreadsheetApp.getActiveSpreadsheet().getId()`. Penggunaan kunci yang sama di akun/file yang berbeda harus diblokir mutlak.

## Aturan Google Apps Script: Simple Triggers UI (REVISI FATAL)
- **Konteks:** Menggunakan onOpen(e) atau onEdit(e) untuk memanipulasi UI.
- **Akar Masalah:** Pelarangan pemanggilan *Sidebar* atau *Modal Dialog* di *simple triggers* bersistem **Hukuman Mati (Abrupt Termination)**. Blok 	ry-catch TIDAK AKAN BISA menyelamatkan skrip dari kematian ini.
- **Solusi Wajib:** Pemanggilan SpreadsheetApp.getUi().showSidebar() atau showModalDialog() HARUS DIHAPUS SEPENUHNYA dari alur pemanggilan onOpen(). Anda tidak boleh menggunakan 	ry-catch sebagai trik.
`n## Aturan Distribusi SaaS (Library Versioning)`n- **Konteks:** Menghubungkan *Client Script* (Google Sheets) dengan *Core Library*.`n- **Akar Masalah:** Menggunakan `"version": "HEAD"` atau `"developmentMode": true` di dalam `appsscript.json` akan menyebabkan koneksi *Library* terputus total saat klien menekan tombol "Salin", karena Google otomatis mencabut mode developer pada file salinan.`n- **Solusi Wajib:** Saat merilis pembaruan atau menyiapkan file untuk disalin/didistribusikan, Anda WAJIB: 1. Membuat versi permanen baru di Core Library menggunakan perintah clasp version. 2. Memperbarui `appsscript.json` di *Client Script* untuk menunjuk ke nomor versi tersebut (misal: `"version": "38"`).
`n## Aturan Pengujian SaaS (Master vs Copy State)`n- **Konteks:** Menguji fitur Onboarding atau Validasi Lisensi.`n- **Akar Masalah:** *Script Properties* (yang menyimpan status aktivasi lisensi) bersifat permanen di *Master Sheet* namun di-reset menjadi kosong saat file disalin.`n- **Solusi Wajib:** Jangan pernah mengevaluasi keberhasilan antarmuka pengguna (UI) perdana/lisensi dengan melihat tampilan di *Master Sheet*. *Master Sheet* selalu berada dalam status "Teraktivasi". Untuk menguji apa yang dilihat klien baru, pengujian WAJIB dilakukan pada *File Salinan* (Copy), bukan *Master*.
`n## Aturan Distribusi SaaS (Library Permissions)`n- **Konteks:** Menyiapkan Google Sheets yang memiliki ketergantungan pada *Core Library* untuk didistribusikan (klien melakukan "Make a Copy").`n- **Akar Masalah:** Jika proyek Apps Script *Core Library* bersifat *Private*, Google akan memblokir eksekusi skrip di file salinan milik klien secara diam-diam (*silent failure*), sehingga menu kustom atau fitur apapun tidak akan muncul.`n- **Solusi Wajib:** Proyek *Core Library* WAJIB dibagikan hak aksesnya. Anda harus membuka Apps Script dari *Core Library*, menekan tombol "Share" (Bagikan) di kanan atas, dan mengubah izinnya menjadi **"Anyone with the link can view" (Siapa saja yang memiliki tautan dapat melihat)**. Ini aman karena mereka hanya bisa mengeksekusi dan melihat kode, tapi tidak bisa mengubah kode rahasia Anda.
`n## Aturan Arsitektur Google Sheets UI (Stub-First UI)`n- **Konteks:** Menggambar menu kustom (Custom Menu) di file klien yang menggunakan *Core Library*.`n- **Akar Masalah:** Menyerahkan pembentukan menu (terutama UI untuk file yang belum diaktivasi/lisensi) ke dalam `onOpen` milik *Library* sering menyebabkan *silent crash* akibat batas waktu muat (loading timeout) atau isu izin Google.`n- **Solusi Wajib:** Pindahkan logika pembentukan menu dasar (*Locked Menu*) ke dalam `onOpen` milik skrip klien (`Stub.js`). Skrip klien harus mengecek `PropertiesService` secara mandiri, dan jika lisensi kosong, klien sendirilah yang menggambar UI-nya. Pemanggilan `CoreLibrary.onOpen(e)` hanya dilakukan jika status sudah teraktivasi.
`n## Aturan Lingkungan Pengujian Google Apps Script`n- **Konteks:** Menguji eksekusi `onOpen` dan UI di Google Sheets.`n- **Akar Masalah:** Bug Multi-Login Google Chrome menyebabkan kegagalan eksekusi skrip tanpa pesan error. Selain itu, memastikan sinkronisasi antara ID Script lokal (`clasp`) dengan ID Script di file Google Sheets Master pengguna adalah wajib.`n- **Solusi Wajib:** `n  1. Selalu instruksikan pengujian (sebagai klien) menggunakan jendela **Incognito** untuk menghindari bug multi-akun.`n  2. Jika kode dipastikan benar namun hasil tidak muncul, periksa keselarasan `scriptId` di `.clasp.json` dengan URL Script Editor pengguna, dan mintalah pengguna untuk memeriksa *Apps Script Editor* mereka secara manual.
`n## Aturan Penyimpanan Memori Library (Library Scope Bug)`n- **Konteks:** Membuat sistem *Library* untuk Google Sheets (*SaaS Template*).`n- **Akar Masalah:** `PropertiesService.getScriptProperties()` yang dipanggil di dalam file *Library* akan menyimpan dan membaca memori milik proyek *Library* secara global, bukan memori milik *Client Script* yang memanggilnya. Ini menyebabkan data pengguna bentrok, atau hilang saat halaman direfresh oleh klien.`n- **Solusi Wajib:** `n  1. WAJIB menggunakan `PropertiesService.getDocumentProperties()` untuk sistem *Client-Library* yang terikat pada Spreadsheet (Container-bound), agar klien dan library berbagi memori yang sama.`n  2. Karena `DocumentProperties` otomatis ikut tersalin jika pengguna menyalin Spreadsheet, WAJIB menerapkan logika **Deteksi Kloning** di klien: Simpan ID Spreadsheet saat aktivasi, lalu di `onOpen`, bandingkan dengan ID saat ini. Jika berbeda, langsung eksekusi `props.deleteAllProperties()` untuk mencegah pencurian lisensi dan kuota API.
`n## Aturan Injeksi Dependensi PropertiesService (Library Scope Fix)`n- **Konteks:** Membuat sistem *Library* untuk Google Sheets (*SaaS Template*).`n- **Akar Masalah:** Baik `ScriptProperties` maupun `DocumentProperties` yang dipanggil di dalam file *Library* yang bersifat *Container-bound* akan menargetkan memori milik proyek *Library* itu sendiri, BUKAN klien yang memanggilnya.`n- **Solusi Wajib:** `n  1. WAJIB menerapkan arsitektur *Dependency Injection* (Injeksi Dependensi).`n  2. Buat fungsi `setClientProperties(props)` di dalam Library.`n  3. Setiap fungsi utama di `Stub.js` (Client) WAJIB memanggil `Library.setClientProperties(PropertiesService)` pada baris pertama eksekusinya, untuk menanamkan referensi memori Klien ke dalam Library pada setiap siklus eksekusi (menembus isolasi memori V8).
`n## Aturan Anti-Tebak & Verifikasi Kode (Zero-Guessing Code Verification)`n- **Konteks:** Saat menangani bug yang terus berulang atau mengonfirmasi keberhasilan sebuah perombakan arsitektur besar.`n- **Akar Masalah:** AI sering kali berasumsi bahwa modifikasi kode telah berhasil tanpa menguji secara menyeluruh (*blind assumption*).`n- **Solusi Wajib:** `n  1. DILARANG KERAS mengasumsikan keberhasilan.`n  2. Saat pengguna meminta verifikasi, WAJIB menggunakan alat bantu kode (misalnya, skrip PowerShell *Static Code Analysis* dengan regex) untuk memindai seluruh *codebase* demi memastikan tidak ada satu baris pun yang terlewat.`n  3. Berikan "Laporan QC (Quality Control)" yang merinci jumlah file, jumlah fungsi, dan persentase kelulusan metrik yang diuji (contoh: 36/36 fungsi terinjeksi).

## Aturan Ekstraksi JSON (Anti-Halusinasi Chatty Models)
- **Konteks:** Mengurai (parsing) JSON dari model AI yang cenderung memberikan teks penjelasan panjang atau *Chain of Thought* (seperti Gemma atau Llama), terutama di dalam lingkungan Google Apps Script (`AIAgent.js`).
- **Akar Masalah:** Model sering menyertakan tanda kurung siku/kurawal palsu (`[` atau `{`) di dalam teks penjelasannya. Fungsi ekstraktor yang hanya mencari index `[` pertama dan `]` terakhir akan mengambil *string* cacat dan menyebabkan error "Gagal memproses output AI menjadi JSON".
- **Aturan Wajib:** 
  1. *System Prompt* WAJIB secara eksplisit melarang penalaran: "DILARANG KERAS memberikan penjelasan, langkah berpikir, atau teks pengantar apapun. Output HARUS langsung berupa JSON array/object."
  2. DILARANG menggunakan metode `indexOf` ke `lastIndexOf` tunggal untuk mengambil JSON. Fungsi ekstraktor (`_extractJson`) WAJIB menggunakan algoritma perulangan (`while`) yang memindai dan menyusutkan batas (mencoba semua kombinasi kurung) dipadukan dengan `try...catch(JSON.parse)` hingga menemukan blok yang valid, untuk menjamin kekebalan terhadap teks acak.

## Aturan Sinkronisasi AI Prompt & Execution Engine (Formatting Bug Fix)
- **Konteks:** Mengembangkan fitur UI berbasis AI (seperti perintah `/format` atau sejenisnya) di mana *System Prompt* menjabarkan parameter dan perintah (cmd) yang bisa dieksekusi.
- **Akar Masalah:** Divergensi (*mismatch*) antara string perintah yang diajarkan ke LLM (misal: `SET_BACKGROUND`) dengan string yang diharapkan oleh mesin *backend* Google Apps Script (misal: `SET_BACKGROUND_COLOR`). Ini menyebabkan perintah gagal diam-diam.
- **Aturan Wajib:** 
  1. Setiap kali menambah atau memodifikasi fungsionalitas di sistem pengeksekusi (*Execution Engine*, seperti `Agent.js`), agen WAJIB memastikan bahwa *System Prompt* di `AIAgent.js` (atau antarmuka AI terkait) diperbarui secara akurat untuk mencerminkan parameter terbaru.
  2. JANGAN PERNAH berasumsi bahwa format yang dihasilkan AI bisa dieksekusi tanpa memeriksa blok `if...else` / `switch` pada fungsi eksekusi aslinya. Pastikan nama *keys* (`cmd`, `range`, `color`) dan fungsinya sinkron 100%.
