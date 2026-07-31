# LLM Gateway for Excel

Ultra-lightweight, high-performance, OpenAI-compatible AI sidebar add-in for Microsoft Excel.

---

## 1. Mục đích dự án (Project Purpose)
`LLM Gateway for Excel` là một add-in giao diện thanh bên (sidebar taskpane) dành cho Microsoft Excel, cho phép người dùng tích hợp trực tiếp các mô hình trí tuệ nhân tạo (LLM) hỗ trợ chuẩn OpenAI Chat Completions API (`/v1/chat/completions`) vào bảng tính Excel.

Dự án giải quyết bài toán:
* Cung cấp trợ lý AI đọc/ghi, viết công thức, định dạng và phân tích dữ liệu bảng tính siêu nhẹ và phản hồi nhanh.
* Tương thích linh hoạt với mọi nhà cung cấp LLM: OpenAI, DeepSeek, Local AI (Ollama, LM Studio, vLLM), OpenRouter, LiteLLM, Groq, Qwen, v.v.
* **Hoàn toàn độc lập:** Tất cả logic AI & Excel chạy 100% trên trình duyệt trong sidebar Excel. Không cần bất kỳ background server hay script Python nào!

---

## 2. Công nghệ áp dụng (Tech Stack)
* **Core Runtime:** Vanilla TypeScript + HTML5 / CSS3.
* **Bundler & Dev Server:** Vite 6.x.
* **Excel Integration API:** Microsoft `Office.js` (giao tiếp trực tiếp với Microsoft Excel trên Desktop, Web và Mac).
* **LLM API Protocol:** OpenAI-compatible REST API Client (`/v1/chat/completions`) với cơ chế Function Calling (`tools`).

---

## 3. Yêu cầu hệ thống (Requirements)

### Phần mềm (Software Requirements)
* **Node.js:** `v18.0.0` trở lên (Khuyến nghị `v20.x` hoặc `v22.x`).
* **Trình duyệt / Excel Host:** Microsoft Excel Desktop (Windows / macOS) hoặc Excel on the Web.
* **Mô hình LLM:** Bất kỳ OpenAI-compatible endpoint nào (Cloud API Key hoặc Local AI server như Ollama/LM Studio).

### Phần cứng tối thiểu (Hardware Requirements)
* CPU: 2 Cores.
* RAM: 4 GB.

---

## 4. Hướng Dẫn Đưa Lên GitHub Pages (Dùng Online Mọi Lúc Không Cần Chạy Server)

### Bước 1: Tạo Repository mới trên GitHub
Tạo 1 repo mới trên GitHub tên `llmgateway-for-excel`.

### Bước 2: Push code lên GitHub & Deploy
Mở terminal tại thư mục dự án và chạy các lệnh sau:

```bash
git init
git add .
git commit -m "Initial commit for LLM Gateway Excel Add-in"
git branch -M main
git remote add origin https://github.com/<YOUR_GITHUB_USERNAME>/llmgateway-for-excel.git
git push -u origin main
```

Sau đó chạy lệnh deploy tự động đẩy bản build lên GitHub Pages:
```bash
npm run deploy
```

Trang web Add-in của bạn sẽ tự động chạy tại URL:
`https://<YOUR_GITHUB_USERNAME>.github.io/llmgateway-for-excel/`

### Bước 3: Cập nhật `manifest.xml` & Nạp vào Excel
Thay thế `YOUR_HOSTED_URL` trong `manifest.prod.xml` bằng link GitHub Pages ở trên (ví dụ `https://username.github.io/llmgateway-for-excel`), lưu thành `manifest.xml` và nạp vào Excel.

Từ giờ về sau, mở Excel lên là thanh Sidebar tự động mở Add-in mà không cần bất kỳ server hay lệnh nào trên máy bạn!

---

## 5. Cấu trúc thư mục (Directory Layout)

```
llmgateway-for-excel/
├── dist/                   # Production static bundle
├── manifest.xml            # Dev/Local manifest (https://localhost:3141)
├── manifest.prod.xml       # Production manifest template for GitHub Pages
├── package.json            # Project dependencies & scripts (với npm run deploy)
├── tsconfig.json           # TypeScript configuration
├── vite.config.ts          # Vite build config với relative base path
├── taskpane.html           # Taskpane Sidebar HTML layout
├── README.md               # Project documentation
└── src/
    ├── main.ts             # App controller & Office.onReady entry
    ├── types.ts            # Type definitions & OpenAI tool interfaces
    ├── api/
    │   └── llmClient.ts    # OpenAI-compatible API client & tool loop
    ├── excel/
    │   └── tools.ts        # Office.js tools implementation
    ├── storage/
    │   └── settingsStore.ts # LocalStorage settings manager
    └── ui/
        └── style.css       # Modern dark/light glassmorphism theme
```
