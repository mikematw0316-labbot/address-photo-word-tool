# 地址照片整理與 Word 產生器（雲端版）

這是可部署的網站服務。部署完成後，使用者只需網址，不必安裝程式，也不綁定 Mac 或 Codex。

## 使用方式

1. 使用者以瀏覽器開啟網站。
2. 拖入照片或 ZIP，按「開始辨識與分類」。
3. 確認系統標示的待確認地址。
4. 產生並下載成果 ZIP。

## 部署

### Render

1. 將此資料夾推送到 GitHub repository。
2. 在 Render 選擇 **New > Blueprint**。
3. 選擇該 repository；Render 會讀取 `render.yaml` 並建立網站。

### 任何 Docker 主機

```bash
docker build -t address-photo-word-tool .
docker run --rm -p 8765:8765 address-photo-word-tool
```

## 處理原則

- 原始照片只會被複製處理，不會修改或覆蓋。
- 使用繁體中文 Tesseract OCR，不依賴使用者的作業系統。
- 暫存檔預設在 24 小時後自動刪除。
- 地址不完整、無法辨識或複合地址不確定時會進入待確認區。
- 每個第二層地址產生一份獨立 Word。
- Word 直接複製提供的公會範本紅框，不重新繪製表格。

## 支援格式

JPG、JPEG、PNG、HEIC、HEIF、TIFF、BMP、WebP，以及包含上述照片的 ZIP。
