const express = require("express");
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const app = express();
const PORT = process.env.PORT || 3000;

// ==============================
// 基本設定
// ==============================

app.use(express.json({ limit: "2mb" }));
app.use(express.urlencoded({ extended: true }));

// サイト本体
app.use(express.static(__dirname));

// Excel商品マスタ
const EXCEL_FILE = path.join(__dirname, "価格表.xlsm");
const PRODUCT_SHEET = "Sheet2";

// 見積データ保存先
const ESTIMATE_FILE = path.join(__dirname, "estimates.json");


// ==============================
// Excelから商品データを読み込む
// Sheet2だけを使用
// ==============================

function loadProducts() {
  if (!fs.existsSync(EXCEL_FILE)) {
    throw new Error(
      "価格表.xlsm が見つかりません。GitHubのリポジトリ直下に価格表.xlsmを置いてください。"
    );
  }

  const workbook = XLSX.readFile(EXCEL_FILE, {
    cellDates: false
  });

  if (!workbook.SheetNames.includes(PRODUCT_SHEET)) {
    throw new Error(
      `Excelに「${PRODUCT_SHEET}」シートがありません。`
    );
  }

  const sheet = workbook.Sheets[PRODUCT_SHEET];

  // 1行目を見出しとしてオブジェクト化
  const rows = XLSX.utils.sheet_to_json(sheet, {
    defval: "",
    raw: true
  });

  const products = rows
    .map((row) => ({
      code: row["品番"],
      size: row["サイズ"],
      a: row["A表"],
      price: row["価格"],
      brand: row["ブランド"],
      pattern: row["パターン"]
    }))
    .filter((p) => {
      return Object.values(p).some(
        (value) => String(value ?? "").trim() !== ""
      );
    });

  return products;
}


// ==============================
// 商品一覧API
// ==============================

app.get("/api/products", (req, res) => {
  try {
    const products = loadProducts();

    res.json({
      success: true,
      count: products.length,
      products
    });
  } catch (error) {
    console.error("商品データ読み込みエラー:", error);

    res.status(500).json({
      success: false,
      message: error.message
    });
  }
});


// ==============================
// 見積依頼API
// ==============================

app.post("/api/estimate", (req, res) => {
  try {
    const body = req.body || {};

    const estimate = {
      id: Date.now(),
      company: body.company || "",
      phone: body.phone || "",
      email: body.email || "",
      note: body.note || "",
      items: Array.isArray(body.items) ? body.items : [],
      createdAt: new Date().toISOString()
    };

    let estimates = [];

    if (fs.existsSync(ESTIMATE_FILE)) {
      try {
        estimates = JSON.parse(
          fs.readFileSync(ESTIMATE_FILE, "utf8")
        );

        if (!Array.isArray(estimates)) {
          estimates = [];
        }
      } catch (error) {
        estimates = [];
      }
    }

    estimates.push(estimate);

    fs.writeFileSync(
      ESTIMATE_FILE,
      JSON.stringify(estimates, null, 2),
      "utf8"
    );

    console.log("見積依頼を受け付けました");

    res.json({
      success: true,
      message: "見積依頼を受け付けました"
    });
  } catch (error) {
    console.error("見積保存エラー:", error);

    res.status(500).json({
      success: false,
      message: "見積依頼の保存に失敗しました"
    });
  }
});


// ==============================
// サーバー起動
// ==============================

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
  console.log(`Excel: ${EXCEL_FILE}`);
  console.log(`商品シート: ${PRODUCT_SHEET}`);
});
