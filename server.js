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

// ==============================
// Excel商品マスタ
// ==============================

const EXCEL_FILE = path.join(__dirname, "価格表.xlsm");
const PRODUCT_SHEET = "Sheet2";

// ==============================
// 見積データ保存先
// ==============================

const ESTIMATE_FILE = path.join(__dirname, "estimates.json");


// ==============================
// Excelから商品データを読み込む
// Sheet2だけを使用
// ==============================

function loadProducts() {

  if (!fs.existsSync(EXCEL_FILE)) {
    throw new Error(
      "価格表.xlsm が見つかりません。"
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
        (value) =>
          String(value ?? "").trim() !== ""
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

    console.error(
      "商品データ読み込みエラー:",
      error
    );

    res.status(500).json({
      success: false,
      message: error.message
    });

  }

});


// ==============================
// 見積データ読み込み
// ==============================

function loadEstimates() {

  if (!fs.existsSync(ESTIMATE_FILE)) {
    return [];
  }

  try {

    const data = fs.readFileSync(
      ESTIMATE_FILE,
      "utf8"
    );

    const estimates = JSON.parse(data);

    if (!Array.isArray(estimates)) {
      return [];
    }

    return estimates;

  } catch (error) {

    console.error(
      "見積データ読み込みエラー:",
      error
    );

    return [];
  }
}


// ==============================
// 見積データ保存
// ==============================

function saveEstimates(estimates) {

  fs.writeFileSync(
    ESTIMATE_FILE,
    JSON.stringify(
      estimates,
      null,
      2
    ),
    "utf8"
  );

}


// ==============================
// 見積依頼受付API
// ==============================

app.post("/api/estimate", (req, res) => {

  try {

    const body = req.body || {};

    const estimate = {

      id: Date.now(),

      company:
        body.company || "",

      phone:
        body.phone || "",

      email:
        body.email || "",

      note:
        body.note || "",

      items:
        Array.isArray(body.items)
          ? body.items
          : [],

      // 納期連絡
      delivery:
        body.delivery || "",

      createdAt:
        new Date().toISOString()

    };

    const estimates =
      loadEstimates();

    estimates.push(estimate);

    saveEstimates(estimates);

    console.log(
      "見積依頼を受け付けました:",
      estimate.id
    );

    res.json({

      success: true,

      message:
        "見積依頼を受け付けました",

      id:
        estimate.id

    });

  } catch (error) {

    console.error(
      "見積保存エラー:",
      error
    );

    res.status(500).json({

      success: false,

      message:
        "見積依頼の保存に失敗しました"

    });

  }

});


// ==============================
// 管理画面用
// 見積依頼一覧取得API
// ==============================

app.get("/api/estimates", (req, res) => {

  try {

    const estimates =
      loadEstimates();

    res.json({

      success: true,

      count:
        estimates.length,

      estimates

    });

  } catch (error) {

    console.error(
      "見積一覧取得エラー:",
      error
    );

    res.status(500).json({

      success: false,

      message:
        "見積依頼の取得に失敗しました"

    });

  }

});


// ==============================
// 納期連絡を保存するAPI
// ==============================

app.patch(
  "/api/estimates/:id/delivery",
  (req, res) => {

    try {

      const id =
        Number(req.params.id);

      const delivery =
        String(
          req.body?.delivery || ""
        ).trim();

      const estimates =
        loadEstimates();

      const estimate =
        estimates.find(
          (item) =>
            Number(item.id) === id
        );

      if (!estimate) {

        return res.status(404).json({

          success: false,

          message:
            "指定された見積依頼が見つかりません"

        });

      }

      estimate.delivery =
        delivery;

      estimate.deliveryUpdatedAt =
        new Date().toISOString();

      saveEstimates(estimates);

      console.log(
        "納期連絡を更新しました:",
        id
      );

      res.json({

        success: true,

        message:
          "納期連絡を保存しました",

        estimate

      });

    } catch (error) {

      console.error(
        "納期保存エラー:",
        error
      );

      res.status(500).json({

        success: false,

        message:
          "納期連絡の保存に失敗しました"

      });

    }

  }
);


// ==============================
// サーバー起動
// ==============================

app.listen(PORT, () => {

  console.log(
    `Server running on port ${PORT}`
  );

  console.log(
    `Excel: ${EXCEL_FILE}`
  );

  console.log(
    `商品シート: ${PRODUCT_SHEET}`
  );

});
