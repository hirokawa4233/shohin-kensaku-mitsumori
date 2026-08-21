const express = require("express");
const fs = require("fs");
const path = require("path");
const https = require("https");
const XLSX = require("xlsx");
const PDFDocument = require("pdfkit");

const app = express();
const PORT = process.env.PORT || 3000;

// ==============================
// 基本設定
// ==============================

app.use(express.json({ limit: "2mb" }));
app.use(express.urlencoded({ extended: true }));
app.use(express.static(__dirname));


// ==============================
// Excel商品マスタ
// ==============================

const EXCEL_FILE = path.join(__dirname, "価格表.xlsm");
const PRODUCT_SHEET = "Sheet2";


// ==============================
// 見積データ保存先
// ==============================

const ESTIMATE_FILE =
  path.join(__dirname, "estimates.json");


// ==============================
// PDF用日本語フォント
// ==============================

const FONT_FILE =
  path.join(__dirname, "NotoSansJP-Regular.otf");

const FONT_URL =
  "https://github.com/notofonts/noto-cjk/raw/main/Sans/SubsetOTF/JP/NotoSansJP-Regular.otf";


// ==============================
// 日本語フォントを取得
// ==============================

function downloadFile(url, destination) {

  return new Promise((resolve, reject) => {

    const file =
      fs.createWriteStream(destination);

    https.get(url, (response) => {

      // リダイレクト
      if (
        response.statusCode >= 300 &&
        response.statusCode < 400 &&
        response.headers.location
      ) {

        file.close();

        try {
          fs.unlinkSync(destination);
        } catch (_) {}

        return downloadFile(
          response.headers.location,
          destination
        )
        .then(resolve)
        .catch(reject);
      }

      if (
        response.statusCode < 200 ||
        response.statusCode >= 400
      ) {

        file.close();

        try {
          fs.unlinkSync(destination);
        } catch (_) {}

        return reject(
          new Error(
            `フォント取得に失敗しました。HTTP ${response.statusCode}`
          )
        );
      }

      response.pipe(file);

      file.on("finish", () => {

        file.close(resolve);

      });

      file.on("error", (error) => {

        file.close();

        try {
          fs.unlinkSync(destination);
        } catch (_) {}

        reject(error);

      });

    }).on("error", (error) => {

      file.close();

      try {
        fs.unlinkSync(destination);
      } catch (_) {}

      reject(error);

    });

  });

}


// ==============================
// 日本語フォントを準備
// ==============================

async function ensureJapaneseFont() {

  if (fs.existsSync(FONT_FILE)) {
    return;
  }

  console.log(
    "日本語フォントを取得しています..."
  );

  await downloadFile(
    FONT_URL,
    FONT_FILE
  );

  console.log(
    "日本語フォントの取得が完了しました。"
  );

}


// ==============================
// Excelから商品データを読み込む
// ==============================

function loadProducts() {

  if (!fs.existsSync(EXCEL_FILE)) {

    throw new Error(
      "価格表.xlsm が見つかりません。"
    );

  }

  const workbook =
    XLSX.readFile(EXCEL_FILE, {
      cellDates: false
    });

  if (
    !workbook.SheetNames.includes(
      PRODUCT_SHEET
    )
  ) {

    throw new Error(
      `Excelに「${PRODUCT_SHEET}」シートがありません。`
    );

  }

  const sheet =
    workbook.Sheets[PRODUCT_SHEET];

  const rows =
    XLSX.utils.sheet_to_json(sheet, {
      defval: "",
      raw: true
    });

  const products =
    rows
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

    const products =
      loadProducts();

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

    const data =
      fs.readFileSync(
        ESTIMATE_FILE,
        "utf8"
      );

    const estimates =
      JSON.parse(data);

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

    const body =
      req.body || {};

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
// 見積依頼一覧API
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
// 納期連絡保存API
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
// 御見積書PDF作成API
// ==============================

app.get(
  "/api/estimates/:id/pdf",
  async (req, res) => {

    try {

      const id =
        Number(req.params.id);

      const estimates =
        loadEstimates();

      const estimate =
        estimates.find(
          (item) =>
            Number(item.id) === id
        );

      if (!estimate) {

        return res.status(404).send(
          "見積データが見つかりません。"
        );

      }

      if (
        !estimate.delivery ||
        !String(
          estimate.delivery
        ).trim()
      ) {

        return res.status(400).send(
          "納期を入力してください。"
        );

      }

      // 日本語フォントを準備
      await ensureJapaneseFont();

      // ==============================
      // PDF作成
      // ==============================

      const doc =
        new PDFDocument({
          size: "A4",
          margin: 50
        });

      res.setHeader(
        "Content-Type",
        "application/pdf"
      );

      res.setHeader(
        "Content-Disposition",
        `attachment; filename="御見積書_${estimate.id}.pdf"`
      );

      doc.pipe(res);

      // 日本語フォント
      doc.font(FONT_FILE);

      // ==============================
      // タイトル
      // ==============================

      doc
        .fontSize(28)
        .text(
          "御 見 積 書",
          {
            align: "center"
          }
        );

      doc.moveDown(2);


      // ==============================
      // 見積番号・受付日時
      // ==============================

      const createdDate =
        estimate.createdAt
          ? new Date(
              estimate.createdAt
            ).toLocaleString(
              "ja-JP"
            )
          : "";

      doc
        .fontSize(11)
        .text(
          `見積番号：${estimate.id}`,
          {
            align: "right"
          }
        )
        .text(
          `受付日時：${createdDate}`,
          {
            align: "right"
          }
        );

      doc.moveDown(2);


      // ==============================
      // お客様情報
      // ==============================

      doc
        .fontSize(16)
        .text(
          "お客様情報",
          {
            underline: true
          }
        );

      doc.moveDown(0.7);

      // ★ここで「様」を付ける
      const customerName =
        String(
          estimate.company || ""
        ).trim();

      const customerDisplay =
        customerName
          ? `${customerName} 様`
          : "";

      doc
        .fontSize(11)
        .text(
          `会社名・氏名：${customerDisplay}`
        )
        .text(
          `電話番号：${estimate.phone || ""}`
        )
        .text(
          `メールアドレス：${estimate.email || ""}`
        );

      doc.moveDown(2);


      // ==============================
      // お見積内容
      // ==============================

      doc
        .fontSize(16)
        .text(
          "お見積内容",
          {
            underline: true
          }
        );

      doc.moveDown(0.8);


      // 表の見出し

      const tableX = 50;

      const colCode = 50;
      const colSize = 150;
      const colBrand = 280;
      const colQty = 410;
      const colAmount = 490;

      doc
        .fontSize(9)
        .text(
          "品番",
          colCode
        )
        .text(
          "サイズ",
          colSize
        )
        .text(
          "ブランド・パターン",
          colBrand
        )
        .text(
          "数量",
          colQty
        )
        .text(
          "金額",
          colAmount
        );

      doc.moveTo(
        tableX,
        doc.y + 5
      )
      .lineTo(
        545,
        doc.y + 5
      )
      .stroke();

      doc.moveDown(0.5);


      // ==============================
      // 商品
      // ==============================

      const items =
        Array.isArray(
          estimate.items
        )
          ? estimate.items
          : [];

      let total = 0;

      items.forEach((item) => {

        const price =
          Number(item.price) || 0;

        const qty =
          Number(item.qty) || 0;

        const subtotal =
          price * qty;

        total += subtotal;

        const brand =
          String(
            item.brand || ""
          );

        const pattern =
          String(
            item.pattern || ""
          );

        const brandPattern =
          `${brand} ${pattern}`.trim();

        doc
          .fontSize(9)
          .text(
            String(
              item.code || ""
            ),
            colCode
          )
          .text(
            String(
              item.size || ""
            ),
            colSize
          )
          .text(
            brandPattern,
            colBrand
          )
          .text(
            `${qty}個`,
            colQty
          )
          .text(
            `${subtotal.toLocaleString("ja-JP")}円`,
            colAmount
          );

        doc.moveDown(0.5);

      });


      // 下線

      doc.moveTo(
        tableX,
        doc.y + 5
      )
      .lineTo(
        545,
        doc.y + 5
      )
      .stroke();

      doc.moveDown(1);


      // ==============================
      // 合計金額
      // ==============================

      doc
        .fontSize(18)
        .text(
          `合計金額：${total.toLocaleString("ja-JP")}円`,
          {
            align: "right"
          }
        );

      doc.moveDown(2);


      // ==============================
      // 納期
      // ==============================

      doc
        .fontSize(16)
        .text(
          "納期",
          {
            underline: true
          }
        );

      doc.moveDown(0.7);

      doc
        .fontSize(11)
        .text(
          estimate.delivery
        );

      doc.moveDown(2);


      // ==============================
      // 備考
      // ==============================

      doc
        .fontSize(16)
        .text(
          "備考",
          {
            underline: true
          }
        );

      doc.moveDown(0.7);

      doc
        .fontSize(11)
        .text(
          estimate.note || "なし"
        );


      // ==============================
      // PDF終了
      // ==============================

      doc.end();

    } catch (error) {

      console.error(
        "PDF作成エラー:",
        error
      );

      if (!res.headersSent) {

        res.status(500).send(
          "御見積書PDFの作成に失敗しました。"
        );

      }

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
