const express = require("express");
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const PDFDocument = require("pdfkit");
const { Pool } = require("pg");
const app = express();
const PORT = process.env.PORT || 3000;
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: {
    rejectUnauthorized: false
  }
});
async function initDatabase() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS estimates (
      id BIGINT PRIMARY KEY,
      company TEXT,
      phone TEXT,
      email TEXT,
      note TEXT,
      delivery TEXT,
      items JSONB,
      created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
  `);

  console.log("PostgreSQL database ready");
}

initDatabase().catch(err => {
  console.error("Database initialization error:", err);
});
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

const ESTIMATE_FILE = path.join(
  __dirname,
  "estimates.json"
);


// ==============================
// PDF保存先
// ==============================

const PDF_DIR = path.join(
  __dirname,
  "pdf"
);

if (!fs.existsSync(PDF_DIR)) {
  fs.mkdirSync(PDF_DIR, {
    recursive: true
  });
}


// ==============================
// 日本語フォント
// ==============================

function findJapaneseFont() {

  const fontCandidates = [

     path.join(
    __dirname,
    "NotoSansCJKjp-Regular.otf"
  ),

  "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",

  "/usr/share/fonts/opentype/noto/NotoSansCJKjp-Regular.otf",

  "/usr/share/fonts/truetype/noto/NotoSansJP-Regular.ttf",

  "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf",

  "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf",

  "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf"

];

  for (const font of fontCandidates) {

    if (fs.existsSync(font)) {

      console.log(
        "Japanese font:",
        font
      );

      return font;

    }

  }

  console.log(
    "Japanese font: NOT FOUND"
  );

  return null;
}

const JAPANESE_FONT =
  findJapaneseFont();


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
    XLSX.readFile(
      EXCEL_FILE,
      {
        cellDates: false
      }
    );

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
    workbook.Sheets[
      PRODUCT_SHEET
    ];

  const rows =
    XLSX.utils.sheet_to_json(
      sheet,
      {
        defval: "",
        raw: true
      }
    );

  const products =
    rows
      .map((row) => {

        return {

          code:
            row["品番"],

          size:
            row["サイズ"],

          a:
            row["A表"],

          price:
            row["価格"],

          brand:
            row["ブランド"],

          pattern:
            row["パターン"]

        };

      })
      .filter((p) => {

        return Object.values(p)
          .some(
            (value) =>
              String(
                value ?? ""
              ).trim() !== ""
          );

      });

  return products;

}


// ==============================
// 商品一覧API
// ==============================

app.get(
  "/api/products",
  (req, res) => {

    try {

      const products =
        loadProducts();

      res.json({

        success: true,

        count:
          products.length,

        products

      });

    } catch (error) {

      console.error(
        "商品データ読み込みエラー:",
        error
      );

      res.status(500).json({

        success: false,

        message:
          error.message

      });

    }

  }
);

// ==============================
// 商品データ保存API
// ==============================

app.post(
  "/api/products",
  (req, res) => {

    try {

      const products =
        Array.isArray(req.body.products)
          ? req.body.products
          : [];

      const workbook =
        XLSX.readFile(
          EXCEL_FILE,
          {
            cellDates: false,
            bookVBA: true
          }
        );

      if (
        !workbook.SheetNames.includes(
          PRODUCT_SHEET
        )
      ) {

        throw new Error(
          `Excelに「${PRODUCT_SHEET}」シートがありません。`
        );

      }

      const rows =
        products.map((product) => ({

          "品番":
            product.code ?? "",

          "サイズ":
            product.size ?? "",

          "A表":
            product.a ?? "",

          "価格":
            product.price ?? "",

          "ブランド":
            product.brand ?? "",

          "パターン":
            product.pattern ?? ""

        }));

      const newSheet =
        XLSX.utils.json_to_sheet(
          rows
        );

      workbook.Sheets[
        PRODUCT_SHEET
      ] =
        newSheet;

      XLSX.writeFile(
        workbook,
        EXCEL_FILE,
        {
          bookType: "xlsm"
        }
      );

      console.log(
        "商品データをExcelに保存しました"
      );

      res.json({

        success: true,

        message:
          "商品データを保存しました"

      });

    } catch (error) {

      console.error(
        "商品データ保存エラー:",
        error
      );

      res.status(500).json({

        success: false,

        message:
          error.message ||
          "商品データの保存に失敗しました"

      });

    }

  }
);
// ==============================
// 見積データ読み込み
// ==============================

function loadEstimates() {

  if (
    !fs.existsSync(
      ESTIMATE_FILE
    )
  ) {

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

    if (
      !Array.isArray(
        estimates
      )
    ) {

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

function saveEstimates(
  estimates
) {

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

app.post(
  "/api/estimate",
  async (req, res) => {

    try {

      const body =
        req.body || {};

      const estimate = {

        id:
          Date.now(),

        company:
          body.company || "",

        phone:
          body.phone || "",

        email:
          body.email || "",

        note:
          body.note || "",

        items:
          Array.isArray(
            body.items
          )
            ? body.items
            : [],

        delivery:
          body.delivery || "",

        createdAt:
          new Date().toISOString()

      };

      const estimates =
        loadEstimates();

      estimates.push(
        estimate
      );

      saveEstimates(
        estimates
      );


      // ==============================
      // 管理者へLINE通知
      // ==============================

      const adminUserId =
        process.env.LINE_ADMIN_USER_ID;

      const channelAccessToken =
        process.env.LINE_CHANNEL_ACCESS_TOKEN;

      if (
        adminUserId &&
        channelAccessToken
      ) {

        try {

          const lineResponse =
            await fetch(
              "https://api.line.me/v2/bot/message/push",
              {
                method: "POST",

                headers: {
                  "Content-Type":
                    "application/json",

                  "Authorization":
                    `Bearer ${channelAccessToken}`
                },

                body: JSON.stringify({

                  to: adminUserId,

                  messages: [
                    {
                      type: "text",

                      text:
                        `新しい見積依頼が届きました。\n\n` +
                        `見積番号：${estimate.id}\n` +
                        `会社名・氏名：${estimate.company}\n` +
                        `電話番号：${estimate.phone}\n` +
                        `メールアドレス：${estimate.email}`
                    }
                  ]

                })

              }
            );


          if (!lineResponse.ok) {

            const lineError =
              await lineResponse.text();

            console.error(
              "LINE通知エラー:",
              lineResponse.status,
              lineError
            );

          } else {

            console.log(
              "管理者へLINE通知を送信しました。"
            );

          }

        } catch (lineError) {

          console.error(
            "LINE通知送信エラー:",
            lineError
          );

        }

      }


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

  }
);


// ==============================
// 管理画面用
// 見積依頼一覧取得API
// ==============================

app.get(
  "/api/estimates",
  (req, res) => {

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

  }
);
// ==============================
// 見積依頼削除API
// ==============================

app.delete(
  "/api/estimates/:id",
  (req, res) => {

    try {

      const id =
        Number(
          req.params.id
        );

      const estimates =
        loadEstimates();

      const index =
        estimates.findIndex(
          (item) =>
            Number(item.id) === id
        );

      if (index === -1) {

        return res.status(404).json({

          success: false,

          message:
            "指定された見積依頼が見つかりません"

        });

      }

      const deletedEstimate =
        estimates[index];

      estimates.splice(
        index,
        1
      );

      saveEstimates(
        estimates
      );

      // ==============================
      // 作成済みPDFも削除
      // ==============================

      const pdfFileName =
        `御見積書_${deletedEstimate.id}.pdf`;

      const pdfPath =
        path.join(
          PDF_DIR,
          pdfFileName
        );

      if (
        fs.existsSync(pdfPath)
      ) {

        fs.unlinkSync(
          pdfPath
        );

        console.log(
          "見積PDFを削除しました:",
          pdfFileName
        );

      }

      console.log(
        "見積依頼を削除しました:",
        id
      );

      res.json({

        success: true,

        message:
          "見積依頼を削除しました",

        id

      });

    } catch (error) {

      console.error(
        "見積削除エラー:",
        error
      );

      res.status(500).json({

        success: false,

        message:
          "見積依頼の削除に失敗しました"

      });

    }

  }
);

// ==============================
// 納期連絡を保存するAPI
// ==============================

app.patch(
  "/api/estimates/:id/delivery",
  (req, res) => {

    try {

      const id =
        Number(
          req.params.id
        );

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

      saveEstimates(
        estimates
      );

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
// 金額フォーマット
// ==============================

function formatNumber(value) {

  return Number(
    value || 0
  ).toLocaleString(
    "ja-JP"
  );

}


// ==============================
// 御見積書PDF作成API
// ==============================

app.post(
  "/api/estimates/:id/pdf",
  (req, res) => {

    try {

      const id =
        Number(
          req.params.id
        );

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


      // ==========================
      // 納期
      // ==========================

      const delivery =
        String(
          req.body?.delivery ??
          estimate.delivery ??
          ""
        ).trim();

      if (!delivery) {

        return res.status(400).json({

          success: false,

          message:
            "納期連絡を入力してください"

        });

      }


      estimate.delivery =
        delivery;

      estimate.deliveryUpdatedAt =
        new Date().toISOString();

      saveEstimates(
        estimates
      );


      // ==========================
      // PDFファイル
      // ==========================

      const pdfFileName =
        `御見積書_${estimate.id}.pdf`;

      const pdfPath =
        path.join(
          PDF_DIR,
          pdfFileName
        );


      // ==========================
      // PDF作成
      // ==========================

      const doc =
        new PDFDocument({

          size: "A4",

          margin: 55

        });


      const stream =
        fs.createWriteStream(
          pdfPath
        );


      doc.pipe(
        stream
      );


      // ==========================
      // 日本語フォント
      // ==========================

      if (JAPANESE_FONT) {

        doc.font(
          JAPANESE_FONT
        );

      }


      // ==========================
      // タイトル
      // ==========================

      doc
        .fontSize(28)
        .text(
          "御 見 積 書",
          {
            align: "center"
          }
        );

      doc.moveDown(2);


      // ==========================
      // 見積情報
      // ==========================

      doc
        .fontSize(11)
        .text(
          `見積番号：${estimate.id}`,
          {
            align: "right"
          }
        );


      const createdDate =
        estimate.createdAt
          ? new Date(
              estimate.createdAt
            ).toLocaleString(
              "ja-JP"
            )
          : "";


      doc.text(
        `受付日時：${createdDate}`,
        {
          align: "right"
        }
      );


      doc.moveDown(2);


      // ==========================
      // お客様情報
      // ==========================

      doc
        .fontSize(16)
        .text(
          "お客様情報",
          {
            underline: true
          }
        );

      doc.moveDown(0.5);


      // ★ ここだけ今回変更
      // お客様のお名前に「様」を付ける

      const customerName =
        String(
          estimate.company || ""
        ).trim();


      doc
        .fontSize(12)
        .text(
          `会社名・氏名：${customerName} 様`
        );


      doc.text(
        `電話番号：${estimate.phone || ""}`
      );


      doc.text(
        `メールアドレス：${estimate.email || ""}`
      );


      doc.moveDown(2);


      // ==========================
      // お見積内容
      // ==========================

      doc
        .fontSize(16)
        .text(
          "お見積内容",
          {
            underline: true
          }
        );

      doc.moveDown(0.8);


      // ==========================
      // 表ヘッダー
      // ==========================

      const headerY =
        doc.y;


      doc
        .fontSize(10)
        .text(
          "品番",
          55,
          headerY,
          {
            width: 100
          }
        );


      doc.text(
        "サイズ",
        155,
        headerY,
        {
          width: 100
        }
      );


      doc.text(
        "ブランド・パターン",
        255,
        headerY,
        {
          width: 170
        }
      );


      doc.text(
        "数量",
        425,
        headerY,
        {
          width: 45
        }
      );


      doc.text(
        "金額",
        470,
        headerY,
        {
          width: 80,
          align: "right"
        }
      );


      doc.y =
        headerY + 24;


      doc
        .moveTo(
          55,
          doc.y
        )
        .lineTo(
          535,
          doc.y
        )
        .stroke();


      doc.y += 10;


      // ==========================
      // 商品
      // ==========================

      let total = 0;

      const items =
        Array.isArray(
          estimate.items
        )
          ? estimate.items
          : [];


      for (
        const item of items
      ) {

        const price =
          Number(
            item.price
          ) || 0;


        const qty =
          Number(
            item.qty
          ) || 0;


        const subtotal =
          price * qty;


        total +=
          subtotal;


        const startY =
          doc.y;


        doc
          .fontSize(10)
          .text(
            String(
              item.code || ""
            ),
            55,
            startY,
            {
              width: 100
            }
          );


        doc.text(
          String(
            item.size || ""
          ),
          155,
          startY,
          {
            width: 100
          }
        );


        doc.text(
          `${String(
            item.brand || ""
          )} ${String(
            item.pattern || ""
          )}`,
          255,
          startY,
          {
            width: 170
          }
        );


        doc.text(
          `${qty}個`,
          425,
          startY,
          {
            width: 45
          }
        );


        doc.text(
          `${formatNumber(
            subtotal
          )}円`,
          470,
          startY,
          {
            width: 80,
            align: "right"
          }
        );


        doc.y =
          startY + 24;

      }


      // ==========================
      // 下線
      // ==========================

      doc
        .moveTo(
          55,
          doc.y
        )
        .lineTo(
          535,
          doc.y
        )
        .stroke();


      doc.moveDown(1);


      // ==========================
      // 合計金額
      // ==========================

      doc
        .fontSize(18)
        .text(
          `合計金額：${formatNumber(
            total
          )}円`,
          {
            align: "right"
          }
        );


      doc.moveDown(2);


      // ==========================
      // 納期
      // ==========================

      doc
        .fontSize(15)
        .text(
          "納期",
          {
            underline: true
          }
        );


      doc.moveDown(0.5);


      doc
        .fontSize(12)
        .text(
          delivery
        );


      doc.moveDown(2);


      // ==========================
      // 備考
      // ==========================

      doc
        .fontSize(15)
        .text(
          "備考",
          {
            underline: true
          }
        );


      doc.moveDown(0.5);


      doc
        .fontSize(12)
        .text(
          estimate.note || "なし"
        );


      // ==========================
      // PDF終了
      // ==========================

      doc.end();


      // ==========================
      // PDF完成後に返す
      // ==========================

      stream.on(
        "finish",
        () => {

          console.log(
            "御見積書PDFを作成しました:",
            pdfFileName
          );


          res.json({

            success: true,

            message:
              "御見積書PDFを作成しました",

            fileName:
              pdfFileName,

            url:
              `/pdf/${encodeURIComponent(
                pdfFileName
              )}`

          });

        }
      );


      stream.on(
        "error",
        (error) => {

          console.error(
            "PDF保存エラー:",
            error
          );

          if (
            !res.headersSent
          ) {

            res.status(500).json({

              success: false,

              message:
                "PDFの保存に失敗しました"

            });

          }

        }
      );


    } catch (error) {

      console.error(
        "PDF作成エラー:",
        error
      );

      if (
        !res.headersSent
      ) {

        res.status(500).json({

          success: false,

          message:
            "御見積書PDFの作成に失敗しました"

        });

      }

    }

  }
);


// ==============================
// PDF公開
// ==============================

app.use(
  "/pdf",
  express.static(
    PDF_DIR
  )
);


// ==============================
// サーバー起動
// ==============================

app.listen(
  PORT,
  () => {

    console.log(
      `Server running on port ${PORT}`
    );

    console.log(
      `Excel: ${EXCEL_FILE}`
    );

    console.log(
      `商品シート: ${PRODUCT_SHEET}`
    );

    console.log(
      `PDF directory: ${PDF_DIR}`
    );

    console.log(
      `Japanese font: ${
        JAPANESE_FONT || "NOT FOUND"
      }`
    );

  }
);
