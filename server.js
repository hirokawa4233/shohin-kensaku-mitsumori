const express = require("express");
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const PDFDocument = require("pdfkit");
const https = require("https");

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

const ESTIMATE_FILE = path.join(__dirname, "estimates.json");


// ==============================
// PDF保存先
// ==============================

const PDF_DIR = path.join(__dirname, "pdf");

if (!fs.existsSync(PDF_DIR)) {
  fs.mkdirSync(PDF_DIR, { recursive: true });
}


// ==============================
// 日本語フォント
// ==============================

const FONT_DIR = path.join(__dirname, "fonts");
const JAPANESE_FONT = path.join(
  FONT_DIR,
  "NotoSansJP-Regular.ttf"
);

if (!fs.existsSync(FONT_DIR)) {
  fs.mkdirSync(FONT_DIR, { recursive: true });
}


// ==============================
// 日本語フォントを取得
// ==============================

function downloadJapaneseFont() {

  return new Promise((resolve, reject) => {

    if (fs.existsSync(JAPANESE_FONT)) {
      console.log(
        "PDF font:",
        JAPANESE_FONT
      );

      return resolve();
    }

    console.log(
      "日本語フォントをダウンロードしています..."
    );

    const fontUrl =
      "https://github.com/google/fonts/raw/main/ofl/notosansjp/NotoSansJP%5Bwght%5D.ttf";

    const file =
      fs.createWriteStream(
        JAPANESE_FONT
      );

    https.get(
      fontUrl,
      (response) => {

        if (
          response.statusCode >= 300 &&
          response.statusCode < 400 &&
          response.headers.location
        ) {

          file.close();

          if (fs.existsSync(JAPANESE_FONT)) {
            fs.unlinkSync(JAPANESE_FONT);
          }

          https.get(
            response.headers.location,
            (redirectResponse) => {

              redirectResponse.pipe(file);

              file.on(
                "finish",
                () => {

                  file.close();

                  console.log(
                    "日本語フォント取得完了:",
                    JAPANESE_FONT
                  );

                  resolve();

                }
              );

            }
          ).on(
            "error",
            (error) => {

              file.close();

              if (
                fs.existsSync(JAPANESE_FONT)
              ) {
                fs.unlinkSync(
                  JAPANESE_FONT
                );
              }

              reject(error);

            }
          );

          return;
        }


        if (
          response.statusCode !== 200
        ) {

          file.close();

          if (
            fs.existsSync(JAPANESE_FONT)
          ) {
            fs.unlinkSync(
              JAPANESE_FONT
            );
          }

          return reject(
            new Error(
              `フォント取得失敗: HTTP ${response.statusCode}`
            )
          );

        }


        response.pipe(file);

        file.on(
          "finish",
          () => {

            file.close();

            console.log(
              "日本語フォント取得完了:",
              JAPANESE_FONT
            );

            resolve();

          }
        );

      }
    ).on(
      "error",
      (error) => {

        file.close();

        if (
          fs.existsSync(JAPANESE_FONT)
        ) {
          fs.unlinkSync(
            JAPANESE_FONT
          );
        }

        reject(error);

      }
    );

  });

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
      .map(
        (row) => ({

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

        })
      )
      .filter(
        (p) =>
          Object.values(p).some(
            (value) =>
              String(
                value ?? ""
              ).trim() !== ""
          )
      );

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
  (req, res) => {

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
// 見積依頼一覧API
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
// 納期連絡保存
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

function formatNumber(
  value
) {

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
  async (req, res) => {

    try {

      // --------------------------
      // 日本語フォント確認
      // --------------------------

      await downloadJapaneseFont();


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


      // --------------------------
      // 納期
      // --------------------------

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


      // --------------------------
      // PDFファイル名
      // --------------------------

      const pdfFileName =
        `御見積書_${estimate.id}.pdf`;

      const pdfPath =
        path.join(
          PDF_DIR,
          pdfFileName
        );


      console.log(
        "PDF作成リクエスト:",
        estimate.id
      );

      console.log(
        "PDF path:",
        pdfPath
      );


      // --------------------------
      // PDF作成
      // --------------------------

      const doc =
        new PDFDocument({

          size: "A4",

          margin: 55

        });


      const stream =
        fs.createWriteStream(
          pdfPath
        );

      doc.pipe(stream);


      // --------------------------
      // 日本語フォント
      // --------------------------

      doc.font(
        JAPANESE_FONT
      );


      console.log(
        "PDF font:",
        JAPANESE_FONT
      );


      // --------------------------
      // タイトル
      // --------------------------

      doc
        .fontSize(28)
        .text(
          "御 見 積 書",
          {
            align: "center"
          }
        );

      doc.moveDown(2);


      // --------------------------
      // 見積情報
      // --------------------------

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


      // --------------------------
      // お客様情報
      // --------------------------

      doc
        .fontSize(16)
        .text(
          "お客様情報",
          {
            underline: true
          }
        );

      doc.moveDown(0.5);


      const customerName =
        String(
          estimate.company || ""
        ).trim();


      // ★ お客様名には必ず「様」
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


      // --------------------------
      // お見積内容
      // --------------------------

      doc
        .fontSize(16)
        .text(
          "お見積内容",
          {
            underline: true
          }
        );

      doc.moveDown(0.8);


      // --------------------------
      // 表ヘッダー
      // --------------------------

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
        headerY + 22;


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

      doc.moveDown(0.5);


      // --------------------------
      // 商品
      // --------------------------

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


      // --------------------------
      // 下線
      // --------------------------

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


      // --------------------------
      // 合計金額
      // --------------------------

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


      // --------------------------
      // 納期
      // --------------------------

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


      // --------------------------
      // 備考
      // --------------------------

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
          estimate.note ||
          "なし"
        );


      // --------------------------
      // PDF終了
      // --------------------------

      doc.end();


      // --------------------------
      // PDF保存完了
      // --------------------------

      stream.on(
        "finish",
        () => {

          try {

            const stat =
              fs.statSync(
                pdfPath
              );

            console.log(
              "御見積書PDFを作成しました:",
              pdfFileName
            );

            console.log(
              "PDF size:",
              stat.size,
              "bytes"
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

          } catch (error) {

            console.error(
              "PDF確認エラー:",
              error
            );

            if (
              !res.headersSent
            ) {

              res.status(500).json({

                success: false,

                message:
                  "PDFファイルを確認できませんでした"

              });

            }

          }

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
      `Japanese font: ${JAPANESE_FONT}`
    );

  }
);
