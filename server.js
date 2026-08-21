const express = require("express");
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const PDFDocument = require("pdfkit");

const app = express();

const PORT = process.env.PORT || 3000;

// ========================================
// 基本設定
// ========================================

app.use(express.json({ limit: "2mb" }));
app.use(express.urlencoded({ extended: true }));

app.use(express.static(__dirname));


// ========================================
// Excel
// ========================================

const EXCEL_FILE = path.join(__dirname, "価格表.xlsm");
const PRODUCT_SHEET = "Sheet2";


// ========================================
// 見積データ
// ========================================

const ESTIMATE_FILE = path.join(
  __dirname,
  "estimates.json"
);


// ========================================
// PDF
// ========================================

const PDF_DIR = path.join(
  __dirname,
  "pdf"
);

if (!fs.existsSync(PDF_DIR)) {
  fs.mkdirSync(PDF_DIR, {
    recursive: true
  });
}


// ========================================
// 日本語フォント
// ★ ダウンロードしない
// ========================================

const FONT_CANDIDATES = [
  path.join(
    __dirname,
    "fonts",
    "NotoSansJP-Regular.ttf"
  ),

  "/usr/share/fonts/truetype/noto/NotoSansJP-Regular.ttf",

  "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",

  "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
];


function findJapaneseFont() {

  for (const font of FONT_CANDIDATES) {

    if (fs.existsSync(font)) {

      console.log(
        "Japanese font:",
        font
      );

      return font;
    }
  }

  console.error(
    "Japanese font: NOT FOUND"
  );

  return null;
}


const JAPANESE_FONT =
  findJapaneseFont();


// ========================================
// 商品データ読み込み
// ========================================

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
    XLSX.utils.sheet_to_json(
      sheet,
      {
        defval: "",
        raw: true
      }
    );

  return rows
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
    .filter((product) => {

      return Object.values(product)
        .some((value) => {

          return String(
            value ?? ""
          ).trim() !== "";

        });

    });

}


// ========================================
// 商品一覧API
// ========================================

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


// ========================================
// 見積データ読み込み
// ========================================

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


// ========================================
// 見積データ保存
// ========================================

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


// ========================================
// 見積依頼受付
// ========================================

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


// ========================================
// 見積一覧
// ========================================

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


// ========================================
// 納期保存
// ========================================

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
          (item) => {

            return Number(
              item.id
            ) === id;

          }
        );

      if (!estimate) {

        return res.status(404)
          .json({

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


// ========================================
// 数値フォーマット
// ========================================

function formatNumber(
  value
) {

  return Number(
    value || 0
  ).toLocaleString(
    "ja-JP"
  );

}


// ========================================
// PDF作成
// ========================================

app.post(
  "/api/estimates/:id/pdf",
  (req, res) => {

    const id =
      Number(
        req.params.id
      );

    console.log(
      "PDF作成リクエスト:",
      id
    );

    try {

      // --------------------------------
      // フォント確認
      // --------------------------------

      if (!JAPANESE_FONT) {

        console.error(
          "日本語フォントが見つかりません"
        );

        return res.status(500)
          .json({

            success: false,

            message:
              "日本語フォントが見つかりません"

          });

      }

      console.log(
        "使用フォント:",
        JAPANESE_FONT
      );


      // --------------------------------
      // 見積データ
      // --------------------------------

      const estimates =
        loadEstimates();

      const estimate =
        estimates.find(
          (item) => {

            return Number(
              item.id
            ) === id;

          }
        );

      if (!estimate) {

        return res.status(404)
          .json({

            success: false,

            message:
              "指定された見積依頼が見つかりません"

          });

      }


      // --------------------------------
      // 納期
      // --------------------------------

      const delivery =
        String(
          req.body?.delivery ??
          estimate.delivery ??
          ""
        ).trim();

      if (!delivery) {

        return res.status(400)
          .json({

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


      // --------------------------------
      // ファイル名
      // --------------------------------

      const pdfFileName =
        `御見積書_${estimate.id}.pdf`;

      const pdfPath =
        path.join(
          PDF_DIR,
          pdfFileName
        );

      console.log(
        "PDF path:",
        pdfPath
      );


      // --------------------------------
      // PDF作成
      // --------------------------------

      const doc =
        new PDFDocument({

          size: "A4",

          margin: 55,

          info: {

            Title:
              "御見積書",

            Author:
              "見積依頼管理システム"

          }

        });


      const stream =
        fs.createWriteStream(
          pdfPath
        );


      let finished = false;


      stream.on(
        "error",
        (error) => {

          console.error(
            "PDF書き込みエラー:",
            error
          );

          if (
            !res.headersSent
          ) {

            res.status(500)
              .json({

                success: false,

                message:
                  "PDFファイルの保存に失敗しました"

              });

          }

        }
      );


      stream.on(
        "finish",
        () => {

          try {

            const stat =
              fs.statSync(
                pdfPath
              );

            console.log(
              "PDF size:",
              stat.size,
              "bytes"
            );


            if (
              stat.size < 1000
            ) {

              console.error(
                "PDFサイズが小さすぎます:",
                stat.size
              );

              if (
                !res.headersSent
              ) {

                return res.status(500)
                  .json({

                    success: false,

                    message:
                      "PDFファイルが正しく生成されませんでした"

                  });

              }

              return;
            }


            finished = true;

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

          } catch (error) {

            console.error(
              "PDF確認エラー:",
              error
            );

            if (
              !res.headersSent
            ) {

              res.status(500)
                .json({

                  success: false,

                  message:
                    "PDFの確認に失敗しました"

                });

            }

          }

        }
      );


      doc.pipe(
        stream
      );


      // --------------------------------
      // 日本語フォント
      // --------------------------------

      doc.font(
        JAPANESE_FONT
      );


      // --------------------------------
      // タイトル
      // --------------------------------

      doc
        .fontSize(26)
        .text(
          "御 見 積 書",
          {
            align: "center"
          }
        );

      doc.moveDown(2);


      // --------------------------------
      // 見積情報
      // --------------------------------

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


      // --------------------------------
      // お客様情報
      // --------------------------------

      doc
        .fontSize(16)
        .text(
          "お客様情報",
          {
            underline: true
          }
        );

      doc.moveDown(0.5);


      doc
        .fontSize(12)
        .text(
          `会社名・氏名：${String(
            estimate.company || ""
          )} 様`
        );

      doc.text(
        `電話番号：${String(
          estimate.phone || ""
        )}`
      );

      doc.text(
        `メールアドレス：${String(
          estimate.email || ""
        )}`
      );


      doc.moveDown(2);


      // --------------------------------
      // お見積内容
      // --------------------------------

      doc
        .fontSize(16)
        .text(
          "お見積内容",
          {
            underline: true
          }
        );

      doc.moveDown(0.8);


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


      doc.y += 12;


      // --------------------------------
      // 商品
      // --------------------------------

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


        const y =
          doc.y;


        doc
          .fontSize(10)
          .text(
            String(
              item.code || ""
            ),
            55,
            y,
            {
              width: 100
            }
          );

        doc.text(
          String(
            item.size || ""
          ),
          155,
          y,
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
          y,
          {
            width: 170
          }
        );

        doc.text(
          `${qty}個`,
          425,
          y,
          {
            width: 45
          }
        );

        doc.text(
          `${formatNumber(
            subtotal
          )}円`,
          470,
          y,
          {
            width: 80,
            align: "right"
          }
        );


        doc.y =
          y + 28;

      }


      // --------------------------------
      // 合計
      // --------------------------------

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


      // --------------------------------
      // 納期
      // --------------------------------

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


      // --------------------------------
      // 備考
      // --------------------------------

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
          String(
            estimate.note || "なし"
          )
        );


      // --------------------------------
      // PDF終了
      // --------------------------------

      doc.end();

    } catch (error) {

      console.error(
        "PDF作成エラー:",
        error
      );

      if (
        !res.headersSent
      ) {

        res.status(500)
          .json({

            success: false,

            message:
              "御見積書PDFの作成に失敗しました",

            error:
              error.message

          });

      }

    }

  }
);


// ========================================
// PDF公開
// ========================================

app.use(
  "/pdf",
  express.static(
    PDF_DIR,
    {
      fallthrough: false
    }
  )
);


// ========================================
// サーバー起動
// ========================================

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
