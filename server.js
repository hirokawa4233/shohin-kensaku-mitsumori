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
// PDF用フォント
// ==============================

const FONT_FILE = path.join(
  __dirname,
  "NotoSansJP-Regular.otf"
);

const FONT_URL =
  "https://github.com/notofonts/noto-cjk/raw/main/Sans/SubsetOTF/JP/NotoSansJP-Regular.otf";


// ==============================
// 日本語フォントを準備
// ==============================

function downloadFile(url, destination) {

  return new Promise((resolve, reject) => {

    const file =
      fs.createWriteStream(destination);

    https.get(url, (response) => {

      // リダイレクト対応
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

      if (response.statusCode !== 200) {

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

    }).on("error", (error) => {

      file.close();

      try {
        fs.unlinkSync(destination);
      } catch (_) {}

      reject(error);

    });

  });

}


async function ensureJapaneseFont() {

  if (fs.existsSync(FONT_FILE)) {
    return;
  }

  console.log(
    "日本語PDF用フォントを取得しています..."
  );

  await downloadFile(
    FONT_URL,
    FONT_FILE
  );

  console.log(
    "日本語PDF用フォントの準備が完了しました。"
  );

}


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

  const products =
    rows
      .map((row) => ({

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

      }))
      .filter((p) => {

        return Object.values(p).some(
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
      !Array.isArray(estimates)
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
// 金額を数値に変換
// ==============================

function toNumber(value) {

  if (
    typeof value === "number"
  ) {

    return value;

  }

  const number =
    Number(
      String(
        value ?? ""
      )
        .replace(/,/g, "")
        .replace(/円/g, "")
        .trim()
    );

  return Number.isFinite(number)
    ? number
    : 0;

}


// ==============================
// 金額表示
// ==============================

function yen(value) {

  return (
    toNumber(value)
      .toLocaleString("ja-JP")
    + "円"
  );

}


// ==============================
// 日付表示
// ==============================

function formatDate(iso) {

  if (!iso) {
    return "";
  }

  const date =
    new Date(iso);

  if (
    Number.isNaN(
      date.getTime()
    )
  ) {

    return String(iso);

  }

  return date.toLocaleString(
    "ja-JP",
    {
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit"
    }
  );

}


// ==============================
// お客様名
// 「様」を自動で付ける
// ==============================

function customerName(
  company
) {

  const name =
    String(
      company || ""
    ).trim();

  if (!name) {
    return "";
  }

  // すでに「様」が付いている場合は二重にしない
  if (
    name.endsWith("様")
  ) {

    return name;

  }

  return `${name} 様`;

}


// ==============================
// 御見積書PDF生成
// ==============================

app.post(
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

        return res.status(404).json({

          success: false,

          message:
            "指定された見積依頼が見つかりません"

        });

      }


      // ==============================
      // 納期が送られてきた場合は保存
      // ==============================

      if (
        req.body &&
        typeof req.body.delivery !==
          "undefined"
      ) {

        estimate.delivery =
          String(
            req.body.delivery || ""
          ).trim();

        estimate.deliveryUpdatedAt =
          new Date().toISOString();

        saveEstimates(
          estimates
        );

      }


      // ==============================
      // 日本語フォント準備
      // ==============================

      await ensureJapaneseFont();


      // ==============================
      // PDFファイル名
      // ==============================

      const fileName =
        `御見積書_${estimate.id}.pdf`;


      // ==============================
      // PDFレスポンス
      // ==============================

      res.setHeader(
        "Content-Type",
        "application/pdf"
      );

      res.setHeader(
        "Content-Disposition",
        `inline; filename="estimate-${estimate.id}.pdf"`
      );


      // ==============================
      // PDF作成
      // ==============================

      const doc =
        new PDFDocument({

          size: "A4",

          margins: {

            top: 50,
            bottom: 50,
            left: 50,
            right: 50

          },

          info: {

            Title:
              "御見積書",

            Author:
              "商品検索・見積依頼サイト"

          }

        });


      doc.pipe(res);


      // 日本語フォント
      doc.font(
        FONT_FILE
      );


      // ==============================
      // タイトル
      // ==============================

      doc
        .fontSize(24)
        .text(
          "御 見 積 書",
          {
            align: "center"
          }
        );

      doc.moveDown(1);


      // ==============================
      // 見積番号・受付日時
      // ==============================

      doc
        .fontSize(10)
        .text(
          `見積番号：${estimate.id}`,
          {
            align: "right"
          }
        );

      doc
        .text(
          `受付日時：${formatDate(
            estimate.createdAt
          )}`,
          {
            align: "right"
          }
        );

      doc.moveDown(1);


      // ==============================
      // お客様情報
      // ==============================

      doc
        .fontSize(13)
        .text(
          "お客様情報",
          {
            underline: true
          }
        );

      doc.moveDown(0.4);


      // ★ここで「様」を付ける
      doc
        .fontSize(11)
        .text(
          `会社名・氏名：${customerName(
            estimate.company
          )}`
        );


      doc
        .text(
          `電話番号：${estimate.phone || ""}`
        );


      doc
        .text(
          `メールアドレス：${estimate.email || ""}`
        );

      doc.moveDown(1);


      // ==============================
      // 商品一覧
      // ==============================

      doc
        .fontSize(13)
        .text(
          "お見積内容",
          {
            underline: true
          }
        );

      doc.moveDown(0.5);


      const items =
        Array.isArray(
          estimate.items
        )
          ? estimate.items
          : [];


      let total = 0;


      // ==============================
      // 表の見出し
      // ==============================

      doc
        .fontSize(9)
        .text(
          "品番",
          50,
          doc.y,
          {
            continued: true,
            width: 90
          }
        )
        .text(
          "サイズ",
          {
            continued: true,
            width: 80
          }
        )
        .text(
          "ブランド・パターン",
          {
            continued: true,
            width: 150
          }
        )
        .text(
          "数量",
          {
            continued: true,
            width: 45
          }
        )
        .text(
          "金額",
          {
            width: 100,
            align: "right"
          }
        );


      doc.moveDown(0.4);


      // 横線
      doc
        .moveTo(
          50,
          doc.y
        )
        .lineTo(
          545,
          doc.y
        )
        .stroke();


      doc.moveDown(0.5);


      // ==============================
      // 商品行
      // ==============================

      items.forEach(
        (item) => {

          const quantity =
            Math.max(
              1,
              parseInt(
                item.qty,
                10
              ) || 1
            );


          const unitPrice =
            toNumber(
              item.price
            );


          const itemTotal =
            unitPrice *
            quantity;


          total +=
            itemTotal;


          const brandPattern =
            [
              item.brand || "",
              item.pattern || ""
            ]
              .filter(Boolean)
              .join(" ");


          const startY =
            doc.y;


          doc
            .fontSize(9)
            .text(
              String(
                item.code || ""
              ),
              50,
              startY,
              {
                continued: true,
                width: 90
              }
            )
            .text(
              String(
                item.size || ""
              ),
              {
                continued: true,
                width: 80
              }
            )
            .text(
              brandPattern,
              {
                continued: true,
                width: 150
              }
            )
            .text(
              String(
                quantity
              ),
              {
                continued: true,
                width: 45
              }
            )
            .text(
              yen(itemTotal),
              {
                width: 100,
                align: "right"
              }
            );


          doc.moveDown(0.5);


          doc
            .moveTo(
              50,
              doc.y
            )
            .lineTo(
              545,
              doc.y
            )
            .stroke();


          doc.moveDown(0.5);

        }
      );


      // ==============================
      // 合計金額
      // ==============================

      doc.moveDown(0.5);

      doc
        .fontSize(16)
        .text(
          `合計金額：${yen(total)}`,
          {
            align: "right"
          }
        );


      // ==============================
      // 納期
      // ==============================

      doc.moveDown(1.5);

      doc
        .fontSize(13)
        .text(
          "納期",
          {
            underline: true
          }
        );

      doc.moveDown(0.4);

      doc
        .fontSize(11)
        .text(
          estimate.delivery ||
          "未定"
        );


      // ==============================
      // 備考
      // ==============================

      doc.moveDown(1.2);

      doc
        .fontSize(13)
        .text(
          "備考",
          {
            underline: true
          }
        );

      doc.moveDown(0.4);

      doc
        .fontSize(11)
        .text(
          estimate.note ||
          "なし"
        );


      // ==============================
      // フッター
      // ==============================

      doc
        .fontSize(9)
        .text(
          "本見積書は見積依頼内容をもとに作成されたものです。",
          50,
          760,
          {
            align: "center",
            width: 495
          }
        );


      // ==============================
      // PDF終了
      // ==============================

      doc.end();


      console.log(
        "御見積書PDFを作成しました:",
        fileName
      );


    } catch (error) {

      console.error(
        "PDF生成エラー:",
        error
      );

      if (!res.headersSent) {

        res.status(500).json({

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

  }
);
