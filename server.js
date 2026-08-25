const express = require("express");
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const PDFDocument = require("pdfkit");
const { Pool } = require("pg");

const app = express();

const PORT = process.env.PORT || 3000;

// ==============================
// 基本設定
// ==============================

app.use(express.json({ limit: "2mb" }));
app.use(express.urlencoded({ extended: true }));

app.use(express.static(__dirname));


// ==============================
// PostgreSQL
// ==============================

if (!process.env.DATABASE_URL) {
  console.error("DATABASE_URL が設定されていません。");
}

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,

  ssl:
    process.env.NODE_ENV === "production"
      ? {
          rejectUnauthorized: false
        }
      : false
});


// ==============================
// Excel商品マスタ
// ==============================
//
// ※ 初回の商品データ取り込み用として残してあります。
// 今後の商品データの保存先はPostgreSQLです。
// ==============================

const EXCEL_FILE = path.join(
  __dirname,
  "価格表.xlsm"
);

const PRODUCT_SHEET = "Sheet2";


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
// PostgreSQL テーブル作成
// ==============================

async function initializeDatabase() {

  if (!process.env.DATABASE_URL) {

    throw new Error(
      "DATABASE_URL が設定されていません。"
    );

  }

  const client =
    await pool.connect();

  try {

    await client.query("BEGIN");


    // ==========================
    // 商品テーブル
    // ==========================

    await client.query(`
      CREATE TABLE IF NOT EXISTS products (
        id SERIAL PRIMARY KEY,
        code TEXT NOT NULL DEFAULT '',
        size TEXT NOT NULL DEFAULT '',
        a TEXT NOT NULL DEFAULT '',
        price NUMERIC NOT NULL DEFAULT 0,
        brand TEXT NOT NULL DEFAULT '',
        pattern TEXT NOT NULL DEFAULT '',
        created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
        updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
      )
    `);


    // ==========================
    // 見積テーブル
    // ==========================

    await client.query(`
      CREATE TABLE IF NOT EXISTS estimates (
        id BIGINT PRIMARY KEY,
        company TEXT NOT NULL DEFAULT '',
        phone TEXT NOT NULL DEFAULT '',
        email TEXT NOT NULL DEFAULT '',
        note TEXT NOT NULL DEFAULT '',
        items JSONB NOT NULL DEFAULT '[]'::jsonb,
        delivery TEXT NOT NULL DEFAULT '',
        created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
        delivery_updated_at TIMESTAMPTZ
      )
    `);


    await client.query("COMMIT");

    console.log(
      "PostgreSQLのテーブルを確認しました。"
    );

  } catch (error) {

    await client.query("ROLLBACK");

    throw error;

  } finally {

    client.release();

  }

}


// ==============================
// PostgreSQL接続確認
// ==============================

async function checkDatabaseConnection() {

  const result =
    await pool.query(
      "SELECT NOW() AS now"
    );

  console.log(
    "PostgreSQL connected:",
    result.rows[0].now
  );

}


// ============================================================
// 商品
// ============================================================


// ==============================
// Excelから商品データを読み込む
// ==============================
//
// 初回取り込みなどに使用できます。
// 通常のAPIではPostgreSQLを使用します。
// ==============================

function loadProductsFromExcel() {

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
    .filter((p) => {

      return Object.values(p)
        .some(
          (value) =>
            String(
              value ?? ""
            ).trim() !== ""
        );

    });

}


// ==============================
// PostgreSQLから商品データ取得
// ==============================

async function getProducts() {

  const result =
    await pool.query(`
      SELECT
        id,
        code,
        size,
        a,
        price,
        brand,
        pattern
      FROM products
      ORDER BY id ASC
    `);

  return result.rows.map(
    (row) => ({

      id:
        row.id,

      code:
        row.code,

      size:
        row.size,

      a:
        row.a,

      price:
        Number(row.price),

      brand:
        row.brand,

      pattern:
        row.pattern

    })
  );

}


// ==============================
// 商品一覧API
// ==============================

app.get(
  "/api/products",
  async (req, res) => {

    try {

      const products =
        await getProducts();

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
          "商品データの読み込みに失敗しました"

      });

    }

  }
);


// ==============================
// 商品データ保存API
// ==============================
//
// products配列をPostgreSQLに保存します。
// 既存商品は一度削除してから入れ直します。
// ==============================

app.post(
  "/api/products",
  async (req, res) => {

    const client =
      await pool.connect();

    try {

      const products =
        Array.isArray(
          req.body.products
        )
          ? req.body.products
          : [];


      await client.query("BEGIN");


      // 現在の商品データを削除
      await client.query(
        "DELETE FROM products"
      );


      // 新しい商品データを登録
      for (
        const product of products
      ) {

        await client.query(
          `
          INSERT INTO products
          (
            code,
            size,
            a,
            price,
            brand,
            pattern
          )
          VALUES
          (
            $1,
            $2,
            $3,
            $4,
            $5,
            $6
          )
          `,
          [

            product.code ?? "",

            product.size ?? "",

            product.a ?? "",

            Number(
              product.price
            ) || 0,

            product.brand ?? "",

            product.pattern ?? ""

          ]
        );

      }


      await client.query("COMMIT");


      console.log(
        "商品データをPostgreSQLに保存しました"
      );


      res.json({

        success: true,

        message:
          "商品データを保存しました",

        count:
          products.length

      });

    } catch (error) {

      await client.query("ROLLBACK");

      console.error(
        "商品データ保存エラー:",
        error
      );

      res.status(500).json({

        success: false,

        message:
          "商品データの保存に失敗しました"

      });

    } finally {

      client.release();

    }

  }
);


// ==============================
// Excel → PostgreSQL 初回取り込みAPI
// ==============================
//
// 必要なときだけPOSTすると、
// 価格表.xlsmの内容をPostgreSQLへ取り込みます。
// ==============================

app.post(
  "/api/products/import-excel",
  async (req, res) => {

    const client =
      await pool.connect();

    try {

      const products =
        loadProductsFromExcel();


      await client.query("BEGIN");


      await client.query(
        "DELETE FROM products"
      );


      for (
        const product of products
      ) {

        await client.query(
          `
          INSERT INTO products
          (
            code,
            size,
            a,
            price,
            brand,
            pattern
          )
          VALUES
          (
            $1,
            $2,
            $3,
            $4,
            $5,
            $6
          )
          `,
          [

            product.code ?? "",

            product.size ?? "",

            product.a ?? "",

            Number(
              product.price
            ) || 0,

            product.brand ?? "",

            product.pattern ?? ""

          ]
        );

      }


      await client.query("COMMIT");


      console.log(
        `Excelから${products.length}件の商品を取り込みました`
      );


      res.json({

        success: true,

        message:
          "Excelの商品データを取り込みました",

        count:
          products.length

      });

    } catch (error) {

      await client.query("ROLLBACK");

      console.error(
        "Excel取り込みエラー:",
        error
      );

      res.status(500).json({

        success: false,

        message:
          error.message ||
          "Excelの商品データ取り込みに失敗しました"

      });

    } finally {

      client.release();

    }

  }
);


// ============================================================
// 見積
// ============================================================


// ==============================
// 見積データ一覧取得
// ==============================

async function getEstimates() {

  const result =
    await pool.query(`
      SELECT
        id,
        company,
        phone,
        email,
        note,
        items,
        delivery,
        created_at,
        delivery_updated_at
      FROM estimates
      ORDER BY created_at DESC
    `);


  return result.rows.map(
    (row) => ({

      id:
        Number(row.id),

      company:
        row.company,

      phone:
        row.phone,

      email:
        row.email,

      note:
        row.note,

      items:
        Array.isArray(row.items)
          ? row.items
          : [],

      delivery:
        row.delivery,

      createdAt:
        row.created_at
          ? new Date(
              row.created_at
            ).toISOString()
          : "",

      deliveryUpdatedAt:
        row.delivery_updated_at
          ? new Date(
              row.delivery_updated_at
            ).toISOString()
          : null

    })
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
          String(
            body.company || ""
          ),

        phone:
          String(
            body.phone || ""
          ),

        email:
          String(
            body.email || ""
          ),

        note:
          String(
            body.note || ""
          ),

        items:
          Array.isArray(
            body.items
          )
            ? body.items
            : [],

        delivery:
          String(
            body.delivery || ""
          ),

        createdAt:
          new Date().toISOString()

      };


      await pool.query(
        `
        INSERT INTO estimates
        (
          id,
          company,
          phone,
          email,
          note,
          items,
          delivery,
          created_at
        )
        VALUES
        (
          $1,
          $2,
          $3,
          $4,
          $5,
          $6::jsonb,
          $7,
          $8
        )
        `,
        [

          estimate.id,

          estimate.company,

          estimate.phone,

          estimate.email,

          estimate.note,

          JSON.stringify(
            estimate.items
          ),

          estimate.delivery,

          estimate.createdAt

        ]
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
  async (req, res) => {

    try {

      const estimates =
        await getEstimates();


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
  async (req, res) => {

    try {

      const id =
        Number(
          req.params.id
        );


      const delivery =
        String(
          req.body?.delivery || ""
        ).trim();


      const result =
        await pool.query(
          `
          UPDATE estimates
          SET
            delivery = $1,
            delivery_updated_at = NOW()
          WHERE id = $2
          RETURNING
            id,
            company,
            phone,
            email,
            note,
            items,
            delivery,
            created_at,
            delivery_updated_at
          `,
          [
            delivery,
            id
          ]
        );


      if (
        result.rowCount === 0
      ) {

        return res.status(404).json({

          success: false,

          message:
            "指定された見積依頼が見つかりません"

        });

      }


      console.log(
        "納期連絡を更新しました:",
        id
      );


      res.json({

        success: true,

        message:
          "納期連絡を保存しました",

        estimate:
          result.rows[0]

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
// 見積依頼削除API
// ==============================
//
// 「削除する」と明示的に呼び出したときだけ削除。
// 自動削除はしません。
// ==============================

app.delete(
  "/api/estimates/:id",
  async (req, res) => {

    try {

      const id =
        Number(
          req.params.id
        );


      if (
        !Number.isSafeInteger(id)
      ) {

        return res.status(400).json({

          success: false,

          message:
            "不正な見積番号です"

        });

      }


      const result =
        await pool.query(
          `
          DELETE FROM estimates
          WHERE id = $1
          RETURNING id
          `,
          [id]
        );


      if (
        result.rowCount === 0
      ) {

        return res.status(404).json({

          success: false,

          message:
            "指定された見積依頼が見つかりません"

        });

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


// ============================================================
// PDF
// ============================================================


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
  async (req, res) => {

    try {

      const id =
        Number(
          req.params.id
        );


      const result =
        await pool.query(
          `
          SELECT
            id,
            company,
            phone,
            email,
            note,
            items,
            delivery,
            created_at
          FROM estimates
          WHERE id = $1
          `,
          [id]
        );


      if (
        result.rowCount === 0
      ) {

        return res.status(404).json({

          success: false,

          message:
            "指定された見積依頼が見つかりません"

        });

      }


      const row =
        result.rows[0];


      const estimate = {

        id:
          Number(row.id),

        company:
          row.company,

        phone:
          row.phone,

        email:
          row.email,

        note:
          row.note,

        items:
          Array.isArray(row.items)
            ? row.items
            : [],

        delivery:
          row.delivery,

        createdAt:
          row.created_at
            ? new Date(
                row.created_at
              ).toISOString()
            : ""

      };


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


      // ==========================
      // 納期をDBへ保存
      // ==========================

      await pool.query(
        `
        UPDATE estimates
        SET
          delivery = $1,
          delivery_updated_at = NOW()
        WHERE id = $2
        `,
        [
          delivery,
          id
        ]
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
// ヘルスチェック
// ==============================

app.get(
  "/api/health",
  async (req, res) => {

    try {

      await pool.query(
        "SELECT 1"
      );

      res.json({

        success: true,

        database:
          "connected"

      });

    } catch (error) {

      console.error(
        "DB health check error:",
        error
      );

      res.status(500).json({

        success: false,

        database:
          "disconnected"

      });

    }

  }
);


// ==============================
// サーバー起動
// ==============================

async function startServer() {

  try {

    await checkDatabaseConnection();

    await initializeDatabase();


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

  } catch (error) {

    console.error(
      "サーバー起動エラー:",
      error
    );

    process.exit(1);

  }

}


startServer();
