const express = require("express");
const fs = require("fs");
const path = require("path");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// サイト本体を表示
app.use(express.static(__dirname));

// 見積依頼を受け取る
app.post("/api/estimate", (req, res) => {
  const estimate = {
    id: Date.now(),
    company: req.body.company || "",
    phone: req.body.phone || "",
    email: req.body.email || "",
    note: req.body.note || "",
    items: req.body.items || [],
    createdAt: new Date().toISOString()
  };

  const filePath = path.join(__dirname, "estimates.json");

  let estimates = [];

  if (fs.existsSync(filePath)) {
    try {
      estimates = JSON.parse(fs.readFileSync(filePath, "utf8"));
    } catch (error) {
      estimates = [];
    }
  }

  estimates.push(estimate);

  fs.writeFileSync(
    filePath,
    JSON.stringify(estimates, null, 2),
    "utf8"
  );

  res.json({
    success: true,
    message: "見積依頼を受け付けました"
  });
});

// 管理画面から見積一覧を取得
app.get("/api/estimates", (req, res) => {
  const filePath = path.join(__dirname, "estimates.json");

  if (!fs.existsSync(filePath)) {
    return res.json([]);
  }

  try {
    const estimates = JSON.parse(
      fs.readFileSync(filePath, "utf8")
    );

    res.json(estimates);
  } catch (error) {
    res.status(500).json({
      success: false,
      message: "データを読み込めませんでした"
    });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
