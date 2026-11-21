
/**
 * 注文選択専用API
 * 既存の見積データを変更せず、注文内容を別保存する
 */

module.exports = function registerOrderSelectionApi(app, pool) {

  // 注文用テーブルの準備
  const initializeOrderSelectionTables = async () => {

    await pool.query(`
      CREATE TABLE IF NOT EXISTS order_selection_orders (
        id BIGSERIAL PRIMARY KEY,
        estimate_id BIGINT NOT NULL UNIQUE
        REFERENCES estimates(id) ON DELETE CASCADE,  
        total NUMERIC NOT NULL DEFAULT 0,
        ordered_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
      );
    `);

    await pool.query(`
      CREATE TABLE IF NOT EXISTS order_selection_items (
        id BIGSERIAL PRIMARY KEY,
        order_id BIGINT NOT NULL
          REFERENCES order_selection_orders(id) ON DELETE CASCADE,
        estimate_item_id BIGINT,
        product_code TEXT,
        size TEXT,
        brand TEXT,
        pattern TEXT,
        was_ordered BOOLEAN NOT NULL DEFAULT false,
        ordered_quantity INTEGER NOT NULL DEFAULT 0,
        unit_price NUMERIC NOT NULL DEFAULT 0,
        subtotal NUMERIC NOT NULL DEFAULT 0,
        remark TEXT DEFAULT ''
      );
    `);

  };

  const tablesReady = initializeOrderSelectionTables();

  
  // お客様画面用：見積商品と商品IDを取得
  app.get(
    "/api/order-selection/:accessToken",
    async (req, res) => {

      try {

        await tablesReady;

        const accessToken =
          String(req.params.accessToken || "").trim();

        const estimateResult = await pool.query(
          `
            SELECT
              id,
              company,
              phone,
              email,
              status,
              delivery,
              note,
              ordered_at AS "orderedAt"
            FROM estimates
            WHERE access_token = $1
            LIMIT 1
          `,
          [accessToken]
        );

        if (!estimateResult.rowCount) {
          return res.status(404).json({
            success: false,
            message: "見積データが見つかりません"
          });
        }

        const estimate = estimateResult.rows[0];

        // すでに注文済みなら保存済みの注文内容を取得
        if (estimate.status === "ordered") {

          const orderResult = await pool.query(
            `
              SELECT
                id,
                total,
                ordered_at AS "orderedAt"
              FROM order_selection_orders
              WHERE estimate_id = $1
              LIMIT 1
            `,
            [estimate.id]
          );

          if (orderResult.rowCount) {

            const order = orderResult.rows[0];

            const itemsResult = await pool.query(
              `
                SELECT
                  estimate_item_id AS id,
                  product_code AS code,
                  size,
                  brand,
                  pattern,
                  was_ordered AS "wasOrdered",
                  ordered_quantity AS "orderedQuantity",
                  unit_price AS price,
                  subtotal,
                  remark
                FROM order_selection_items
                WHERE order_id = $1
                ORDER BY id
              `,
              [order.id]
            );

            estimate.items = itemsResult.rows.map(item => ({
              id: Number(item.id),
              code: item.code || "",
              size: item.size || "",
              brand: item.brand || "",
              pattern: item.pattern || "",
              price: Number(item.price) || 0,
              qty: Number(item.orderedQuantity) || 0,
              wasOrdered: item.wasOrdered,
              subtotal: Number(item.subtotal) || 0,
              remark: item.remark || ""
            }));

            estimate.order = {
              id: Number(order.id),
              total: Number(order.total) || 0,
              orderedAt: order.orderedAt
            };

          }

        } else {

          // 注文前は元の見積商品を取得
          const itemsResult = await pool.query(
            `
              SELECT
                id,
                product_code AS code,
                size,
                brand,
                pattern,
                price,
                quantity AS qty
              FROM estimate_items
              WHERE estimate_id = $1
              ORDER BY id
            `,
            [estimate.id]
          );

          estimate.items = itemsResult.rows.map(item => ({
            id: Number(item.id),
            code: item.code || "",
            size: item.size || "",
            brand: item.brand || "",
            pattern: item.pattern || "",
            price: Number(item.price) || 0,
            qty: Number(item.qty) || 1
          }));

        }

        return res.json({
          success: true,
          estimate
        });

      } catch (error) {

        console.error("注文選択データ取得エラー:", error);

        return res.status(500).json({
          success: false,
          message: "注文内容の取得に失敗しました"
        });

      }

    }
  );

  // お客様画面用：選択商品の注文確定
  app.post(
    "/api/order-selection/:accessToken",
    async (req, res) => {

      const client = await pool.connect();

      let notificationData = null;

      try {

        await tablesReady;

        const accessToken =
          String(req.params.accessToken || "").trim();

        const selectedItems = req.body?.items;

        if (
          !Array.isArray(selectedItems) ||
          selectedItems.length === 0
        ) {
          return res.status(400).json({
            success: false,
            message: "注文する商品を1つ以上選択してください"
          });
        }

        await client.query("BEGIN");

        // 見積をロックして二重注文を防ぐ
        const estimateResult = await client.query(
          `
            SELECT id, company, phone, email, status
            FROM estimates
            WHERE access_token = $1
            FOR UPDATE
          `,
          [accessToken]
        );

        if (!estimateResult.rowCount) {
          await client.query("ROLLBACK");

          return res.status(404).json({
            success: false,
            message: "見積データが見つかりません"
          });
        }

        const estimate = estimateResult.rows[0];

        if (estimate.status === "ordered") {
          await client.query("ROLLBACK");

          return res.status(409).json({
            success: false,
            message: "この見積はすでに注文済みです"
          });
        }

        // 見積に登録されている商品をDBから取得
        const estimateItemsResult = await client.query(
          `
            SELECT
              id,
              product_code,
              size,
              brand,
              pattern,
              price,
              quantity
            FROM estimate_items
            WHERE estimate_id = $1
            ORDER BY id
          `,
          [estimate.id]
        );

        const estimateItems = estimateItemsResult.rows;

        if (!estimateItems.length) {
          await client.query("ROLLBACK");

          return res.status(400).json({
            success: false,
            message: "見積商品がありません"
          });
        }

        // 選択内容を検証
        const selectedMap = new Map();

        for (const selected of selectedItems) {

          const itemId = Number(selected.estimateItemId);
          const quantity = Number(selected.quantity);
          const remark = String(selected.remark || "").trim();

          if (
            !Number.isSafeInteger(itemId) ||
            !Number.isInteger(quantity) ||
            quantity < 1 ||
            quantity > 100000 ||
            remark.length > 1000 ||
            selectedMap.has(itemId)
          ) {
            await client.query("ROLLBACK");

            return res.status(400).json({
              success: false,
              message: "注文商品の指定内容が正しくありません"
            });
          }

          selectedMap.set(itemId, {
            quantity,
            remark
          });

        }

        // 見積に存在しない商品IDを拒否
        const estimateItemIds = new Set(
          estimateItems.map(item => Number(item.id))
        );

        for (const itemId of selectedMap.keys()) {

          if (!estimateItemIds.has(itemId)) {
            await client.query("ROLLBACK");

            return res.status(400).json({
              success: false,
              message: "見積にない商品が指定されています"
            });
          }

        }

        const orderedAt = new Date();

        // 注文ヘッダーを保存
        const orderResult = await client.query(
          `
            INSERT INTO order_selection_orders (
              estimate_id,
              total,
              ordered_at
            )
            VALUES ($1, 0, $2)
            RETURNING id
          `,
          [estimate.id, orderedAt]
        );

        const orderId = orderResult.rows[0].id;

        let total = 0;
        const notificationItems = [];

        // 注文した商品・注文しなかった商品をすべて保存
        for (const item of estimateItems) {

          const selected = selectedMap.get(Number(item.id));
          const wasOrdered = Boolean(selected);

          const quantity = wasOrdered
            ? selected.quantity
            : 0;

          const price = Number(item.price) || 0;
          const subtotal = wasOrdered
            ? price * quantity
            : 0;

          const remark = wasOrdered
            ? selected.remark
            : "";

          total += subtotal;

          await client.query(
            `
              INSERT INTO order_selection_items (
                order_id,
                estimate_item_id,
                product_code,
                size,
                brand,
                pattern,
                was_ordered,
                ordered_quantity,
                unit_price,
                subtotal,
                remark
              )
              VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11)
            `,
            [
              orderId,
              item.id,
              item.product_code || "",
              item.size || "",
              item.brand || "",
              item.pattern || "",
              wasOrdered,
              quantity,
              price,
              subtotal,
              remark
            ]
          );

          if (wasOrdered) {
            notificationItems.push({
              code: item.product_code || "",
              size: item.size || "",
              brand: item.brand || "",
              pattern: item.pattern || "",
              quantity,
              price,
              subtotal,
              remark
            });
          }

        }

        // 合計金額を保存
        await client.query(
          `
            UPDATE order_selection_orders
            SET total = $1
            WHERE id = $2
          `,
          [total, orderId]
        );

        // 既存の注文管理にも記録
        await client.query(
          `
            INSERT INTO orders (
              estimate_id,
              status,
              created_at
            )
            VALUES ($1, 'ordered', $2)
          `,
          [estimate.id, orderedAt]
        );

        // 見積を注文済みにする
        await client.query(
          `
            UPDATE estimates
            SET
              status = 'ordered',
              ordered_at = $1
            WHERE id = $2
          `,
          [orderedAt, estimate.id]
        );

        await client.query("COMMIT");

        notificationData = {
          estimate,
          orderedAt,
          notificationItems,
          total
        };

        // 注文完了をLINEで通知
        const adminUserId =
          process.env.LINE_ADMIN_USER_ID;

        const channelAccessToken =
          process.env.LINE_CHANNEL_ACCESS_TOKEN;

        if (adminUserId && channelAccessToken) {

          try {

            const itemText = notificationItems.map(item => {

              return (
                `・${item.code} ${item.size}\n` +
                `  ${item.brand} ${item.pattern}\n` +
                `  数量：${item.quantity}個\n` +
                `  単価：${item.price.toLocaleString("ja-JP")}円\n` +
                `  小計：${item.subtotal.toLocaleString("ja-JP")}円\n` +
                `  備考：${item.remark || "なし"}`
              );

            }).join("\n\n");

            const message =
              `🛒 注文が入りました！\n\n` +
              `見積番号：${estimate.id}\n` +
              `会社名・氏名：${estimate.company || ""}\n` +
              `電話番号：${estimate.phone || ""}\n\n` +
              `【注文商品】\n${itemText}\n\n` +
              `合計金額：${total.toLocaleString("ja-JP")}円（税抜）\n` +
              `注文日時：${orderedAt.toLocaleString("ja-JP")}`;

            const lineResponse = await fetch(
              "https://api.line.me/v2/bot/message/push",
              {
                method: "POST",
                headers: {
                  "Content-Type": "application/json",
                  "Authorization": `Bearer ${channelAccessToken}`
                },
                body: JSON.stringify({
                  to: adminUserId,
                  messages: [
                    {
                      type: "text",
                      text: message
                    }
                  ]
                })
              }
            );

            if (!lineResponse.ok) {
              console.error(
                "注文LINE通知エラー:",
                lineResponse.status,
                await lineResponse.text()
              );
            }

          } catch (lineError) {
            console.error("注文LINE通知送信エラー:", lineError);
          }

        }

        return res.json({
          success: true,
          message: "注文を受け付けました",
          estimate: {
            id: estimate.id,
            company: estimate.company,
            status: "ordered",
            orderedAt
          },
          order: {
            id: orderId,
            total,
            orderedAt,
            items: notificationItems
          }
        });

      } catch (error) {

        await client.query("ROLLBACK");

        console.error("注文選択APIエラー:", error);

        return res.status(500).json({
          success: false,
          message: "注文の受付に失敗しました"
        });

      } finally {

        client.release();

      }

    }
  );


  // 管理画面用：注文内容を取得
  app.get(
    "/api/admin/order-selection/:estimateId",
    async (req, res) => {

      try {

        await tablesReady;

        const estimateId = Number(req.params.estimateId);

        if (!Number.isSafeInteger(estimateId)) {
          return res.status(400).json({
            success: false,
            message: "見積番号が正しくありません"
          });
        }

        const orderResult = await pool.query(
          `
            SELECT
              id,
              estimate_id AS "estimateId",
              total,
              ordered_at AS "orderedAt"
            FROM order_selection_orders
            WHERE estimate_id = $1
          `,
          [estimateId]
        );

        if (!orderResult.rowCount) {
          return res.json({
            success: true,
            order: null
          });
        }

        const order = orderResult.rows[0];

        const itemsResult = await pool.query(
          `
            SELECT
              estimate_item_id AS "estimateItemId",
              product_code AS code,
              size,
              brand,
              pattern,
              was_ordered AS "wasOrdered",
              ordered_quantity AS "orderedQuantity",
              unit_price AS "unitPrice",
              subtotal,
              remark
            FROM order_selection_items
            WHERE order_id = $1
            ORDER BY id
          `,
          [order.id]
        );

        order.total = Number(order.total) || 0;
        order.items = itemsResult.rows.map(item => ({
          ...item,
          unitPrice: Number(item.unitPrice) || 0,
          subtotal: Number(item.subtotal) || 0
        }));

        return res.json({
          success: true,
          order
        });

      } catch (error) {

        console.error("管理画面注文取得エラー:", error);

        return res.status(500).json({
          success: false,
          message: "注文内容の取得に失敗しました"
        });

      }

    }
  );

  console.log("注文選択APIを登録しました");

};
