/**
 * 第1回メルマガ（デザイン・文言再調整版）
 * シート前提：
 * A:名前 / B:社員番号 / C:生年月日 / D:メール / E:PDF送信日 / F:第1回済 / G:第2回済（推奨）
 */
function onFormSubmit(e) {
  sendConsentEmail(e);
}

/**
 * フォームの「同意する」回答者へ自動送信メール
 * 送信条件：同意質問に「同意する」と回答していること
 */
function sendConsentEmail(e) {
  if (!e || !e.namedValues) {
    console.warn("フォーム送信イベントが取得できません。");
    return;
  }

  // ==================================================
  // 【設定エリア】フォームの質問文に合わせて変更してください
  // ==================================================
  const LAST_NAME_QUESTION = "姓（漢字）（例：山田）";
  const FIRST_NAME_QUESTION = "名（漢字）（例：花子）";
  const EMAIL_QUESTION = "メールアドレス（※社内アドレスを推奨）";
  const CONSENT_QUESTION =
    "あなたはこの研究に参加するにあたり、上記の事項について十分な説明を受け、内容等を十分理解の上、本研究に参加することに同意しますか？";
  const CONSENT_VALUE = "同意する";
  const SUBJECT = "Well-Bone Health Study お申込みありがとうございます";
  // ==================================================

  const lastName = getFirstAnswer_(e.namedValues[LAST_NAME_QUESTION]);
  const firstName = getFirstAnswer_(e.namedValues[FIRST_NAME_QUESTION]);
  const email = getFirstAnswer_(e.namedValues[EMAIL_QUESTION]);
  const consent = getFirstAnswer_(e.namedValues[CONSENT_QUESTION]);

  if (!email) {
    console.warn("メールアドレスが取得できません。");
    return;
  }

  if (!consent || !consent.includes(CONSENT_VALUE)) {
    console.log(`同意が確認できないため送信をスキップしました: ${email}`);
    return;
  }

  const fullName = [lastName, firstName].filter(Boolean).join(" ");
  const displayName = fullName ? `${fullName} 様` : "参加者 様";
  const htmlBody = buildConsentEmailHtml_(displayName);

  GmailApp.sendEmail(email, SUBJECT, "HTMLメールを表示できる環境でご覧ください。", {
    htmlBody,
    name: "ウェルネス事務局",
  });
}

function buildConsentEmailHtml_(displayName) {
  return `
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Well-Bone Health Study</title>
    <style>
        body {
            margin: 0;
            padding: 0;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Hiragino Kaku Gothic ProN', 'Hiragino Sans', Meiryo, sans-serif;
            background-color: #f5f5f5;
            line-height: 1.6;
        }
        .container {
            max-width: 600px;
            margin: 0 auto;
            background-color: #ffffff;
        }
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: #ffffff;
            padding: 30px 20px;
            text-align: center;
        }
        .header h1 {
            margin: 0;
            font-size: 20px;
            font-weight: 600;
        }
        .header p {
            margin: 10px 0 0 0;
            font-size: 14px;
            opacity: 0.95;
        }
        .greeting {
            padding: 25px 20px 15px;
            font-size: 18px;
            font-weight: 600;
            color: #333333;
        }
        .content {
            padding: 0 20px 20px;
        }
        .gift-box {
            background: linear-gradient(135deg, #ffecd2 0%, #fcb69f 100%);
            border-radius: 12px;
            padding: 25px;
            margin: 20px 0;
            text-align: center;
        }
        .gift-box h2 {
            margin: 0 0 15px 0;
            font-size: 20px;
            color: #d63031;
        }
        .gift-item {
            background-color: rgba(255, 255, 255, 0.9);
            border-radius: 8px;
            padding: 12px;
            margin: 8px 0;
            font-weight: 500;
            color: #2d3436;
        }
        .cta-section {
            background-color: #f8f9fa;
            border-radius: 12px;
            padding: 25px;
            margin: 25px 0;
            text-align: center;
            border: 2px solid #667eea;
        }
        .cta-section h3 {
            margin: 0 0 10px 0;
            font-size: 18px;
            color: #333333;
        }
        .cta-section .time {
            color: #6c757d;
            font-size: 14px;
            margin-bottom: 20px;
        }
        .cta-button {
            display: inline-block;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: #ffffff !important;
            text-decoration: none;
            padding: 16px 40px;
            border-radius: 50px;
            font-weight: 600;
            font-size: 16px;
            box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);
            transition: transform 0.2s;
        }
        .cta-button:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(102, 126, 234, 0.5);
        }
        .benefits {
            background-color: #f0f7ff;
            border-radius: 12px;
            padding: 25px;
            margin: 25px 0;
        }
        .benefits h3 {
            margin: 0 0 20px 0;
            font-size: 18px;
            color: #333333;
            text-align: center;
        }
        .benefit-item {
            display: flex;
            align-items: flex-start;
            margin: 15px 0;
            padding: 15px;
            background-color: #ffffff;
            border-radius: 8px;
            border-left: 4px solid #667eea;
        }
        .benefit-number {
            background-color: #667eea;
            color: #ffffff;
            width: 28px;
            height: 28px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: 600;
            flex-shrink: 0;
            margin-right: 15px;
        }
        .benefit-text {
            flex: 1;
            color: #2d3436;
        }
        .benefit-text strong {
            display: block;
            color: #333333;
            margin-bottom: 3px;
        }
        .warning-box {
            background-color: #fff3cd;
            border-left: 4px solid #ffc107;
            border-radius: 8px;
            padding: 15px 20px;
            margin: 20px 0;
        }
        .warning-box strong {
            color: #856404;
            display: block;
            margin-bottom: 8px;
            font-size: 16px;
        }
        .warning-box p {
            margin: 0;
            color: #856404;
            font-size: 14px;
        }
        .info-box {
            background-color: #e7f3ff;
            border-radius: 8px;
            padding: 20px;
            margin: 20px 0;
        }
        .info-box h4 {
            margin: 0 0 12px 0;
            font-size: 16px;
            color: #0056b3;
        }
        .info-box ul {
            margin: 0;
            padding-left: 20px;
        }
        .info-box li {
            margin: 8px 0;
            color: #333333;
        }
        .footer {
            background-color: #2d3436;
            color: #ffffff;
            padding: 30px 20px;
            text-align: center;
        }
        .footer h4 {
            margin: 0 0 10px 0;
            font-size: 16px;
            font-weight: 600;
        }
        .footer p {
            margin: 5px 0;
            font-size: 14px;
            opacity: 0.9;
        }
        .footer a {
            color: #74b9ff;
            text-decoration: none;
        }
        .divider {
            height: 1px;
            background: linear-gradient(to right, transparent, #dfe6e9, transparent);
            margin: 25px 0;
        }
        @media only screen and (max-width: 600px) {
            .container {
                width: 100% !important;
            }
            .header h1 {
                font-size: 18px;
            }
            .gift-box h2 {
                font-size: 18px;
            }
            .cta-button {
                padding: 14px 30px;
                font-size: 15px;
            }
            .benefit-item {
                flex-direction: column;
            }
            .benefit-number {
                margin-bottom: 10px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <!-- Header -->
        <div class="header">
            <h1>🏥 順天堂大学スポートロジーセンター</h1>
            <p>Well-Bone Health Study</p>
        </div>

        <!-- Greeting -->
        <div class="greeting">
            ${displayName}
        </div>

        <div class="content">
            <p style="margin: 0 0 20px 0; color: #555555;">
                お申込みありがとうございます。
            </p>

            <!-- Gift Box -->
            <div class="gift-box">
                <h2>🎁 あと1ステップで受け取れます</h2>
                <div class="gift-item">
                    ✅ 無料の骨密度AI検査
                </div>
                <div class="gift-item">
                    ✅ 順天堂監修・あなた専用ウェルネスレポート
                </div>
            </div>

            <!-- CTA Section -->
            <div class="cta-section">
                <h3>📋 まず、こちらにご回答ください</h3>
                <p style="font-weight: 600; color: #333333; margin: 10px 0;">【必須】女性ウェルネス診断アンケート</p>
                <p class="time">⏱️ 所要時間：約12分</p>
                <a href="https://forms.gle/7PagAPxh6ajKUxo19" class="cta-button">
                    📝 アンケートに回答する
                </a>
            </div>

            <!-- Benefits -->
            <div class="benefits">
                <h3>📦 アンケート回答後にお届けする内容</h3>
                
                <div class="benefit-item">
                    <div class="benefit-number">1</div>
                    <div class="benefit-text">
                        <strong>骨密度AI検査</strong>
                        胸部レントゲン画像を使用
                    </div>
                </div>

                <div class="benefit-item">
                    <div class="benefit-number">2</div>
                    <div class="benefit-text">
                        <strong>順天堂監修 特別ウェルネスレポート</strong>
                        あなた専用の健康レポート
                    </div>
                </div>

                <div class="benefit-item">
                    <div class="benefit-number">3</div>
                    <div class="benefit-text">
                        <strong>動画付き特別メルマガ（全8回）</strong>
                        健康管理に役立つ情報をお届け
                    </div>
                </div>
            </div>

            <!-- Warning Box -->
            <div class="warning-box">
                <strong>⚠️ 重要なお知らせ</strong>
                <p>アンケート未完了の場合、検査・レポート・配信はいずれも行われませんのでご注意ください。</p>
            </div>

            <div class="divider"></div>

            <!-- Info Box -->
            <div class="info-box">
                <h4>📧 メールが届かない場合</h4>
                <ul>
                    <li>迷惑メールフォルダをご確認ください</li>
                    <li><strong>well-bone@juntendo.ac.jp</strong> からの受信設定をご確認ください</li>
                </ul>
            </div>

            <p style="color: #6c757d; font-size: 14px; text-align: center; margin-top: 30px;">
                ご不明点は本メールへの返信、または下記までご連絡ください。
            </p>
        </div>

        <!-- Footer -->
        <div class="footer">
            <h4>💬 お問い合わせ</h4>
            <p>順天堂大学スポートロジーセンター</p>
            <p>Well-Bone Health Study 事務局</p>
            <p>📩 <a href="mailto:well-bone@juntendo.ac.jp">well-bone@juntendo.ac.jp</a></p>
        </div>
    </div>
</body>
</html>
  `;
}

function getFirstAnswer_(value) {
  if (!value) return "";
  if (Array.isArray(value)) {
    return value[0] ? String(value[0]).trim() : "";
  }
  return String(value).trim();
}

function sendFirstNewsletterNow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート1");
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  // ==================================================
  // 【設定エリア】
  // ==================================================
  const SUBJECT = '【第1回】AI骨密度の見方 ＆ “なんとなく不調”のヒント（動画・プレゼントあり）';
  const PANOPTO_URL = 'https://juntendo.ap.panopto.com/Panopto/Pages/Viewer.aspx?id=c68e47d2-9eaa-4508-bd77-b3c9006fde42';
  const FORM_URL = 'https://forms.gle/7PagAPxh6ajKUxo19';
  const IMG_URL_1 = 'https://lh3.googleusercontent.com/d/1cvTOCxbLBWbO4pgPqT9x9NeGM3RqYHEr';
  const IMG_URL_2 = 'https://lh3.googleusercontent.com/d/1oN6Gip01Ckx1jnoqNvCRS4iYJoT-r7vp';
  // ==================================================

  // A〜Fまで読む（E=PDF送信日、F=第1回済）
  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  data.forEach((row, index) => {
    const name = row[0];        // A列: 名前
    const email = row[3];       // D列: メールアドレス
    const sentDate = row[4];    // E列: PDF送信日
    const alreadyDone = row[5]; // F列: 第1回メルマガ済みフラグ
    const rowNumber = index + 2;

    // 送信条件：E（PDF送信日）が入っている & Fが未送信
    if (sentDate !== "" && alreadyDone !== "済") {

      const htmlBody = `
        <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f5f7fb; margin: 0; padding: 0; width: 100%;">
          <tr>
            <td align="center" style="padding: 24px 12px;">
              <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="max-width: 600px; width: 100%; background-color: #ffffff; border: 1px solid #e6e9ef; border-radius: 12px;">
                <tr>
                  <td style="padding: 24px 22px; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; color: #2b2f33; line-height: 1.6; font-size: 16px;">
                    <p style="margin: 0 0 12px 0;"><strong>${name} 様</strong></p>
                    <h2 style="margin: 0 0 16px 0; color: #1f2d3d; font-size: 20px; line-height: 1.4; text-align: center; padding-bottom: 12px; border-bottom: 2px solid #2b7de9;">
                      【第1回】AI骨密度の見方 ＆ “なんとなく不調”のヒント<br>
                      <span style="font-size: 14px; font-weight: normal; color: #5f6b76;">（動画・プレゼントあり）</span>
                    </h2>

                    <p style="margin: 0 0 16px 0; font-size: 15px; color: #3d4650;">
                      今回の内容は「AI骨密度の見方」動画と、生活習慣セルフチェックです。<br>
                      やることは<span style="font-weight: bold;">動画を見る → チェックを完了</span>の2ステップだけ。
                    </p>

                    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="margin: 12px 0 20px 0;">
                      <tr><td align="center">
                        <img src="${IMG_URL_1}" alt="AI骨密度の見方イメージ" width="556" style="display: block; width: 100%; max-width: 556px; height: auto; border: 0; border-radius: 10px;">
                      </td></tr>
                    </table>

                    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="margin: 0 0 20px 0;">
                      <tr><td align="center">
                        <img src="${IMG_URL_2}" alt="女性の不調サインイメージ" width="556" style="display: block; width: 100%; max-width: 556px; height: auto; border: 0; border-radius: 10px;">
                      </td></tr>
                    </table>

                    <p style="margin: 0 0 12px 0;">AI骨密度検査の結果、もう見ましたか？<br>
                    <strong>実は20〜30代でも“やせ”がきっかけで骨が弱くなる人が続出中。</strong></p>

                    <p style="margin: 0 0 12px 0;">また最近、働く女性の間で<br>
                    ■ 疲れやすい<br>
                    ■ 眠れない<br>
                    ■ 月経の乱れ・PMS<br>
                    ■ イライラ<br>
                    などの“なんとなく不調”が増えています。</p>

                    <p style="margin: 0 0 20px 0;">実はそれ、<strong style="text-decoration: underline; background: linear-gradient(transparent 60%, #fff2a6 0%);">やせ傾向</strong>や<strong style="text-decoration: underline; background: linear-gradient(transparent 60%, #fff2a6 0%);">”ちょこちょこダイエット”</strong>が原因かも。<br>
                    このメルマガでは、8回にわたり<br>
                    <strong>骨の健康＋女性の不調をまるっと改善するヒント</strong>を<br>
                    分かりやすくお届けします。</p>

                    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border: 1px solid #cfe0f6; background-color: #f5f9ff; border-radius: 12px; margin-bottom: 18px;">
                      <tr><td style="padding: 18px;">
                        <p style="margin: 0 0 8px 0; font-weight: bold; font-size: 16px; color: #1b5bbf;">動画（最優先）</p>
                        <p style="margin: 0 0 14px 0; font-size: 15px; color: #36414b;">
                          【3分で分かる】AI骨密度の見方 ＆ 女性の不調サイン
                        </p>
                        <table role="presentation" cellpadding="0" cellspacing="0" align="center">
                          <tr><td align="center" bgcolor="#2b7de9" style="border-radius: 24px;">
                            <a href="${PANOPTO_URL}" style="display: inline-block; padding: 12px 28px; font-size: 16px; font-weight: bold; color: #ffffff; text-decoration: none; border-radius: 24px; line-height: 20px; min-width: 200px;">
                              動画を見る
                            </a>
                          </td></tr>
                        </table>
                        <p style="margin: 12px 0 0 0; font-size: 13px; color: #5a6672;">あなたのAI結果をどう理解すればいいか、ここで一気にわかります。</p>
                      </td></tr>
                    </table>

                    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border: 1px solid #f0d5b2; background-color: #fff7ed; border-radius: 12px; margin-bottom: 18px;">
                      <tr><td style="padding: 18px;">
                        <p style="margin: 0 0 8px 0; font-weight: bold; font-size: 16px; color: #a55a00;">生活習慣セルフチェック</p>
                        <p style="margin: 0 0 14px 0; font-size: 15px; color: #36414b;">すべて回答して、あなた専用の「ウェルネス解析レポート」をgetしよう</p>
                        <table role="presentation" cellpadding="0" cellspacing="0" align="center">
                          <tr><td align="center" bgcolor="#f39c12" style="border-radius: 24px;">
                            <a href="${FORM_URL}" style="display: inline-block; padding: 12px 28px; font-size: 16px; font-weight: bold; color: #ffffff; text-decoration: none; border-radius: 24px; line-height: 20px; min-width: 220px;">
                              セルフチェックを開始する
                            </a>
                          </td></tr>
                        </table>
                      </td></tr>
                    </table>

                    <p style="margin: 0 0 16px 0; padding-top: 16px; border-top: 1px dashed #d7dde4;">
                      次回は、<br>
                      <strong>骨が弱い人・体調がわるい人に共通する、ある"習慣"とは？</strong><br>
                      多くの方に当てはまる、意外な事実をご紹介します。
                    </p>
                    <p style="margin: 0;">どうぞお楽しみに！</p>
                  </td>
                </tr>

                <tr>
                  <td style="padding: 0 22px 22px; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; color: #7a8794; font-size: 12px; text-align: center;">
                    ※本メールはウェルネス検査を受診された方にお送りしています。
                  </td>
                </tr>

              </table>
            </td>
          </tr>
        </table>
      `;

      try {
        GmailApp.sendEmail(email, SUBJECT, "HTMLメールを表示できる環境でご覧ください。", {
          htmlBody,
          name: "ウェルネス事務局",
        });

        // 第1回「済」→ F列（6列目）
        sheet.getRange(rowNumber, 6).setValue("済");
        console.log(`[行${rowNumber}] 第1回 送信完了: ${name}様 / ${email}`);

      } catch (e) {
        console.error(`[行${rowNumber}] 第1回 エラー: ${name}様 / ${email} - ${e.message}`);
      }
    }
  });
}

/**
 * 第2回メルマガ（第1回デザイン踏襲）
 * 第2回の送信済みフラグは G列（7列目）推奨
 */
function sendSecondNewsletterNow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート1");
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const SUBJECT = '【第2回】あなたもFUSかも？〜“やせ”と不調の関係〜';
  const PANOPTO_URL = 'https://juntendo.ap.panopto.com/Panopto/Pages/Viewer.aspx?id=c68e47d2-9eaa-4508-bd77-b3c9006fde42';
  const FORM_URL = 'https://forms.gle/7PagAPxh6ajKUxo19';
  const IMG_URL_1 = 'https://lh3.googleusercontent.com/d/1cvTOCxbLBWbO4pgPqT9x9NeGM3RqYHEr';
  const IMG_URL_2 = 'https://lh3.googleusercontent.com/d/1oN6Gip01Ckx1jnoqNvCRS4iYJoT-r7vp';

  // A〜Gまで読む（E=PDF送信日、G=第2回済）
  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

  data.forEach((row, index) => {
    const name = row[0];        // A
    const email = row[3];       // D
    const sentDate = row[4];    // E（PDF送信日）
    const alreadyDone2 = row[6];// G（第2回済）
    const rowNumber = index + 2;

    if (sentDate !== "" && alreadyDone2 !== "済") {

      const htmlBody = `
        <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f5f7fb; margin: 0; padding: 0; width: 100%;">
          <tr>
            <td align="center" style="padding: 24px 12px;">
              <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="max-width: 600px; width: 100%; background-color: #ffffff; border: 1px solid #e6e9ef; border-radius: 12px;">
                <tr>
                  <td style="padding: 24px 22px; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; color: #2b2f33; line-height: 1.6; font-size: 16px;">
                    <p style="margin: 0 0 12px 0;"><strong>${name} 様</strong></p>
                    <h2 style="margin: 0 0 16px 0; color: #1f2d3d; font-size: 20px; line-height: 1.4; text-align: center; padding-bottom: 12px; border-bottom: 2px solid #2b7de9;">
                      【第2回】あなたもFUSかも？<br>
                      <span style="font-size: 14px; font-weight: normal; color: #5f6b76;">〜“やせ”と不調の関係〜</span>
                    </h2>

                    <p style="margin: 0 0 8px 0; font-size: 15px; color: #3d4650;">
                      今回やること：<strong>①動画 → ②セルフチェック</strong>
                    </p>

                    <table role="presentation" cellpadding="0" cellspacing="0" align="center" style="margin: 0 0 18px 0;">
                      <tr><td align="center" bgcolor="#2b7de9" style="border-radius: 24px;">
                        <a href="${PANOPTO_URL}" style="display: inline-block; padding: 16px 32px; font-size: 16px; font-weight: bold; color: #ffffff; text-decoration: none; border-radius: 24px; line-height: 20px; min-width: 200px;">
                          第1回動画を見る
                        </a>
                      </td></tr>
                    </table>

                    <p style="margin: 0 0 16px 0; font-size: 15px; color: #3d4650;">
                      疲れやすい、眠れない、イライラなどの“なんとなく不調”。<br>
                      原因として最近注目されているのが、<strong>FUS（女性の低体重・低栄養症候群）</strong>です！
                    </p>

                    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="margin: 12px 0 20px 0;">
                      <tr><td align="center">
                        <img src="${IMG_URL_1}" alt="FUSのイメージ" width="556" style="display: block; width: 100%; max-width: 556px; height: auto; border: 0; border-radius: 10px;">
                      </td></tr>
                    </table>

                    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="margin: 0 0 20px 0;">
                      <tr><td align="center">
                        <img src="${IMG_URL_2}" alt="女性の不調サインイメージ" width="556" style="display: block; width: 100%; max-width: 556px; height: auto; border: 0; border-radius: 10px;">
                      </td></tr>
                    </table>

                    <p style="margin: 0 0 12px 0;"><strong>■ FUSとは？</strong><br>
                      食事が偏ることで、痩せていなくても低栄養になり、<br>
                      疲れやすい・眠れない・イライラ・月経の乱れ<br>
                      といった不調や、骨の老化につながりうる状態のことです。
                    </p>

                    <p style="margin: 0 0 12px 0;"><strong>こんなあなたは、FUS要注意！</strong><br>
                      ・「やせたい」から、ちょこちょこダイエット中（サラダだけ／炭水化物少なめ／油控えめ…）<br>
                      ・ダイエットのつもりはないが、実際に食べる量が少なめ（忙しくて軽食・朝食抜きがち…）
                    </p>

                    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border: 1px solid #cfe0f6; background-color: #f5f9ff; border-radius: 12px; margin-bottom: 18px;">
                      <tr><td style="padding: 18px;">
                        <p style="margin: 0 0 8px 0; font-weight: bold; font-size: 16px; color: #1b5bbf;">▶︎【再送】第1回動画はこちら</p>
                        <p style="margin: 0 0 14px 0; font-size: 15px; color: #36414b;">
                          「AI骨密度の見方 &amp; 女性の不調サイン」<br>
                          短い動画なので、まだの方はこの機会にぜひ。
                        </p>
                        <table role="presentation" cellpadding="0" cellspacing="0" align="center" style="margin: 0 0 8px 0;">
                          <tr><td align="center" bgcolor="#2b7de9" style="border-radius: 24px;">
                            <a href="${PANOPTO_URL}" style="display: inline-block; padding: 16px 32px; font-size: 16px; font-weight: bold; color: #ffffff; text-decoration: none; border-radius: 24px; line-height: 20px; min-width: 200px;">
                              動画を見る
                            </a>
                          </td></tr>
                        </table>
                      </td></tr>
                    </table>

                    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border: 1px solid #f0d5b2; background-color: #fff7ed; border-radius: 12px; margin-bottom: 18px;">
                      <tr><td style="padding: 18px;">
                        <p style="margin: 0 0 8px 0; font-weight: bold; font-size: 16px; color: #a55a00;">📝【再案内】事前アンケート</p>
                        <p style="margin: 0 0 14px 0; font-size: 15px; color: #36414b;">回答すると、あなた専用のウェルネス解析レポートをお届けします🎁</p>
                        <table role="presentation" cellpadding="0" cellspacing="0" align="center" style="margin: 0 0 8px 0;">
                          <tr><td align="center" bgcolor="#f39c12" style="border-radius: 24px;">
                            <a href="${FORM_URL}" style="display: inline-block; padding: 16px 32px; font-size: 16px; font-weight: bold; color: #ffffff; text-decoration: none; border-radius: 24px; line-height: 20px; min-width: 220px;">
                              セルフチェックを開始する
                            </a>
                          </td></tr>
                        </table>
                      </td></tr>
                    </table>

                    <p style="margin: 0 0 16px 0; padding-top: 16px; border-top: 1px dashed #d7dde4;">
                      次回は、今日からできる<strong>“食事改善のポイント”</strong>をご紹介します。<br>
                      どうぞお楽しみに！
                    </p>
                  </td>
                </tr>

                <tr>
                  <td style="padding: 0 22px 22px; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; color: #7a8794; font-size: 12px; text-align: center;">
                    ※本メールはウェルネス検査を受診された方にお送りしています。
                  </td>
                </tr>

              </table>
            </td>
          </tr>
        </table>
      `;

      try {
        GmailApp.sendEmail(email, SUBJECT, "HTMLメールを表示できる環境でご覧ください。", {
          htmlBody,
          name: "ウェルネス事務局",
        });

        // 第2回「済」→ G列（7列目）
        sheet.getRange(rowNumber, 7).setValue("済");
        console.log(`[行${rowNumber}] 第2回 送信完了: ${name}様 / ${email}`);

      } catch (e) {
        console.error(`[行${rowNumber}] 第2回 エラー: ${name}様 / ${email} - ${e.message}`);
      }
    }
  });
}
