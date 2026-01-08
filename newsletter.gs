/**

 * 第1回メルマガ（デザイン・文言再調整版）

 */

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



  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();



  data.forEach((row, index) => {

    const name = row[0];        // A列: 名前

    const email = row[3];       // D列: メールアドレス

    const sentDate = row[4];    // E列: PDF送信日

    const alreadyDone = row[5]; // F列: メルマガ済みフラグ

    const rowNumber = index + 2;



    if (sentDate !== "" && alreadyDone !== "済") {

      
      // プレビューは、GASエディタでこのHTMLをHtmlServiceに貼り付けてプレビューするのがおすすめです。
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
                      <tr>
                        <td align="center">
                          <img src="${IMG_URL_1}" alt="AI骨密度の見方イメージ" width="556" style="display: block; width: 100%; max-width: 556px; height: auto; border: 0; border-radius: 10px;">
                        </td>
                      </tr>
                    </table>
                    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="margin: 0 0 20px 0;">
                      <tr>
                        <td align="center">
                          <img src="${IMG_URL_2}" alt="女性の不調サインイメージ" width="556" style="display: block; width: 100%; max-width: 556px; height: auto; border: 0; border-radius: 10px;">
                        </td>
                      </tr>
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
                      <tr>
                        <td style="padding: 18px;">
                          <p style="margin: 0 0 8px 0; font-weight: bold; font-size: 16px; color: #1b5bbf;">動画（最優先）</p>
                          <p style="margin: 0 0 14px 0; font-size: 15px; color: #36414b;">
                            【3分で分かる】AI骨密度の見方 ＆ 女性の不調サイン
                          </p>
                          <table role="presentation" cellpadding="0" cellspacing="0" align="center">
                            <tr>
                              <td align="center" bgcolor="#2b7de9" style="border-radius: 24px;">
                                <a href="${PANOPTO_URL}" style="display: inline-block; padding: 12px 28px; font-size: 16px; font-weight: bold; color: #ffffff; text-decoration: none; border-radius: 24px; line-height: 20px; min-width: 200px;">
                                  動画を見る
                                </a>
                              </td>
                            </tr>
                          </table>
                          <p style="margin: 12px 0 0 0; font-size: 13px; color: #5a6672;">あなたのAI結果をどう理解すればいいか、ここで一気にわかります。</p>
                        </td>
                      </tr>
                    </table>
                    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border: 1px solid #f0d5b2; background-color: #fff7ed; border-radius: 12px; margin-bottom: 18px;">
                      <tr>
                        <td style="padding: 18px;">
                          <p style="margin: 0 0 8px 0; font-weight: bold; font-size: 16px; color: #a55a00;">
                            生活習慣セルフチェック
                          </p>
                          <p style="margin: 0 0 14px 0; font-size: 15px; color: #36414b;">すべて回答して、あなた専用の「ウェルネス解析レポート」をgetしよう</p>
                          <table role="presentation" cellpadding="0" cellspacing="0" align="center">
                            <tr>
                              <td align="center" bgcolor="#f39c12" style="border-radius: 24px;">
                                <a href="${FORM_URL}" style="display: inline-block; padding: 12px 28px; font-size: 16px; font-weight: bold; color: #ffffff; text-decoration: none; border-radius: 24px; line-height: 20px; min-width: 220px;">
                                  セルフチェックを開始する
                                </a>
                              </td>
                            </tr>
                          </table>
                        </td>
                      </tr>
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

          htmlBody: htmlBody,

          name: "ウェルネス事務局"

        });

        sheet.getRange(rowNumber, 6).setValue("済");

        console.log(`[行${rowNumber}] 送信完了: ${name}様`);

      } catch (e) {

        console.error(`[行${rowNumber}] エラー: ${name}様 - ${e.message}`);

      }

    }

  });

}

/**

 * 第2回メルマガ（第1回デザイン踏襲）

 */

function sendSecondNewsletterNow() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート1");

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return;



  // ==================================================

  // 【設定エリア】

  // ==================================================

  const SUBJECT = '【第2回】あなたもFUSかも？〜“やせ”と不調の関係〜';

  

  const PANOPTO_URL = 'https://juntendo.ap.panopto.com/Panopto/Pages/Viewer.aspx?id=c68e47d2-9eaa-4508-bd77-b3c9006fde42'; 

  const FORM_URL = 'https://forms.gle/7PagAPxh6ajKUxo19';



  const IMG_URL_1 = 'https://lh3.googleusercontent.com/d/1cvTOCxbLBWbO4pgPqT9x9NeGM3RqYHEr'; 

  const IMG_URL_2 = 'https://lh3.googleusercontent.com/d/1oN6Gip01Ckx1jnoqNvCRS4iYJoT-r7vp'; 

  // ==================================================



  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();



  data.forEach((row, index) => {

    const name = row[0];        // A列: 名前

    const email = row[3];       // D列: メールアドレス

    const sentDate = row[4];    // E列: PDF送信日

    const alreadyDone = row[6]; // G列: 第2回メルマガ済みフラグ

    const rowNumber = index + 2;



    if (sentDate !== "" && alreadyDone !== "済") {

      
      // プレビューは、GASエディタでこのHTMLをHtmlServiceに貼り付けてプレビューするのがおすすめです。
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
                      <tr>
                        <td align="center" bgcolor="#2b7de9" style="border-radius: 24px;">
                          <a href="${PANOPTO_URL}" style="display: inline-block; padding: 16px 32px; font-size: 16px; font-weight: bold; color: #ffffff; text-decoration: none; border-radius: 24px; line-height: 20px; min-width: 200px;">
                            動画を見る
                          </a>
                        </td>
                      </tr>
                    </table>
                    <p style="margin: 0 0 16px 0; font-size: 15px; color: #3d4650;">
                      疲れやすい、眠れない、イライラなどの“なんとなく不調”。<br>
                      原因として最近注目されているのが、<strong>FUS（女性の低体重・低栄養症候群）</strong>です！
                    </p>
                    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="margin: 12px 0 20px 0;">
                      <tr>
                        <td align="center">
                          <img src="${IMG_URL_1}" alt="FUSのイメージ" width="556" style="display: block; width: 100%; max-width: 556px; height: auto; border: 0; border-radius: 10px;">
                        </td>
                      </tr>
                    </table>
                    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="margin: 0 0 20px 0;">
                      <tr>
                        <td align="center">
                          <img src="${IMG_URL_2}" alt="女性の不調サインイメージ" width="556" style="display: block; width: 100%; max-width: 556px; height: auto; border: 0; border-radius: 10px;">
                        </td>
                      </tr>
                    </table>
                    <p style="margin: 0 0 12px 0;"><strong>■ FUSとは？</strong><br>
                    食事が偏ることで、痩せていなくても低栄養になり、<br>
                    疲れやすい・眠れない・イライラ・月経の乱れ<br>
                    といった不調や、骨の老化につながりうる状態のことです。</p>
                    <p style="margin: 0 0 12px 0;"><strong>こんなあなたは、FUS要注意！</strong><br>
                    ・「やせたい」から、ちょこちょこダイエット中（サラダだけ／炭水化物少なめ／油控えめ…）<br>
                    ・ダイエットのつもりはないが、実際に食べる量が少なめ（忙しくて軽食・朝食抜きがち…）</p>
                    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border: 1px solid #cfe0f6; background-color: #f5f9ff; border-radius: 12px; margin-bottom: 18px;">
                      <tr>
                        <td style="padding: 18px;">
                          <p style="margin: 0 0 8px 0; font-weight: bold; font-size: 16px; color: #1b5bbf;">▶︎【再送】第1回動画はこちら</p>
                          <p style="margin: 0 0 14px 0; font-size: 15px; color: #36414b;">
                            「AI骨密度の見方 &amp; 女性の不調サイン」<br>
                            短い動画なので、まだの方はこの機会にぜひ。
                          </p>
                          <table role="presentation" cellpadding="0" cellspacing="0" align="center" style="margin: 0 0 8px 0;">
                            <tr>
                              <td align="center" bgcolor="#2b7de9" style="border-radius: 24px;">
                                <a href="${PANOPTO_URL}" style="display: inline-block; padding: 16px 32px; font-size: 16px; font-weight: bold; color: #ffffff; text-decoration: none; border-radius: 24px; line-height: 20px; min-width: 200px;">
                                  動画を見る
                                </a>
                              </td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                    </table>
                    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border: 1px solid #f0d5b2; background-color: #fff7ed; border-radius: 12px; margin-bottom: 18px;">
                      <tr>
                        <td style="padding: 18px;">
                          <p style="margin: 0 0 8px 0; font-weight: bold; font-size: 16px; color: #a55a00;">
                            📝【再案内】事前アンケート
                          </p>
                          <p style="margin: 0 0 14px 0; font-size: 15px; color: #36414b;">回答すると、あなた専用のウェルネス解析レポートをお届けします🎁</p>
                          <table role="presentation" cellpadding="0" cellspacing="0" align="center" style="margin: 0 0 8px 0;">
                            <tr>
                              <td align="center" bgcolor="#f39c12" style="border-radius: 24px;">
                                <a href="${FORM_URL}" style="display: inline-block; padding: 16px 32px; font-size: 16px; font-weight: bold; color: #ffffff; text-decoration: none; border-radius: 24px; line-height: 20px; min-width: 220px;">
                                  セルフチェックを開始する
                                </a>
                              </td>
                            </tr>
                          </table>
                        </td>
                      </tr>
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

          htmlBody: htmlBody,

          name: "ウェルネス事務局"

        });

        sheet.getRange(rowNumber, 7).setValue("済");

        console.log(`[行${rowNumber}] 送信完了: ${name}様`);

      } catch (e) {

        console.error(`[行${rowNumber}] エラー: ${name}様 - ${e.message}`);

      }

    }

  });

}
