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

      

      const htmlBody = `

        <div style="font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; color: #333; line-height: 1.8; max-width: 600px; margin: auto; border: 1px solid #eee; padding: 25px;">

          

          <p style="font-size: 16px;"><strong>${name} 様</strong></p>

          

          <h2 style="color: #2c3e50; font-size: 19px; border-bottom: 2px solid #3498db; padding-bottom: 10px; text-align: center;">

            【第1回】AI骨密度の見方 ＆ “なんとなく不調”のヒント<br>

            <span style="font-size: 14px; font-weight: normal;">（動画・プレゼントあり）</span>

          </h2>

          

          <div style="text-align: center; margin: 20px 0;">

            <img src="${IMG_URL_1}" style="width: 100%; max-width: 500px;" alt="Image1">

          </div>



          <div style="text-align: center; margin: 20px 0;">

            <img src="${IMG_URL_2}" style="width: 100%; max-width: 500px;" alt="Image2">

          </div>



          <p>AI骨密度検査の結果、もう見ましたか？<br>

          <strong>実は20〜30代でも“やせ”がきっかけで骨が弱くなる人が続出中。</strong></p>

          

          <p>また最近、働く女性の間で<br>

          ■ 疲れやすい<br>

          ■ 眠れない<br>

          ■ 月経の乱れ・PMS<br>

          ■ イライラ<br>

          などの“なんとなく不調”が増えています。</p>

          

          <p>実はそれ、<strong style="text-decoration: underline; background: linear-gradient(transparent 60%, #ffff66 0%);">やせ傾向</strong>や<strong style="text-decoration: underline; background: linear-gradient(transparent 60%, #ffff66 0%);">”ちょこちょこダイエット”</strong>が原因かも。<br>

          このメルマガでは、8回にわたり<br>

          <strong>骨の健康＋女性の不調をまるっと改善するヒント</strong>を<br>

          分かりやすくお届けします。</p>



          <div style="background-color: #f0f7ff; padding: 20px; border-radius: 10px; border: 1px dashed #3498db; margin: 25px 0; text-align: center;">

            <p style="margin-top: 0; font-weight: bold; font-size: 17px; color: #2980b9;">[ ▶ まずはこの動画から！ ]</p>

            <a href="${PANOPTO_URL}" style="color: #c0392b; text-decoration: underline; font-size: 18px; font-weight: bold;">

              【3分で分かる】AI骨密度の見方 ＆ 女性の不調サイン

            </a>

            <p style="font-size: 14px; color: #666; margin-top: 10px;">あなたのAI結果をどう理解すればいいか、ここで一気にわかります。</p>

          </div>



          <div style="background-color: #fff9f0; padding: 20px; border-radius: 10px; border: 1px solid #f39c12; margin: 25px 0; text-align: center;">

            <p style="margin-top: 0; font-weight: bold; font-size: 16px;">

              生活習慣セルフチェック！<br>


            </p>

            <div style="text-align: center; margin-top: 15px;">

              <a href="${FORM_URL}" style="display: inline-block; background-color: #f39c12; color: white; padding: 12px 30px; text-decoration: none; border-radius: 50px; font-weight: bold; box-shadow: 0 4px 0 #d35400;">

                セルフチェックを開始する

              </a>
                <p style="font-size: 14px; color: #666; margin-top: 10px;">すべて回答して、あなた専用の「ウェルネス解析レポート」をgetしよう</p>

            </div>

          </div>



          <p style="border-top: 2px dotted #eee; padding-top: 20px;">

            次回は、<br>

            <strong>骨が弱い人・体調がわるい人に共通する、ある"習慣"とは？</strong><br>

            多くの方に当てはまる、意外な事実をご紹介します。

          </p>

          

          <p>どうぞお楽しみに！</p>



          <p style="font-size: 12px; color: #999; text-align: center; margin-top: 30px;">

            ※本メールはウェルネス検査を受診された方にお送りしています。

          </p>

        </div>

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
