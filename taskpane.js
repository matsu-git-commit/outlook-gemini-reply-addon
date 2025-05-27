const GEMINI_API_KEY = "YOUR_GEMINI_API_KEY"; // AIzaSyBfsxQbBMpbAcGhgGyxVpwnss-6qVLrZM0

document.addEventListener("DOMContentLoaded", function() {
  document.getElementById('generate').onclick = async function() {
    const policy = document.getElementById('policy').value.trim();
    if (!policy) {
      alert('方針を入力してください');
      return;
    }
    document.getElementById('result').innerText = 'AI生成中...';

    // Gemini API呼び出し
    const prompt = `以下の方針に従って、ビジネスメールの返信文を日本語で丁寧に作成してください。\n方針: ${policy}`;
    try {
      const res = await fetch(
        'https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=' + GEMINI_API_KEY,
        {
          method: 'POST',
          headers: {'Content-Type': 'application/json'},
          body: JSON.stringify({
            contents: [{parts: [{text: prompt}]}]
          })
        }
      );
      const data = await res.json();
      const reply = data.candidates?.[0]?.content?.parts?.[0]?.text || '生成失敗';
      document.getElementById('result').innerText = reply;

      // Office.jsで本文に挿入
      if (Office.context.mailbox.item.body) {
        Office.context.mailbox.item.body.setSelectedDataAsync(
          reply,
          {coercionType: Office.CoercionType.Text},
          function(asyncResult) {
            if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
              alert('本文への挿入に失敗しました: ' + asyncResult.error.message);
            }
          }
        );
      }
    } catch (e) {
      document.getElementById('result').innerText = 'エラー: ' + e.message;
    }
  };
}); 
