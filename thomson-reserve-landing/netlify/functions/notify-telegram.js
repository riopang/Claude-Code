exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method Not Allowed' };
  }

  const token = process.env.TELEGRAM_BOT_TOKEN;
  const chatId = process.env.TELEGRAM_CHAT_ID;

  if (!token || !chatId) {
    return { statusCode: 500, body: 'Telegram not configured' };
  }

  let data;
  try {
    data = JSON.parse(event.body);
  } catch (e) {
    return { statusCode: 400, body: 'Invalid JSON' };
  }

  const { name, phone, purpose, budget } = data;

  const text =
    `🏠 *New Lead — Thomson Reserve*\n\n` +
    `*Name:* ${name || '-'}\n` +
    `*Phone:* ${phone || '-'}\n` +
    `*Own Stay/Investment:* ${purpose || '-'}\n` +
    `*Budget:* ${budget || '-'}`;

  const res = await fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      chat_id: chatId,
      text,
      parse_mode: 'Markdown'
    })
  });

  if (!res.ok) {
    const errText = await res.text();
    return { statusCode: 502, body: `Telegram error: ${errText}` };
  }

  return { statusCode: 200, body: JSON.stringify({ ok: true }) };
};
