const json = (statusCode, body) => new Response(JSON.stringify(body), { status: statusCode, headers: { 'content-type': 'application/json' } });

export default async request => {
  if ((request.method || request.httpMethod) !== 'POST') return json(405, { error: 'POST required.' });
  const provider = process.env.OCR_PROVIDER || 'none';
  if (provider !== 'openai' || !process.env.OPENAI_API_KEY) return json(503, { error: 'Server OCR is not configured. Add OCR_PROVIDER=openai and OPENAI_API_KEY in Netlify environment variables.' });
  try {
    const body = JSON.parse(request.body || '{}');
    const response = await fetch('https://api.openai.com/v1/responses', { method: 'POST', headers: { authorization: `Bearer ${process.env.OPENAI_API_KEY}`, 'content-type': 'application/json' }, body: JSON.stringify({ model: process.env.OCR_MODEL || 'gpt-4.1-mini', input: [{ role: 'user', content: [{ type: 'input_text', text: 'Read this receipt carefully. Return JSON only. Extract PURCHASED LINE ITEMS, not tax, subtotal, total, discounts, payment, card, loyalty, barcode, invoice number, table number, or phone number. Receipt rows commonly look like quantity, item description, unit price, currency, line total. Preserve each item as a separate object. For each item: quantity is the leading quantity if visible, description is only the product/service name, unitPrice is the per-item price, lineTotal is quantity multiplied by unitPrice when visible or the row total. Do not confuse dates, tax rates, receipt numbers, or totals with item prices. Detect the receipt date and currency once and repeat them on each item. If a value is uncertain, return null instead of inventing it.' }, { type: 'input_image', image_url: `data:${body.mimeType || 'image/jpeg'};base64,${body.imageBase64}` }] }], text: { format: { type: 'json_schema', name: 'receipt', strict: true, schema: { type: 'object', properties: { items: { type: 'array', items: { type: 'object', properties: { quantity: { type: ['number', 'null'] }, description: { type: 'string' }, unitPrice: { type: ['number', 'null'] }, lineTotal: { type: ['number', 'null'] }, currency: { type: ['string', 'null'] }, date: { type: ['string', 'null'] }, category: { type: ['string', 'null'] } }, required: ['quantity', 'description', 'unitPrice', 'lineTotal', 'currency', 'date', 'category'], additionalProperties: false } } }, required: ['items'], additionalProperties: false } } } }) });
    if (!response.ok) return json(502, { error: 'Receipt AI provider failed.' });
    const result = await response.json();
    return json(200, JSON.parse(result.output_text));
  } catch (error) { return json(500, { error: 'Receipt processing failed.' }); }
}
