/**
 * Receipt OCR - one photo to Gemini and back. The prompt and the answer's shape were tried on
 * mom's first receipts (handwritten slips, a return slip, a notebook page): without the list of
 * product names the handwriting came back as nonsense; with it most lines matched a real name.
 */

// '"Төв", "Тос" = Тосгон, "Ачит", "Myhome" = My home' (OCR_PLACE_ALIASES, ocr-config.gs).
const ocrAliasText = () => Object.entries(OCR_PLACE_ALIASES).map(([place, names]) => `${names.map((name) => `"${name}"`).join(', ')} = ${place}`).join(', ');

const OCR_PROMPT_LINES = [
  'Энэ бол Монголын дэлгүүрийн цаасан баримтын зураг (гараар бичсэн хэсэгтэй байж болно). Зураг дээр хэд хэдэн баримт харагдаж болно: зөвхөн гол (хамгийн том, төвд байгаа) баримтыг унш, булан дахь бусад баримтыг "other_documents_visible"-д тэмдэглэ.',
  'Дүрэм:',
  '- Тааж бүү бич. Уншигдахгүй эсвэл эргэлзээтэй бол утгыг null болгож, тэр мөрийн confidence-ийг бага тавь. Ганц үл хамаарах зүйл нь тоо (qty), доорхыг үз.',
  '- "qty" (тоо ширхэг)-ийг мөр бүрт ЗААВАЛ бич, хэзээ ч null биш. "Тоо" баганад "ш", "1ш", ✓, зураас, гарын тэмдэг байвал 1. Гараар бичсэн "1ш" нь "1ш", "1u", "lu", "Lu", "tu", "1w" мэт харагдаж болно: энэ бүгд 1 (ширхэг). "2ш" = 2, "3ш" = 3: "ш"-ийн өмнөх тоо. Тоо бичигдсэн бол тэр тоо. Огт уншигдахгүй бол хамгийн боломжит тоог (ихэвчлэн 1) бичээд тэр мөрийн confidence-ийг 0.5-аас бага тавь.',
  '- Тоог яг зураг дээрх шиг бич, өөрөө тооцоолж засахгүй. Нийлбэр зөрвөл issues-д бич.',
  '- "date": баримт дээр бичсэн огноо, "YYYY-MM-DD" хэлбэрээр: "20__ он 10 сар 18" бол 2026-10-18 (он бүтэн бичээгүй бол одоогийн он гэж үз). Огноо байхгүй эсвэл уншигдахгүй бол null, бүү таа.',
  '- document_title: баримтын гарчиг яг зураг дээрх шиг ("БАРАА МАТЕРИАЛЫН БУЦААЛТ", "ЗАРЛАГЫН ПАДААН", ...), байхгүй бол null.',
  '- header_receiver / header_sender / shop_line: толгойд "Хүлээн авагч" (эсвэл "Хаана", "→"-ийн дараах газар), "Илгээгч" (эсвэл "Хаанаас", "→"-ээс өмнөх газар, жишээ нь "Тос → Соронзонд" бол илгээгч "Тос"), "Дэлгүүр:" гэж бичсэн утгууд, байхгүй бол null. "ЗАРЛАГЫН БАРИМТ" маягт дээр "Байгууллагын нэр" нь илгээгч (header_sender), "Худалдан авагчийн нэр" нь хүлээн авагч (header_receiver).',
  '- Тосгон (агуулах) Соронзон эсвэл My home дэлгүүр рүү бараа гаргахдаа "ЗАРЛАГЫН БАРИМТ" бичдэг: энэ нэг баримт нь Тосгоны зарлага, тэр дэлгүүрийн орлого хоёулаа. Зураг дэлгүүрийн орлогын хавтсанд байсан ч илгээгчийг header_sender-д ("Тосгон"), хүлээн авагч дэлгүүрийг header_receiver-д ЗААВАЛ бич (систем үүгээр хоёр талд нь бичнэ); document_type нь "zarlaga" хэвээр.',
  '- "Байгууллагын нэр" (илгээгч) дээр "Хотоос", "Хот" гэж бичсэн бол бараа хотоос Тосгонд ирээд, Тосгоноос тэр дэлгүүрт очсон: энэ нэг баримт нь Тосгоны орлого, Тосгоны зарлага, тэр дэлгүүрийн орлого гурвуулаа. header_sender-д "Хот" (Тосгон биш), хүлээн авагч дэлгүүрийг header_receiver-д ЗААВАЛ бич (систем үүгээр гурван газарт нь бичнэ); document_type нь "zarlaga" хэвээр.',
  `- Бичсэн хүн газруудыг өөрийнхөөрөө нэрлэдэг: ${ocrAliasText()}. Эдгээрийн аль нэгийг уншвал header_sender / header_receiver-д тэр газрын жинхэнэ нэрийг бич ("Тосгон", "My home", "Соронзон"); бусад нэрийг зураг дээрх шиг.`,
  '- document_type: "orlogo" (нийлүүлэгчээс орсон), "zarlaga" (борлуулалт/зарлага), "shiljuuleg" (дэлгүүр/агуулах хоорондын шилжүүлэг), "butsaalt", "toollogo" (тооллогын хуудас: бараа тус бүрийн тоолж гаргасан үлдэгдэл), "other" (маягтгүй дэвтрийн хуудас, чөлөөт тэмдэглэл), "blank" (хоосон/цуцалсан).',
  '- Хэвлэсэн маягттай борлуулалтын баримт (жишээ нь "JNN Network" маягт: "Дэлгүүр: Соронзон орон", "Мөнгө хүлээн авагч", "Бараа материалын нэр / Нэг бүрийн үнэ / Тоо / Нийт үнэ" баганатай) бол "zarlaga". Маягт дээр "Илгээгч", "Хүлээн авагч" гэсэн мөр хэвлэгдсэн байсан ч энэ нь шилжүүлэг биш. "shiljuuleg" / "butsaalt" зөвхөн гарчиг нь шилжүүлэг/буцаалт гэсэн, эсвэл "Тос → Соронзон" мэт хоёр газрын хооронд бараа хөдөлсөнийг тодорхой бичсэн баримтад.',
  '- Тооллогын хуудас бол "qty"-д тоолсон үлдэгдлийг бичнэ (0 ч байж болно).',
  '- Барааны нэрийг зураг дээрх хэлбэрээр нь "name"-д бич.',
  '- Доорх жагсаалт бол манай барааны нэрс, "|"-ийн ард одоогийн зарах үнэ (₮). Мөр бүрт "matched_name"-д жагсаалтаас ЗААВАЛ ганц нэр сонго: хамгийн магадлалтайг нь, "|"-ээс өмнөх нэрийг ЯГ тэр чигээр нь (үнийг бичихгүй). Эргэлзэж байсан ч хамгийн ойрыг нь сонго, хоосон бүү үлдээ. Жагсаалтад тохирох бараа огт байхгүй бол "Хамааралгүй".',
  '- Бичсэн хүн латин нэрийг алдаатай бичдэг, брэндийг товчилдог: Panu/Pana/Пану = Panasonic, Mei = Meiling, CMe/CMedia/Cmidea = Midea, Royszan/Ronsh = Ronshen. "8.0" = 80l, "6.0" = 60l (литр). Тоо, хэмжээ нь нэрийн чухал хэсэг: ижил брэндээс тоо нь таарахыг сонго.',
  '- Үнэ нь нэр сонгоход туслана (ижил төстэй хоёр нэрээс), бас хүснэгтэд "Баримтын үнэ" болж бичигдэнэ. Баримт дээрх үнэ МЯНГААР бичигддэг: 430 = 430 000₮, "130.0" эсвэл "130,0" = 130 000₮ (цэгийн ард 0 нь мянгат биш), "65.0" = 65 000₮, 1280 = 1 280 000₮. Шилжүүлгийн үнэ зарах үнээс өөр байж болно. "unit_price"-д БҮТЭН төгрөгөөр бич (430 гэж бичсэн бол 430000).',
  '- "unit_price" бол НЭГ ШИРХЭГИЙН үнэ: маягттай баримтад "Нэг бүрийн үнэ" (эсвэл "Үнэ") баганын тоо. "Нийт үнэ" баганын тоо, огноо, тэмдэглэлийг "unit_price"-д бүү бич. Мөрийн үнэ зураг дээр харагдаж байвал ЗААВАЛ бич, огт бичээгүй бол л null.',
  '- "15x426", "15×426", "15 х 426" хэлбэрийн бичлэг: эхний тоо (x-ийн өмнөх) нь ТОО ШИРХЭГ ("qty" = 15), x-ийн арын тоо нь НЭГЖ ҮНЭ (426). Тоог үнэтэй сольж болохгүй.',
  '- Хятадаас авсан барааны ЮАНИЙН үнэ: маягтгүй дэвтрийн хуудас (ихэвчлэн орлого), "тоо x үнэ" хэлбэрээр бичсэн, эсвэл ¥, 元, "юань", "yuan" тэмдэгтэй бол үнэ нь Хятад юань, ойролцоо ӨРТӨГ ҮНЭ. Энэ үнэ мянгаар биш: 426 = 426 юань (426 000₮ БИШ). Тэгвэл "price_currency" = "CNY", "unit_price"-д зураг дээрх тоог өөрчлөхгүй бич (426), жагсаалтын зарах үнэтэй харьцуулахгүй (юань нь төгрөгөөс хэдэн зуу дахин бага). Бусад бүх баримтад "price_currency" = "MNT".',
  '',
  'Барааны нэрс:',
];

// matched_name: one of `names` or OCR_NOT_RELEVANT, never empty (the owner, 2026-10-23: mom corrects a
// wrong pick, an empty cell helps nobody).
const ocrSchema = (names) => ({
  type: 'OBJECT',
  properties: {
    document_type: { type: 'STRING' },
    date: { type: 'STRING', nullable: true },
    document_title: { type: 'STRING', nullable: true },
    header_receiver: { type: 'STRING', nullable: true },
    header_sender: { type: 'STRING', nullable: true },
    shop_line: { type: 'STRING', nullable: true },
    price_currency: { type: 'STRING', enum: ['MNT', 'CNY'], nullable: true },
    items: {
      type: 'ARRAY',
      items: {
        type: 'OBJECT',
        properties: {
          name: { type: 'STRING', nullable: true },
          matched_name: { type: 'STRING', enum: [...new Set([...names, OCR_NOT_RELEVANT])] },
          qty: { type: 'NUMBER' },
          unit_price: { type: 'NUMBER', nullable: true },
          line_total: { type: 'NUMBER', nullable: true },
          confidence: { type: 'NUMBER' },
        },
        required: ['name', 'matched_name', 'qty', 'confidence'],
      },
    },
    total_amount: { type: 'NUMBER', nullable: true },
    other_documents_visible: { type: 'STRING', nullable: true },
    issues: { type: 'ARRAY', items: { type: 'STRING' } },
    overall_confidence: { type: 'NUMBER' },
  },
  required: ['document_type', 'items', 'issues', 'overall_confidence'],
});

// The key or the model is wrong: every photo would fail the same way. So does a schema Gemini refuses
// (400 naming the schema); any other 400 is about one photo.
const OCR_FATAL_CODES = [401, 403, 404];
const ocrIsFatal = (code, text) => OCR_FATAL_CODES.includes(code) || (code === 400 && /schema/i.test(String(text)));

// catalogLines: "name" or "name | price₮" (ocr-catalog.gs)
function ocrPrompt(catalogLines) {
  return [...OCR_PROMPT_LINES, ...catalogLines].join('\n');
}

// names: the catalog's names, the only answers matched_name may take (with OCR_NOT_RELEVANT).
function ocrRequestBody(base64, mimeType, prompt, names) {
  return {
    contents: [{ role: 'user', parts: [{ inline_data: { mime_type: mimeType, data: base64 } }, { text: prompt }] }],
    generationConfig: {
      temperature: 0,
      thinkingConfig: { thinkingBudget: OCR_THINKING_BUDGET },
      responseMimeType: 'application/json',
      responseSchema: ocrSchema(names),
    },
  };
}

// { ok: true, result } | { ok: false, error, fatal } - fatal: the run stops and the photo stays
// on top, unread; any other failure files the photo with a "Шалгах" row (read only once).
function ocrParseResponse(code, text) {
  if (code !== 200) {
    return { ok: false, error: `HTTP ${code}: ${String(text).slice(0, 200)}`, fatal: ocrIsFatal(code, text) };
  }
  try {
    const parts = JSON.parse(text).candidates[0].content.parts.filter((part) => typeof part.text === 'string' && !part.thought);
    const result = JSON.parse(parts[parts.length - 1].text);
    if (!result || !Array.isArray(result.items)) {
      return { ok: false, error: 'хариу буруу хэлбэртэй', fatal: false };
    }
    return { ok: true, result };
  } catch (error) {
    return { ok: false, error: `хариуг уншиж чадсангүй: ${error.message}`, fatal: false };
  }
}

// catalog: { lines, names } (loadOcrCatalog)
function readReceiptWithGemini(blob, mimeType, key, catalog) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${OCR_MODEL}:generateContent`;
  const body = ocrRequestBody(Utilities.base64Encode(blob.getBytes()), mimeType, ocrPrompt(catalog.lines), catalog.names);
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': key },
      payload: JSON.stringify(body),
      muteHttpExceptions: true,
    });
    return ocrParseResponse(response.getResponseCode(), response.getContentText());
  } catch (error) {
    // a timeout or a dropped connection: not asked again (read once), a person looks at it
    return { ok: false, error: String(error.message || error), fatal: false };
  }
}
