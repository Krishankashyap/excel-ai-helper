// Excel AI Helper — knowledge base and matching engine.
// IMPORTANT, stated honestly in the README too: this is a rule-based helper,
// NOT a large language model. It needs NO API key, NO server, NO internet.
// It matches your question (Hindi/Hinglish/English words) to the best Excel
// topic in its knowledge base and returns that topic's formula + DEO example.
// No data leaves your browser.

const TOPICS = [
  { id: "sum", title: "SUM - total nikalo", keywords: ["sum", "total", "jod", "jodo", "add", "yog"], formula: "=SUM(C2:C100)", explanation: "Din bhar ki sales, attendance ya amount ka total ek formula se. C2:C100 me jo numbers hain, sab jud jayenge.", example: "DEO example: Aaj ki total entry =SUM(C2:C100)" },
  { id: "average", title: "AVERAGE - ausat nikalo", keywords: ["average", "ausat", "avg", "mean"], formula: "=AVERAGE(B2:B50)", explanation: "Numbers ka ausat deta hai. Khaali cells apne aap ignore ho jaati hain.", example: "DEO example: Class ke marks ka ausat =AVERAGE(B2:B50)" },
  { id: "averageif", title: "AVERAGEIF - ek group ka ausat", keywords: ["averageif", "group average", "ausat group"], formula: '=AVERAGEIF(A2:A100, "Delhi", B2:B100)', explanation: "Sirf ek group ka ausat — jaise sirf Delhi wale students ka ausat marks.", example: "DEO example: Sirf Delhi ka ausat =AVERAGEIF(A2:A100, \"Delhi\", B2:B100)" },
  { id: "count", title: "COUNT - kitni entries hain", keywords: ["count", "gino", "ginti", "kitne"], formula: "=COUNT(A2:A500)", explanation: "Range me kitne cells me numbers hain, ye ginata hai. Text wale cells nahi ginta.", example: "DEO example: Aaj kitni number entries hui =COUNT(A2:A500)" },
  { id: "countif", title: "COUNTIF - Present / Absent gino", keywords: ["countif", "present", "absent", "ginna shart"], formula: '=COUNTIF(D2:D100, "Present")', explanation: "Shart par ginata hai — jaise attendance me Present kitne hain. Absent ke liye 'Absent' likh do.", example: "DEO example: Present ki ginti =COUNTIF(D2:D100, \"Present\")" },
  { id: "countifs", title: "COUNTIFS - do sharton par gino", keywords: ["countifs", "do shart", "dono shart"], formula: '=COUNTIFS(A2:A100, "Delhi", D2:D100, "Present")', explanation: "Do ya zyada sharton par ek saath ginata hai — jaise Delhi + Present dono ho.", example: "DEO example: Delhi se Present =COUNTIFS(A2:A100, \"Delhi\", D2:D100, \"Present\")" },
  { id: "sumif", title: "SUMIF - ek group ka total", keywords: ["sumif", "group total", "ek ka total"], formula: '=SUMIF(A2:A100, "Delhi", C2:C100)', explanation: "Sirf un rows ka total karta hai jo shart poori karti hain — jaise sirf Delhi ki sales.", example: "DEO example: Delhi ki total sales =SUMIF(A2:A100, \"Delhi\", C2:C100)" },
  { id: "sumifs", title: "SUMIFS - mahina + sheher ka total", keywords: ["sumifs", "mahina total", "monthly total"], formula: '=SUMIFS(C2:C100, A2:C100, "Delhi", B2:B100, "Nov")', explanation: "Do ya zyada sharton par total — jaise November + Delhi dono ho. Monthly report me roz kaam aata hai.", example: "DEO example: Nov + Delhi ka total =SUMIFS(C2:C100, A2:A100, \"Delhi\", B2:B100, \"Nov\")" },
  { id: "if", title: "IF - Pass / Fail lagao", keywords: ["if", "pass", "fail", "faisla", "shart"], formula: '=IF(B2>=33, "Pass", "Fail")', explanation: "Shart sahi ho to pehla result, galat ho to doosra. Marks sheet me Pass/Fail apne aap lagao.", example: "DEO example: 33 se zyada par Pass =IF(B2>=33, \"Pass\", \"Fail\")" },
  { id: "iferror", title: "IFERROR - error chhupao", keywords: ["iferror", "error", "galti", "div"], formula: "=IFERROR(A2/B2, 0)", explanation: "Formula me error aaye to uski jagah diya hua value dikhata hai, #DIV/0! nahi dikhega.", example: "DEO example: Divide error par 0 =IFERROR(A2/B2, 0)" },
  { id: "vlookup", title: "VLOOKUP - ID se naam / rate nikalo", keywords: ["vlookup", "lookup", "dhoondho", "id se naam", "rate"], formula: "=VLOOKUP(A2, $D$2:$F$100, 2, FALSE)", explanation: "Pehle column me value dhoondhta hai aur usi row ke doosre column ka value deta hai. FALSE ka matlab exact match.", example: "DEO example: Employee ID se naam =VLOOKUP(A2, $D$2:$F$100, 2, FALSE)" },
  { id: "xlookup", title: "XLOOKUP - Advance lookup", keywords: ["xlookup", "advance lookup", "naya lookup"], formula: "=XLOOKUP(A2, D2:D100, F2:F100)", explanation: "VLOOKUP ka naya, simple version. Sirf naye Excel / Microsoft 365 me chalta hai, purane me VLOOKUP use karo.", example: "DEO example: ID se naam =XLOOKUP(A2, D2:D100, F2:F100)" },
  { id: "trim", title: "TRIM - extra space hatao", keywords: ["trim", "space", "saaf", "extra"], formula: "=TRIM(A2)", explanation: "Text ke aage-peeche aur beech ke extra space hata deta hai. Saaf data ke liye zaroori.", example: "DEO example: Naam saaf karo =TRIM(A2)" },
  { id: "proper", title: "PROPER - naam ek jaisa likho", keywords: ["proper", "naam", "capital", "case"], formula: "=PROPER(A2)", explanation: "Har word ka pehla letter capital karta hai — 'krishan kashyap' se 'Krishan Kashyap'.", example: "DEO example: Naam theek karo =PROPER(A2)" },
  { id: "left", title: "LEFT / RIGHT - code ka hissa nikalo", keywords: ["left", "right", "shuru", "aakhri", "hissa"], formula: "=LEFT(A2, 4)", explanation: "Text ke shuru ke letters LEFT se, aakhri ke RIGHT se, beech ke MID se nikalte hain.", example: "DEO example: Code ke pehle 4 digit =LEFT(A2, 4)" },
  { id: "textjoin", title: "TEXTJOIN - naam + sheher jodo", keywords: ["textjoin", "jodo text", "combine"], formula: '=TEXTJOIN(", ", TRUE, A2:B2)', explanation: "Range ka text comma se jodta hai. TRUE ka matlab khaali cells skip karo.", example: "DEO example: Naam + sheher =TEXTJOIN(\", \", TRUE, A2:B2)" },
  { id: "filter", title: "Data Filter - chhant kar dikhao", keywords: ["filter", "chhant", "chhanto", "sheher wise"], formula: "Data > Filter (Ctrl + Shift + L)", explanation: "Sirf chahiye wali rows dikhata hai, baaki chhup jaati hain. Sheher-wise list nikalne ka sabse tez tareeka.", example: "DEO example: Sirf Delhi wali rows dekhne ke liye Filter lagao, phir Delhi chun lo." },
  { id: "duplicate", title: "Duplicate - double entry hatao / pakdo", keywords: ["duplicate", "double", "nakli", "do baar", "repeat"], formula: "Data > Remove Duplicates", explanation: "Do tareeke: (1) Remove Duplicates se double rows ek click me hat jayengi. (2) COUNTIF se pehle check karo: =COUNTIF($A$2:$A$100, A2) — 1 se zyada aaye to duplicate hai.", example: "DEO example: Customer do baar entry ho gaya ho to Remove Duplicates use karo, pehle backup le lena." },
  { id: "pivot", title: "Pivot Table - summary minute me", keywords: ["pivot", "summary", "report", "mahina summary"], formula: "Insert > PivotTable", explanation: "Hazaaron rows ka summary minute me banata hai — har sheher ki total sales, har item ki ginti. Naya data aane par right-click > Refresh karna mat bhoolna.", example: "DEO example: Mahine ki sales ka summary Pivot se banao, haath se nahi." },
  { id: "freeze", title: "Freeze Panes - header jamayo", keywords: ["freeze", "header", "jamao", "upar"], formula: "View > Freeze Panes", explanation: "Upar ki header row jama deta hai, neeche scroll karne par bhi header dikhega.", example: "DEO example: 1000 rows wali sheet me header jamane ke liye Freeze Panes." },
  { id: "validation", title: "Data Validation - galat entry roko", keywords: ["validation", "galat entry", "dropdown", "list"], formula: "Data > Data Validation > List", explanation: "Ek column me sirf diye hue options hi daale ja sakein — jaise sirf Present / Absent. Galat spelling ki entry hi nahi hogi.", example: "DEO example: Attendance column me sirf Present/Absent ka dropdown lagao." },
  { id: "texttocolumns", title: "Text to Columns - ek cell todo", keywords: ["text to columns", "todo", "alag column", "comma"], formula: "Data > Text to Columns", explanation: "Ek cell me comma se juda data alag columns me tod deta hai.", example: "DEO example: 'Krishan, Delhi, 25' ko naam, sheher, umar me todo." },
  { id: "pastespecial", title: "Paste Special Values - sirf result rakho", keywords: ["paste special", "values", "formula hatao"], formula: "Ctrl + Alt + V > Values", explanation: "Formula ki jagah sirf uska result chipkata hai, taaki aage formula toot na jaye.", example: "DEO example: Boss ko bhejne se pehle sheet ko Values me badlo." },
  { id: "findreplace", title: "Find & Replace - spelling ek saath theek karo", keywords: ["find", "replace", "spelling", "badlo"], formula: "Ctrl + H", explanation: "Ek word ko poori sheet me doosre word se badalta hai.", example: "DEO example: 'Dehli' ko 'Delhi' poori sheet me ek baar me badlo." },
  { id: "sumproduct", title: "SUMPRODUCT - Advance total", keywords: ["sumproduct", "advance", "bina column"], formula: '=SUMPRODUCT((A2:A100="Delhi")*(C2:C100))', explanation: "Bina extra column banaye, shart wali rows ka total. Advance level ka formula.", example: "DEO example: Delhi ka total bina helper column ke =SUMPRODUCT((A2:A100=\"Delhi\")*(C2:C100))" },
];

function normalize(q) {
  return (q || "").toLowerCase().replace(/[^a-z0-9 ]/g, " ");
}

function findAnswer(question) {
  const q = normalize(question);
  if (!q.trim()) return null;
  const tokens = new Set(q.split(/\s+/).filter(Boolean));
  const hasKeyword = (kw) => {
    if (kw.includes(" ")) return q.includes(kw);
    return tokens.has(kw);
  };
  let best = null;
  let bestScore = 0;
  for (const topic of TOPICS) {
    let score = 0;
    for (const kw of topic.keywords) {
      if (hasKeyword(kw)) score += kw.length;
    }
    // Bonus for exact topic id as a whole word in the question
    if (tokens.has(topic.id)) score += 5;
    if (score > bestScore) {
      bestScore = score;
      best = topic;
    }
  }
  return bestScore > 0 ? best : null;
}

function getTopicById(id) {
  
    return TOPICS.find((t) => t.id === id) || null;
}
function listTopics() {
  return TOPICS.map((t) => t.title);
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { TOPICS, findAnswer, listTopics, getTopicById };
}
