# Excel AI Helper — Bina Kisi API Key Ke

Apna Excel sawaal Hindi, Hinglish ya English me likho, system sahi formula + DEO daily-life example ke saath jawab dega.

## Kaise chalta hai — saaf baat

Ye ek **rule-based helper** hai, koi bada language model (LLM) nahi. Ye tumhare sawaal ke shabdon ko apne knowledge base (25 topics: Sum, Average, CountIF, VLOOKUP, Pivot Table, Duplicate, Data Filter, Advance Excel, aur zyada) se milata hai aur sabse achha match wala jawab deta hai.

Isliye isme:
- **Koi API key nahi chahiye**
- **Koi server nahi, koi internet nahi**
- **Koi data browser se bahar nahi jaata**

## Kaise use karein

`index.html` ko browser me kholo, sawaal likho, Enter dabao.

Examples:
- "total kaise nikale" → SUM
- "Present kitne hain kaise ginu" → COUNTIF
- "duplicate hatane hain" → Remove Duplicates
- "ID se naam kaise nikalu" → VLOOKUP
- "mahine ka summary kaise banau" → Pivot Table

Agar sawaal samajh na aaye to system saare topics ki list dikha dega.

## Files

- `index.html` — helper ka page
- `helper.js` — knowledge base + matching engine (browser aur Node dono me chalta hai)

## Test

```bash
node -e "const h = require('./helper.js'); console.log(h.findAnswer('duplicate hatane hain').title); console.log(h.findAnswer('total kaise nikale').formula);"
```

## Note

Ye ek learning project hai, DEO practice ke liye. Ye practice aid hai, test ya job guarantee nahi.
