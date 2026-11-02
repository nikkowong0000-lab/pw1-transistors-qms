# Advanced Time Management — PW1 Transistors（QMS Operations）

**可提交主檔：** [`PW1_TM_SUBMIT.html`](./PW1_TM_SUBMIT.html)  
**Mahara 貼文：** [`PW1_TM_Mahara_Paste.md`](./PW1_TM_Mahara_Paste.md)  
**任務 CSV：** [`PW1_GANTT_tasks.csv`](./PW1_GANTT_tasks.csv)  
**教學頁（可不交）：** [`PW1_Time_Management.html`](./PW1_Time_Management.html)

---

## 怎麼交

1. 打開 `PW1_TM_SUBMIT.html`，改黃色欄（組員／工時／風險）
2. 列印 PDF 或截圖 GANTT（含紅／藍／橙關鍵路徑）
3. 風險短文（50–100 字）貼 Mahara
4. 本資料夾檔案列入 Documents gallery

---

## 工時假設（預設，請在 SUBMIT 頁改成你們組的）

- 6 人；Mon–Fri 09:30–17:00；午休 12:30–13:30
- 淨工時 ≈ 6.5 h／人／日
- 並行：A 量測｜B 閘／加器｜C QMS／文件

## PERT

\( t_e = (O + 4M + P) / 6 \)  
例：Half adder O=1.5, M=2.5, P=4 → \(t_e≈2.58\) h

## Critical paths

- **CP-A 紅：** Kit → Gates → Half adder → FA 演練 → Fri 10:15
- **CP-B 藍：** QMS setup → Audit prep → Thu 11:15
- **CP-C 橙：** Research → 簡報／影片 → Thu 晚 → Fri 12:00

## Buffers

| 類型 | 位置 | 時長 |
|------|------|------|
| 任務 | 半加器後 Wed PM | 1 h |
| 任務 | QMS 後 | 1 h |
| 專案末 | Thu 16:00–17:00 | 1 h |
| 專案末 | Fri 09:30–10:00 | 0.5 h |

## 前三大排程風險（≈90 字）

半加器除錯逾時會直接擠壓週五全加器演練；QMS 檔名未凍結會讓週四 11:15 稽核找檔失敗；個人影片常低估剪輯而拖到週五中午。對策：週三設半加器檢查點並配置任務緩衝、週三凍結 QMS 結構並演練 30 秒找檔、週四晚強制影片定稿再於週五 12:00 前提交 Moodle。
