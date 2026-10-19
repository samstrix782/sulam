# סולם — שאלון ערכים לבחירות 2026

הדף שמוגש לגולשים, והמקור שממנו הוא נבנה.

| קובץ | מה זה |
|---|---|
| `index.html` | השאלון כולו בקובץ אחד. **לא עורכים ידנית** — נבנה מ-`app/` על ידי `build_bundle.py` |
| `app/parties.json` | המפלגות ותיוג העמדות שלהן (רמה, עיקרון, הערה, קישור, מקור) |
| `app/questions.json` | הערכים, השאלות והשלבים בסולם |
| `app/index.html` | תבנית הדף; הנתונים נכנסים במקום `/*DATA*/` |
| `build_bundle.py` | מאחד את שלושת הקבצים שב-`app/` ל-`index.html` |
| `config.js` | כתובת ה-Web app של Apps Script. השורה היחידה שמשתנה בין סביבות |
| `.nojekyll` | מבקש מ-GitHub Pages להגיש את הקבצים כמו שהם |

## לעדכן את השאלון
1. לערוך את הקבצים ב-`app/`
2. `python3 build_bundle.py`
3. לבדוק שהדף נטען, ולדחוף

## תיוג
כל תיוג ב-`app/parties.json` הוא רמה בסולם (1–5), סימון עיקרון, הערה קצרה וקישור למקור.
בסבב עדכון: לבדוק כל רשימה על אותן שאלות ליבה, לשנות רק מול מקור שנפתח ונבדק,
ולתת לאותה הצבעה את אותה רמה אצל כל המפלגות שהצביעו כך. הסבב האחרון: 19.10.2026.

תיוג שמאושר במסך התיוג נשמר בגיליון (`tags`) וגובר בזמן ריצה על מה שב-`app/parties.json`.
`source: "estimate"` מסמן הערכה שטרם אושרה.

## השרת (Apps Script)
השמירה, הניתוח המילולי והמיילים רצים ב-Google Apps Script שצמוד לגיליון.
`config.js` הוא מה שמחבר ביניהם. הקוד נמצא ב-`apps-script/Code.gs`.

**פריסה אוטומטית:** כל מיזוג ל-main שמשנה את `apps-script/` מריץ את
`.github/workflows/deploy-apps-script.yml`, שדוחף את הקוד ל-Apps Script ופורס גרסה
חדשה באותה כתובת. אפשר גם להריץ ידנית: Actions → פריסת Apps Script → Run workflow.

הגדרה חד-פעמית:
1. https://script.google.com/home/usersettings → להפעיל את Google Apps Script API
2. במחשב (או ב-Google Cloud Shell, shell.cloud.google.com, שעובד גם מהטלפון):
   `npx @google/clasp@3.4.1 login --no-localhost`, לאשר בדפדפן, ואז `cat ~/.clasprc.json`
3. בריפו: Settings → Secrets and variables → Actions
   - Secrets → New repository secret: השם `CLASPRC_JSON`, הערך — כל התוכן של הקובץ מסעיף 2
   - Variables → New repository variable: השם `APPS_SCRIPT_ID`, הערך `1bi0quvVKx-Kxzwm2kCm-1IoMz08lKUKJbpr7nSmmch-3Iml42Z71TDLo`

בלי ההגדרה הזאת הפריסה האוטומטית מדלגת (אזהרה ב-Actions, לא כישלון).

**כשהקוד צריך הרשאה חדשה מגוגל** (למשל טריגרים או שירות שלא היה בשימוש), הפריסה
האוטומטית לא יכולה לאשר אותה: אחרי הפריסה צריך להריץ פונקציה אחת מהעורך (Run) ולאשר.
עד אז ה-Web app עלול להחזיר שגיאת הרשאה — לכן שינוי כזה עדיף להדביק ולאשר ידנית לפני המיזוג.

בפעם הראשונה בלבד: להריץ את `setupTrigger` מהעורך.

`setupTrigger` יוצר טריגר שרץ כל 5 דקות: משלים ניתוחים שנכשלו (כולל משיבים ותיקים,
בלי לשלוח להם מייל) ושולח את מייל הניתוח למי שסימן זאת בטופס.
`previewMail` שולח לך מייל לדוגמה, ו-`pingModel` בודק שמפתח הניתוח עובד.
