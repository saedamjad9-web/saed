# تقرير التدقيق والنقد — نموذج الاستهلاك الشهري الفعلي للمواد الخام 2026
# Audit & Critique Report — RM Actual Monthly Consumption Model 2026

الملف الخاضع للمراجعة: `output/RM_Consumption_Model_2026.xlsx` (562 مادة × 9 أشهر، كانون الثاني – أيلول 2026).
Reviewed file: `output/RM_Consumption_Model_2026.xlsx` (562 items × 9 months, Jan–Sep 2026).
مُولِّد الملف: `tools/build_rm_consumption_model.py`، والمصدر: `data/rm/rm_1.xlsx … rm_9.xlsx`.

---

## 1. منهجية العمل / Method

| الدور / Role | المهمة / Task | النتيجة / Result |
|---|---|---|
| وكيل البناء / Builder | توليد المصنف بمعادلات حية تشير إلى الأوراق الخام M01..M09 | تم / Done |
| المدقق أ (رقمي) / Auditor A (numeric) | إعادة حساب كل خلية من ملفات المصدر بـ pandas بشكل مستقل ومقارنتها بنتائج LibreOffice | PASS |
| المدقق ب (معادلات وبنية) / Auditor B (formulas & structure) | فحص كل معادلة ومرجع وتنسيق، ومطابقة الأوراق الخام مع المصدر | PASS |
| إعادة التدقيق / Re-audit | إعادة تشغيل المدقق أ بعد تعديل قاعدة كلفة الإدخال | PASS |
| الناقد 1 / Critic 1 | رأي من منظور محاسبة التكاليف | أدناه / below |
| الناقد 2 / Critic 2 | رأي من منظور تخطيط المخزون والمشتريات | أدناه / below |

---

## 2. نتائج التدقيق / Audit results

### 2.1 التدقيق الرقمي / Numeric audit (Auditor A, re-run on final build)

- 60,696 خلية شهرية (562 × 9 × 12) أُعيد حسابها من المصدر: صفر فروقات، أقصى فرق مطلق 1.9e-09.
- 13 عمود ملخص (DH:DT) × 562: صفر فروقات. عمود Check يساوي صفراً لكل المواد (أقصى قيمة مطلقة 6.9e-11).
- تسلسل الأرصدة بين الأشهر: صفر انقطاع في الكمية وصفر انقطاع في الكلفة عبر 8 انتقالات × 562 مادة.
- الأوراق M01..M09 مطابقة لملفات المصدر خلية بخلية (9 × 18,612 خلية، شاملة التواريخ وصف الإجماليات).
- ورقة Summary: 22 فئة فرعية بلا نقص أو تكرار، 154 خلية مطابقة، الإجمالي العام = مجموع الفئات = إجماليات Model.
- صفر خلايا خطأ أو نص في منطقة الأرقام.

Numeric audit: 60,696 monthly cells and all 13 summary columns recompute from the raw files with zero mismatches; Check = 0 for all 562 items; month-to-month continuity holds for quantity and cost; raw sheets are cell-for-cell identical to the source.

### 2.2 تدقيق المعادلات والبنية / Formula & structure audit (Auditor B)

- 67,882 معادلة في D3:DT564 تطابق النمط المتوقع تماماً: صفر مراجع إلى ورقة خاطئة، صفر انزياح صفوف، صفر مراجع مطلقة، لا ثوابت رقمية سوى 0 و1 و9.
- كل عملية قسمة على قيمة خلية (26,976) محاطة بـ IFERROR.
- تنسيقات الأرقام متسقة في كل الأعمدة (0.0000 لكلفة الوحدة، #,##0.00 للكميات والكلف، 0.0% للنسبة).
- التجميد والفلترة وعروض الأعمدة سليمة؛ الملف يُعاد حسابه في LibreOffice في 3 ثوانٍ بلا أخطاء.

Formula audit: all 67,882 Model formulas match the expected pattern; every division is IFERROR-guarded; number formats are consistent; the file recalculates without errors.

### 2.3 ما تغيّر بعد التدقيق / Changes made after audit

1. **قاعدة كلفة الإدخال.** في النسخة الأولى، عند وجود مشتريات وإدخالات غير شرائية في الشهر نفسه (تسويات AddQty، استلام إنتاج ProdRecQty، مرتجعات)، كانت الوحدات غير الشرائية تدخل بكلفة صفر ويمتصها رقم كلفة الإخراج. أثّر ذلك في 94 شهر-مادة (~72 ألف وحدة). القاعدة الآن:
   `InCost = Purchase Amount + (TotalInQty − RecQty) × UC` عند وجود شراء، وإلا `TotalInQty × UC`، حيث `UC` = كلفة الوحدة الافتتاحية إن وُجد رصيد افتتاحي مقيَّم، وإلا الكلفة المعيارية.
   الأثر: إجمالي كلفة الإخراج 9 أشهر ارتفع من 38,244,633 إلى 38,385,238 (+140,604)، وحالات كلفة الإخراج السالبة انخفضت من 11 إلى 5، وأكبر ثلاث حالات (−2,978 و−2,562 و−1,307) اختفت.
   InCost rule fixed: non-purchase inflow units are now valued at the opening unit cost; 9-month out cost rose by 140,604; negative-OutCost cases fell from 11 to 5.
2. **README.** أُضيف نص القاعدة الدقيق، وتحذير بأن كلفة الإخراج رقم موازنة يمتص إعادات التقييم، وتحذير من فرز أو حذف صفوف الأوراق الخام.
3. **Summary.** عمود النسبة صار "حصة من إجمالي كلفة الإخراج %" بدل "حصة من الإجمالي %".

### 2.4 ملاحظات متبقية في بيانات المصدر (ليست أخطاء في النموذج) / Residual source-data notes

| الملاحظة / Note | العدد / Count | أمثلة / Examples |
|---|---|---|
| كلفة إخراج سالبة (إعادة تقييم في النظام) / Negative OutCost | 5 شهر-مادة | RM.NUT.0058 M02 −40.49؛ RM.NUT.00010 M01 −11.01؛ RM.ADTV.0072 M06 −8.02 |
| كلفة إخراج بلا كمية / OutCost with zero OutQty | 5 (كلها ≤ 8.02) | RM.BTQ.0008 M01 −8.41 |
| كلفة نهاية بلا كمية / EndCost with zero EndQty | 1 | RM.FLV.0073 M07 −0.038 (كلفة معيارية سالبة في النظام) |
| إدخال بكمية وبلا كلفة متاحة / InQty>0, InCost=0 | 1 | RM.BIS.0023 M03 (الكلفة الافتتاحية والمعيارية صفر) |
| مواد كلفتها المعيارية صفر / Items with zero standard cost | 58 (10 منها بحركة) | — |
| إدخالات بلا قيمة شراء في المصدر (مقدَّرة) / Inflows without Purchase Amount (estimated) | 204 شهر-مادة (142 مادة) | — |
| فرق فعلي/معياري يتجاوز ±50% / Actual vs standard beyond ±50% | 9 مواد | RM.ADTV.00024 +4073%؛ RM.NUT.00004 +128%؛ RM.CLR.00035 −98% |

### 2.5 حدود التدقيق / Audit limits

- القيم حُسبت بـ LibreOffice لا بـ Excel؛ الدوال المستخدمة (IF, AND, MAX, IFERROR, SUM, COUNTIF, COUNTIFS, SUMIF) متطابقة السلوك في الاثنين، لكن Excel نفسه لم يُشغَّل.
- التدقيق يؤكد أن النموذج ينفّذ القواعد المعلنة بدقة، ولا يحكم على صحة سياسة التكلفة نفسها؛ ذلك موضوع قسم النقد.
- الملف لا يخزّن قيماً محسوبة؛ أي أداة لا تعيد الحساب (pandas أو openpyxl أو بعض المعاينات) ستقرأ فراغات حتى يُفتح الملف في Excel ويُحفظ مرة.

Values were obtained via LibreOffice, not Excel; the audit verifies implementation of the stated rules, not the costing policy itself; the file stores no cached values until opened and saved once in Excel.

---
