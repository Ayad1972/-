# نقل أرقام الموظفين والخدمة من Excel إلى MFILE.DBF

أداة تنقل بيانات **رقم الموظف** و**الخدمة** من ملف Excel إلى جدول FoxPro القديم `MFILE.DBF` مع الحفاظ على:

- نفس ترميز الكتابة (غالباً `cp1256` للعربي)
- نفس Language Driver في هيدر DBF
- نفس عرض الحقول والرموز حتى لا تظهر مشاكل في نظام الفوكس القديم

## الملفات المطلوبة

- `H:\New Microsoft Excel Worksheet.xlsx`
- `H:\MFILE.DBF`

## التشغيل السريع على جهازك (Windows)

1. ثبّت Python إن لم يكن مثبتاً.
2. انسخ مجلد المشروع إلى جهازك، أو شغّل من نفس المجلد.
3. انقر مرتين على:

```bat
run_update_mfile.bat
```

أو من سطر الأوامر:

```bat
pip install -r requirements.txt
python update_mfile_from_excel.py --inspect-only
python update_mfile_from_excel.py --dry-run
python update_mfile_from_excel.py
```

## ماذا تفعل الأداة؟

1. تقرأ أعمدة رقم الموظف والخدمة من Excel تلقائياً.
2. تكتشف حقول DBF المناسبة تلقائياً.
3. تنشئ نسخة احتياطية من `MFILE.DBF` قبل أي تعديل.
4. تحدّث السجل إن وُجد نفس رقم الموظف، وإلا تضيف سجلاً جديداً.
5. تعيد كتابة بايت لغة الهيدر كما كان لضمان توافق الفوكس القديم.

## إذا لم تتعرف على أسماء الأعمدة

```bat
python update_mfile_from_excel.py --inspect-only
python update_mfile_from_excel.py --emp-excel-col "رقم الموظف" --service-excel-col "الخدمة" --emp-dbf-field "EMPNO" --service-dbf-field "SERVICE"
```

## أوضاع العمل

- `upsert` (افتراضي): تحديث إن وجد، وإلا إدراج
- `update`: تحديث فقط
- `append`: إدراج دائماً

مثال:

```bat
python update_mfile_from_excel.py --mode update
```

## ملاحظة مهمة

هذه البيئة السحابية لا تصل إلى قرص `H:` على جهازك.  
شغّل الأداة على Windows مباشرة حيث توجد الملفات، أو انسخ الملفين إلى مجلد `data` داخل المشروع.
