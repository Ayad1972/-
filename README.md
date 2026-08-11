# نقل أرقام الموظفين والخدمة من Excel إلى MFILE.DBF

أداة تنقل بيانات **رقم الموظف** و**الخدمة** من ملف Excel إلى جدول FoxPro القديم `MFILE.DBF` مع الحفاظ على:

- نفس ترميز الكتابة (غالباً `cp1256` للعربي)
- نفس Language Driver في هيدر DBF
- نفس عرض الحقول والرموز حتى لا تظهر مشاكل في نظام الفوكس القديم

## الملفات المطلوبة

- `H:\New Microsoft Excel Worksheet.xlsx`
- `H:\MFILE.DBF`

## تثبيت Python على Windows (إذا غير موجود)

1. انقر مرتين على:

```bat
install_python.bat
```

سيقوم الملف بـ:
- تثبيت Python تلقائياً
- إضافته إلى PATH
- تثبيت المكتبات المطلوبة (`openpyxl`, `dbfread`, `dbf`)

إذا فشل التلقائي، حمّل Python من:
https://www.python.org/downloads/

**مهم جداً أثناء التثبيت:** فعّل الخيار  
`Add python.exe to PATH`

## التشغيل السريع على جهازك (Windows)

1. ثبّت Python عبر `install_python.bat` إن لم يكن مثبتاً.
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

## حصر الملفات/المجلدات المتطابقة 100%

أولاً تقرير فقط، ثم الإبقاء على الأحدث:

```bat
find_dupes.bat
find_dupes.bat APPLY
```

الدليل: [`FIND_DUPLICATES.md`](FIND_DUPLICATES.md)

## تسريع الحاسوب خطوة بخطوة

إذا كان الجهاز بطيئاً أثناء العمل على الفوكس/Excel، ضع الملف التالي على سطح المكتب وشغّله كمسؤول:

```bat
go.bat
```

من القائمة اختر `9` ثم أعد تشغيل الجهاز.

الدليل: [`SPEED_UP_PC.md`](SPEED_UP_PC.md)  
(بديل متقدم: `speed_up_pc.bat` مع `speed_up_pc.ps1`)

## ملاحظة مهمة

هذه البيئة السحابية لا تصل إلى قرص `H:` على جهازك.  
شغّل الأداة على Windows مباشرة حيث توجد الملفات، أو انسخ الملفين إلى مجلد `data` داخل المشروع.
