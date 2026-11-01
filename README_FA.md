# Appearance AI — Dataset 2 / 90 images / 7-category supervised classification

این پروژه فقط برای **90 تصویر دوم** نوشته شده و هیچ دسترسی‌ای به دیتاست 75 تصویر اول ندارد. کل پردازش روی رایانه شخصی انجام می‌شود؛ هیچ API و توکنی مصرف نمی‌شود.

## هفت دسته

`Color`, `Transparency`, `Glossiness`, `Texture`, `Shape`, `Material`, `Other`.

چهار دسته **ظاهر** فقط `Color`, `Transparency`, `Glossiness`, `Texture` هستند. سه دسته دیگر در گزارش وجود دارند، اما هنگام شمارش تعداد ویژگی‌های اصلی ظاهر، حساب نمی‌شوند.

`Unassigned` دسته هشتم نیست؛ فقط وضعیت برچسب‌هایی است که هنوز توسط پژوهشگر مشخص نشده‌اند.

## فایل‌های ورودی

فایل‌های شخصی خود را در این مسیرها قرار دهید:

```text
Appearance_AI_90_7Categories_VSCode/
├── dataset_2/
│   ├── input/
│   │   ├── Appearance_Part2_90_FINAL.xlsx
│   │   └── images2.zip
│   ├── database/
│   └── output/
├── src/
│   ├── core.py
│   └── reports.py
├── tests/
│   └── test_pipeline.py
├── main.py
├── app.py
└── requirements.txt
```

اکسل واقعی باید شیت `Word_Level_1800` و ستون‌های `row_id`, `image_name`, `model_key`, `position`, `word` را داشته باشد. دو ستون `five_words` و `image_md5` در صورت وجود برای نمایش و پیگیری داده استفاده می‌شوند. رکوردها باید دقیقاً 1800 مورد (90 × 4 × 5) باشند.

تصاویر را **از ZIP استخراج نکنید**؛ برنامه آنها را بر اساس نام `001.jfif` تا `090.jfif` از ZIP می‌خواند، حتی اگر در پوشه داخلی `images/` باشند.

## نصب در PowerShell داخل VS Code

پیش از اجرا، از درست بودن مسیر ترمینال در پوشه اصلی پروژه اطمینان حاصل کنید.

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

این نسخه از `artifact-tool` استفاده نمی‌کند و فایل‌های `exe`, `bat`, `dll` ندارد.

## اجرای برنامه

```powershell
.\.venv\Scripts\python.exe main.py init
.\.venv\Scripts\python.exe main.py status
.\.venv\Scripts\python.exe -m streamlit run app.py
```

Streamlit معمولاً در `http://localhost:8501` باز می‌شود. ترمینال Streamlit را باز نگه دارید. برای توقف برنامه از `Ctrl+C` استفاده کنید.

## دسته‌بندی تعاملی

پنل برای هر مورد **تصویر اصلی، نام مدل، شماره تصویر، موقعیت واژه و پنج‌واژه پاسخ مدل** را نمایش می‌دهد. دو نوع تصمیم وجود دارد:

- `All occurrences of this word`: برچسب انتخابی روی همه تکرارهای آن واژه در **همین دیتاست 90 تصویری** اعمال می‌شود.
- `Only this image + model + word position`: برچسب فقط برای همان استفاده از آن واژه ثبت می‌شود. این تصمیم اختصاصی حتی اگر برچسب عمومی واژه بعدها عوض شود، باقی می‌ماند.

در تب `Review / Edit decisions` می‌توان برچسب‌های قبلی را اصلاح کرد، واژه استاندارد (canonical) ثبت کرد یا تصمیم اختصاصی را حذف کرد تا واژه از برچسب عمومی پیروی کند.

تمام تصمیم‌ها داخل SQLite در `dataset_2/database/classifications.sqlite3` ذخیره می‌شوند و پس از بستن Streamlit باقی می‌مانند. ورودی Excel و ZIP هرگز تغییر نمی‌کنند.

**مهم:** اگر Excel ورودی پس از راه‌اندازی تغییر کند، `init` برای حفظ تصمیم‌های قبلی از تغییر خودکار دیتابیس امتناع می‌کند. قبل از جایگزینی فایل ورودی، از دیتابیس بکاپ بگیرید و آگاهانه انتقال دهید.

## تفاوت Cleaning و دسته‌بندی

`raw_word` متن اصلی استخراج‌شده از Excel است. `word_key` فقط حروف و فاصله‌ها را یکسان‌سازی می‌کند. `canonical_word` واژه استانداردی است که پژوهشگر می‌تواند تعیین کند؛ مثلا `grey` را به `gray` تبدیل کند. خود واژه خام برای بازبینی پژوهشی همیشه حفظ می‌شود.

دسته‌بندی و نرمال‌سازی دو عملیات جدا هستند. تغییر canonical به معنی تغییر دسته یا یکی دانستن واژه‌هایی با معنی متفاوت نیست.

## خروجی پژوهشی

بعد از دسته‌بندی دلخواه، در پنل `Export reports` کلیک کنید یا در یک ترمینال مستقل VS Code بنویسید:

```powershell
.\.venv\Scripts\python.exe main.py report
```

خروجی هر بار در مسیر جدید `dataset_2/output/analysis_<timestamp>/` ساخته می‌شود:

- یک Excel با شیت‌های `Word_Dictionary`, `Word_Instances_1800`, `Image_Model_360`, `Image_Summary_90`, `Coverage_Images`, `Coverage_Model_Responses`, `Raw_Frequency`, `Clean_Frequency`, `Raw_Model_Matrix`, `Clean_Model_Matrix`, `Category_Model`, `Top_Words_By_Category`, `Needs_Review`, `Decision_History`, `Embedded_Charts` و `Read_Me`
- CSV متناظر برای تمام جدول‌ها
- 14 نمودار PNG؛ نمودار مقایسه واژه‌ها در چهار مدل، Top15 هر مدل، تعداد و درصد هفت دسته و پوشش چهار ویژگی ظاهر
- Snapshot دیتابیس SQLite و `summary.json` برای مستندسازی نتایج

### محاسبه درصد تصاویر دارای 4، 3، 2 یا 1 ویژگی ظاهر

برای هر تصویر، تمام 20 توصیف‌گر چهار مدل تجمیع می‌شوند و حضور دسته‌های یکتای **Color / Transparency / Glossiness / Texture** بررسی می‌شود. مثلاً دو واژه `blue` و `red` هر دو Color هستند و همچنان فقط **یک دسته** حساب می‌شوند. Shape، Material و Other در شمارش این چهار دسته منظور نمی‌شوند.

علاوه بر آن، محاسبه **جداگانه برای هر مدل و تصویر** (5 توصیف‌گر) با واحد آماری `image-model` نیز ارائه می‌شود. بدین ترتیب واحدهای آماری با هم مخلوط نمی‌شوند.

تا زمانی که یک تصویر یا پاسخ مدل واژه `Unassigned` داشته باشد، عدد نهایی تنوع ویژگی‌هایش `N/A` می‌ماند. فایل `Coverage_Images` دو درصد را از هم جدا می‌کند:

1. درصد نسبت به **تمام 90 تصویر** (شامل موارد دسته‌بندی‌نشده در مخرج)
2. درصد نسبت به **تصاویر کاملاً دسته‌بندی‌شده** (فقط موارد واجد شرایط در مخرج)

در هر دو حالت تعداد تصاویر دسته‌بندی‌نشده هم درج می‌شود. پس از دسته‌بندی کامل همه واژه‌ها، این دو درصد برابر خواهند بود.

## تست‌ها

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

تست‌ها با یک Excel آزمایشی ساختگی 1800 سطری انجام می‌شوند و هیچ‌کدام از فایل‌های اصلی را تغییر نمی‌دهند. خروجی واقعی پژوهش تنها با فایل واقعی شما تولید می‌شود.
