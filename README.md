# 🌐 Universal Dataset → Excel Website

## 🚀 LIVE WEBSITE

### 👉 https://universal-dataset-to-excel.onrender.com

এই website ব্যবহার করে কোনো Python, Terminal, Command বা software install না করেই একটি public dataset/file URL থেকে Excel (`.xlsx`) file তৈরি করা যাবে।

---

## ✨ কীভাবে ব্যবহার করবেন?

শুধু ৪টি ধাপ:

### 1️⃣ Website খুলুন

👉 https://universal-dataset-to-excel.onrender.com

### 2️⃣ Dataset URL দিন

**Dataset URL** box-এ আপনার public dataset/file URL paste করুন।

### 3️⃣ Convert করুন

**Convert to Excel** button-এ click করুন।

### 4️⃣ Excel Download করুন

Conversion শেষ হলে **Download Excel** button দেখা যাবে।

সেখানে click করলেই `.xlsx` file download হবে।

---

# 🔗 কোন ধরনের URL দেওয়া যাবে?

URL অবশ্যই:

- `http://` অথবা `https://` দিয়ে শুরু হতে হবে
- Publicly accessible হতে হবে
- Login/Password ছাড়া server থেকে access করা যেতে হবে

---

## ✅ 1. GitHub URL

GitHub-এর `blob` URL দেওয়া যাবে।

উদাহরণ:

```text
https://github.com/user/repository/blob/main/data.csv

GitHub RAW URL-ও দেওয়া যাবে:

https://raw.githubusercontent.com/user/repository/main/data.csv

GitHub blob URL হলে converter সেটিকে RAW URL-এ convert করার চেষ্টা করবে।

✅ 2. Hugging Face URL

Public Hugging Face dataset URL দেওয়া যাবে।

উদাহরণ:

https://huggingface.co/datasets/username/dataset

অথবা direct file URL:

https://huggingface.co/datasets/username/dataset/resolve/main/data.csv
✅ 3. Direct File URL

যেকোনো public server-এর direct dataset/file URL দেওয়া যাবে।

উদাহরণ:

https://example.com/data.csv
https://example.com/data.json
https://example.com/data.parquet

Direct downloadable file URL ব্যবহার করাই সবচেয়ে ভালো।

📦 Supported / Detectable File Formats

Converter বিভিন্ন ধরনের dataset/file format detect ও process করার চেষ্টা করে।

📊 Common Data Formats
CSV
TSV
TXT
JSON
JSONL / NDJSON
JSONL.GZ
Parquet
Excel (.xlsx, .xls, .xlsb)
ODS
🌐 Web / Text Formats
HTML tables
Markdown tables
XML
YAML / YML
📦 Archive Formats
ZIP
TAR
TAR.GZ
TGZ

Archive-এর ভিতরে supported dataset থাকলে সেটি extract করে process করার চেষ্টা করা হয়।

🗄️ Database / Scientific Formats
SQLite
NPY
NPZ
MAT
HDF
HDF5
ARFF
📄 Document Formats
PDF
DOCX
PPTX

Decoder/library available থাকলে readable বা structured content Excel-এ export করার চেষ্টা করা হয়।

🧩 Unknown File Extension

File-এর extension পরিচিত না হলেও converter file-এর content/signature দেখে format detect করার চেষ্টা করে।

উদাহরণ:

https://example.com/mydataset.xyz

অথবা extension ছাড়াই:

https://example.com/data

Unknown binary file হলে এবং সেটিকে dataset হিসেবে decode করা সম্ভব না হলে metadata/preview Excel fallback তৈরি হতে পারে।

📁 Multiple Files

একটি archive বা dataset থেকে একাধিক Excel file তৈরি হলে প্রতিটি file-এর জন্য আলাদা download option থাকতে পারে।

উদাহরণ:

dataset_1.xlsx    [Download Excel]

dataset_2.xlsx    [Download Excel]

dataset_3.xlsx    [Download Excel]

                  [Download All (.ZIP)]
❌ কোন URL কাজ নাও করতে পারে?
🔒 Private GitHub Repository

Private repository হলে login ছাড়া file access করা সম্ভব নয়।

উদাহরণ:

https://github.com/private-user/private-repository/blob/main/data.csv
🔐 Login Required Website

যদি URL খুলতে username/password লাগে, converter login bypass করবে না।

🤖 CAPTCHA Required

যদি website CAPTCHA চায়, converter CAPTCHA bypass করবে না।

🔑 API Key / Token Required

File download করতে API key বা authentication প্রয়োজন হলে public access ছাড়া conversion fail করতে পারে।

⚠️ সাধারণ Webpage বনাম Direct File URL

Dataset-এর webpage এবং actual dataset file একই জিনিস নয়।

যেমন:

https://example.com/datasets

এটি একটি webpage।

যদি actual file হয়:

https://example.com/files/dataset.csv

তাহলে দ্বিতীয় URL ব্যবহার করা ভালো।

☁️ Google Drive / Dropbox

সাধারণ Google Drive বা Dropbox sharing URL সবসময় direct file download URL নয়।

যেখানে সম্ভব direct downloadable file URL ব্যবহার করুন।

🧪 প্রথমবার Test করার নিয়ম

প্রথমে ছোট একটি public CSV file দিয়ে test করা ভালো।

উদাহরণ:

https://raw.githubusercontent.com/user/repository/main/test.csv

তারপর:

Paste URL
      ↓
Convert to Excel
      ↓
Wait
      ↓
Download Excel
⏳ Conversion চলার সময়

বড় dataset বা complex file হলে conversion করতে কিছু সময় লাগতে পারে।

বিশেষ করে:

বড় CSV/JSON
Parquet
ZIP/TAR archive
PDF
DOCX
PPTX
complex structured data

এসবের ক্ষেত্রে processing time বেশি হতে পারে।

⚠️ Convert button বারবার চাপবেন না

একবার Convert to Excel চাপার পর result আসা পর্যন্ত অপেক্ষা করুন।

🟣 Render Free Hosting

এই website বর্তমানে Render Free Web Service-এ চলছে।

অনেকক্ষণ website ব্যবহার না হলে Free service sleep করতে পারে।

তাই অনেকক্ষণ পরে প্রথমবার website ব্যবহার করলে কিছুটা বেশি সময় লাগতে পারে।

এটি সবসময় converter error নয়।

⚠️ Large Dataset

বর্তমানে website Free server resources ব্যবহার করছে।

তাই খুব বড় dataset অথবা একই সময়ে অনেক user conversion করলে:

Conversion slow হতে পারে
Memory error হতে পারে
Timeout হতে পারে
Server resource limit-এ পৌঁছাতে পারে

ছোট ও মাঝারি dataset দিয়ে শুরু করা সবচেয়ে ভালো।

🚨 Error হলে কী করবেন?

যদি কোনো error আসে, ভয় পাওয়ার কিছু নেই।

আমাকে নিচের ৩টি জিনিস পাঠান:

1️⃣ Error-এর Screenshot

যে error message দেখাচ্ছে তার screenshot দিন।

2️⃣ Dataset URL

আপনি যে URL দিয়েছিলেন সেটি পাঠান।

3️⃣ কখন Error হয়েছে

যেমন:

Convert to Excel চাপার পর error হয়েছে।

অথবা:

Download Excel চাপার পর error হয়েছে।

এই তথ্যগুলো দিলে error-এর কারণ identify করে fix করার চেষ্টা করা যাবে।

🔧 Common Errors
Invalid URL

URL সঠিক নয়।

সঠিক URL সাধারণত এমন হবে:

https://example.com/data.csv
403 Forbidden

Source server file access করতে দিচ্ছে না।

সম্ভবত permission/authentication প্রয়োজন।

404 Not Found

File বা URL পাওয়া যায়নি।

URL আবার check করুন।

429 Too Many Requests

Source server সাময়িকভাবে বেশি request block করেছে।

কিছুক্ষণ পরে আবার চেষ্টা করুন।

Unsupported dataset format

Converter file-এর format বুঝতে পারেনি।

করণীয়:

Error screenshot + dataset URL পাঠান।

No Excel file could be created

File download হয়েছে, কিন্তু usable table/data extract করা যায়নি।

করণীয়:

Error screenshot + dataset URL পাঠান।

Application Error

Server-side সমস্যা হতে পারে।

প্রথমে:
Page refresh করুন
কিছুক্ষণ অপেক্ষা করুন
আবার URL দিয়ে চেষ্টা করুন

তারপরও error থাকলে screenshot পাঠান।

🔒 Privacy & Security

এই website মূলত publicly accessible dataset/file URL process করার জন্য তৈরি।

Private/login-protected resource-এর authentication bypass করা হয় না।

CAPTCHA bypass করা হয় না।

Sensitive/private URL ব্যবহার না করাই ভালো।

💻 User-এর Computer-এ কী লাগবে?

User-এর computer-এ কিছু install করার প্রয়োজন নেই।

প্রয়োজন:
🌐 Internet connection
💻 / 📱 Browser
🔗 Public dataset URL
প্রয়োজন নেই:
❌ Python
❌ Terminal
❌ Command
❌ Dataset converter software
❌ Package installation

সব conversion server-side-এ করা হয়।

🖥️ Website কীভাবে কাজ করে?
                 USER
                   │
                   ▼
          Public Dataset URL
                   │
                   ▼
        🌐 Universal Converter
                   │
                   ▼
            Render Server
                   │
                   ▼
          Download Dataset
                   │
                   ▼
          Detect File Format
                   │
                   ▼
        Process Dataset/File
                   │
                   ▼
             Excel (.xlsx)
                   │
                   ▼
           Download Button
                   │
                   ▼
                  USER
🔄 Website Update

Website GitHub repository-এর সঙ্গে connected।

Code update করে GitHub repository-তে push করলে Render automatic deployment করতে পারে।

Code Update
     ↓
GitHub
     ↓
Render
     ↓
Automatic Deploy
     ↓
Website Updated
👨‍💻 Developer / Local Setup

যদি developer হিসেবে local machine-এ চালাতে চান:

pip install -r requirements.txt

তারপর:

python app.py

Local website:

http://localhost:5000

Production deployment-এর জন্য:

Procfile
render.yaml

included আছে।

📂 Project Structure
universal_dataset_website/
│
├── app.py
├── universal_dataset_engine.py
├── requirements.txt
├── Procfile
├── render.yaml
├── README.md
├── .gitignore
│
├── templates/
│   └── index.html
│
└── static/
    └── style.css
🌐 LIVE WEBSITE
👉 https://universal-dataset-to-excel.onrender.com

এই link-টাই অন্যদের দিতে পারবেন।

📌 QUICK USER GUIDE
1. Open:

https://universal-dataset-to-excel.onrender.com

2. Paste a public dataset/file URL.

3. Click:

Convert to Excel

4. Wait for conversion.

5. Click:

Download Excel
Important:
Use a public HTTP/HTTPS URL.

Private/login/CAPTCHA-protected URLs may not work.

If an error appears:
send the error screenshot + URL to the administrator.
🎯 এক লাইনে
🔗 Public URL দিন → 🔄 Convert to Excel চাপুন → 📥 Excel Download করুন
🌐 https://universal-dataset-to-excel.onrender.com
