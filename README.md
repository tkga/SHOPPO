# Pokémon GO Shop

## รันบนเครื่องตัวเอง (ทดสอบก่อน deploy)

```
npm install
npm run dev
```

แล้วเปิด http://localhost:5173

## Build สำหรับ deploy

```
npm run build
```

จะได้โฟลเดอร์ `dist/` พร้อมอัปโหลด

## Deploy ขึ้น GitHub Pages (อัตโนมัติด้วย GitHub Actions)

โปรเจกต์นี้มีไฟล์ `.github/workflows/deploy.yml` ให้แล้ว ทำตามนี้ครั้งเดียว:

1. Push โค้ดทั้งโฟลเดอร์นี้ (ทุกไฟล์ รวม `.github/`) ขึ้น branch `main` ของ repo
2. เข้า repo บน GitHub → **Settings → Pages**
3. ตรง **Build and deployment → Source** เลือก **"GitHub Actions"** (ไม่ใช่ "Deploy from a branch")
4. รอสักครู่ ไปที่แท็บ **Actions** ของ repo จะเห็น workflow "Deploy to GitHub Pages" รันอยู่ พอเสร็จ (เขียวติ๊กถูก) เว็บจะขึ้นที่ `https://<username>.github.io/<repo>/`
5. ทุกครั้งที่ push โค้ดใหม่เข้า `main` มันจะ build และ deploy ให้อัตโนมัติ ไม่ต้อง `npm run build` มือเอง

**ข้อควรระวัง:** ห้าม deploy ไฟล์ซอร์ส (`src/`, `index.html` ตัวเปล่าๆ) ขึ้น Pages ตรงๆ เพราะเบราว์เซอร์รันไฟล์ `.jsx` หรือ resolve `import react` เองไม่ได้ (จะเจอจอดำ) ต้องผ่าน `npm run build` ให้ได้โฟลเดอร์ `dist/` (ไฟล์ JS ที่ bundle แล้ว) ก่อนเสมอ — ซึ่ง workflow ด้านบนทำให้อัตโนมัติแล้ว

## Deploy ขึ้น Vercel (แนะนำ, ฟรี)

1. สมัคร/ล็อกอิน https://vercel.com ด้วย GitHub
2. อัปโหลดโฟลเดอร์นี้ขึ้น GitHub repo (หรือใช้ `vercel` CLI: `npx vercel`)
3. Vercel จะ detect Vite อัตโนมัติ กด Deploy
4. ได้ลิงก์ เช่น `pokemon-go-shop.vercel.app`

## ติดตั้งบนมือถือ (PWA)

1. เปิดลิงก์ที่ deploy แล้วบน Safari (iOS) หรือ Chrome (Android)
2. iOS: กดปุ่มแชร์ → "เพิ่มไปยังหน้าจอโฮม"
   Android: กดเมนู 3 จุด → "เพิ่มไปยังหน้าจอโฮม" / "ติดตั้งแอป"
3. จะมีไอคอนแอปที่หน้าจอโฮม เปิดแล้วเต็มจอเหมือนแอปจริง

## เรื่องข้อมูล

- ข้อมูลเก็บด้วย **IndexedDB** ในเครื่อง/เบราว์เซอร์นั้นๆ (เดิมใช้ `localStorage` ซึ่งมีโควตาแค่ 5-10MB — พอมีออเดอร์เยอะขึ้นพร้อมรูปสลิปที่แนบ อาจ save ไม่ผ่านแบบเงียบๆ จึงย้ายมาใช้ IndexedDB ที่โควตาสูงกว่ามาก)
- รูปสลิป/รูปกิจกรรมที่แนบต่อออเดอร์ถูกบีบอัดเป็น JPEG อัตโนมัติก่อนบันทึก (ไม่ใช่ PNG แบบเดิม) เพื่อไม่ให้ไฟล์ข้อมูลบวมเร็วเกินไปเมื่อออเดอร์เยอะขึ้นเรื่อยๆ
- ใช้ปุ่ม **Backup ข้อมูล (.json)** ในหน้าตั้งค่าเป็นประจำ เพื่อกันข้อมูลหายกรณีล้างข้อมูลเบราว์เซอร์หรือเปลี่ยนเครื่อง
- หรือเปิดใช้ **ซิงค์ Google Sheets + Drive อัตโนมัติ** (ดูหัวข้อถัดไป) เป็นสำรองชั้นที่สอง ไม่ต้องพึ่งการกด backup มือทุกครั้ง

## ตั้งค่า Google Sync (สำรองข้อมูลอัตโนมัติ ฟรี ไม่มีค่าใช้จ่าย)

เมื่อเชื่อมต่อแล้ว แอปจะ: เขียนข้อมูลลูกค้า/ออเดอร์/ไอดีเกม/การเงินเป็นแถวใน **Google Sheets** ของร้านเอง และอัปโหลดรูปสลิป/รูปกิจกรรมขึ้น **Google Drive** ของร้านเอง (โฟลเดอร์แยกต่างหาก) แล้วอ้างอิงรูปนั้นกลับมาโชว์ในตารางด้วยสูตร `=IMAGE(...)` ทำงานอัตโนมัติทุกครั้งที่มีการแก้ไขข้อมูล (หน่วงเวลาประมาณ 45 วินาทีหลังหยุดแก้ เพื่อไม่ให้ยิง request ถี่เกินไป) และมีปุ่ม "ซิงค์เดี๋ยวนี้" ให้กดเองได้ทุกเมื่อในหน้าตั้งค่า

ต้องตั้งค่าใน Google Cloud Console **ครั้งเดียว** (ฟรี ไม่ต้องผูกบัตรเครดิต):

1. เข้า https://console.cloud.google.com/ ด้วย Google account ของร้าน แล้วสร้างโปรเจกต์ใหม่ (หรือใช้โปรเจกต์เดิมก็ได้)
2. ไปที่ **APIs & Services → Library** ค้นหาแล้วกด "Enable" ทั้งสองตัว: **Google Sheets API** และ **Google Drive API**
3. ไปที่ **APIs & Services → OAuth consent screen** เลือก **External** → กรอกชื่อแอป (เช่น "Pokémon GO Shop") + อีเมลติดต่อ → บันทึก ไม่ต้องกด submit เพื่อ verify เพราะใช้เองในร้าน
4. ในหน้าเดียวกัน ไปแท็บ **Test users** → กด Add users → ใส่อีเมล Gmail ของคนที่จะใช้แอปนี้ (ตัวเอง/ทีมงาน ได้สูงสุด 100 คน) — ข้ามขั้นนี้ไม่ได้ ไม่งั้นจะล็อกอินไม่ผ่าน
5. ไปที่ **APIs & Services → Credentials** → **Create Credentials → OAuth client ID** → เลือกประเภท **Web application**
6. ที่ช่อง **Authorized JavaScript origins** ใส่โดเมนที่ deploy จริง เช่น `https://<username>.github.io` หรือ `https://pokemon-go-shop.vercel.app` และเพิ่ม `http://localhost:5173` ไว้ด้วยเผื่อทดสอบในเครื่อง
7. กด Create จะได้ **Client ID** (รูปแบบ `xxxxxxxxxxxx.apps.googleusercontent.com`) — คัดลอกมาวางในแอป หน้า **ตั้งค่า → สำรองข้อมูลอัตโนมัติ (Google Sheets + Drive)** แล้วกด "เชื่อมต่อ Google"

**หมายเหตุ:**
- ตอนล็อกอินครั้งแรก Google จะโชว์หน้าเตือน "This app isn't verified" เพราะยังไม่ได้ submit ให้ Google ตรวจ (ปกติสำหรับแอปที่ใช้เองในทีมเล็กๆ ไม่จำเป็นต้อง verify) — กด **Advanced → Go to (ชื่อแอป) (unsafe)** เพื่อดำเนินการต่อได้เลย ปลอดภัย เพราะเป็นแอปที่คุณสร้าง OAuth client เองกับ Google โดยตรง
- ทุกคนที่จะล็อกอินได้ต้องถูกเพิ่มไว้ในขั้นตอนที่ 4 (Test users) ก่อน
- ไม่มีเซิร์ฟเวอร์กลางใดๆ — ข้อมูลวิ่งจากเบราว์เซอร์ตรงไป Google API ของบัญชีร้านเองเท่านั้น
