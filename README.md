# OST → PST Donusturucu

Outlook **OST** dosyalarini, **Outlook'un sorunsuz acabilecegi PST** dosyalarina
donusturen, sade ve okunabilir bir masaustu programi (Tkinter arayuzu, Turkce).

> **Neden iki motor var?** Bir OST dosyasini Outlook'un sorunsuz acabilecegi bir
> PST'ye cevirmenin en guvenilir yolu, PST'yi **Outlook'un kendisine olusturtmaktir**.
> Bu yuzden program, ortaminiza gore en iyi yontemi otomatik secer.

---

## Ekran / Kullanim ozeti

Program iki sekme sunar.

### 1) OST Dosyasi → PST  *(ana akis)*
Diskteki bir `.ost` dosyasini **elle secersiniz** — dosyanin bir Outlook
profiline bagli olmasi gerekmez (eski yedekler, orphan dosyalar dahil).

1. **Dosya sec...** ile kaynak `.ost` dosyasini secin.
2. **Cikti bicimi** secin: PST (Outlook gerekli) / EML klasoru / MBOX.
3. **Klasor sec...** ile yalnizca bir **hedef klasor** secin.
4. **Donustur**. Cikti dosyasi, kaynak OST adina gore **bu klasorde otomatik
   olusturulur** (orn. `Outlook.ost` → `Outlook.pst`). Ayni ada sahip dosya
   varsa sonuna sayi eklenir.

> PST cikti icin Outlook gerekir (dosya libpff ile okunur, gercek PST Outlook
> ile uretilir). Outlook yoksa **EML klasoru** veya **MBOX** secerek icerigi
> kurtarabilirsiniz.

### 2) Posta Kutusu → PST (Outlook hesabi)  *(en yuksek sadakat)*
Outlook'ta **tanimli** bir hesabin OST onbellegini tam sadakatle PST'ye kopyalar.
Hesap listesi bos gelirse Outlook'ta yapilandirilmis hesap yok demektir; bu
durumda 1. sekmeyi kullanin. Hedef yine yalnizca bir **klasordur**; PST dosyasi
posta kutusu adina gore otomatik olusturulur.

Alttaki **Durum** panelinde ilerleme cubugu ve ayrintili gunluk gosterilir.

---

## Kurulum

Gereken: **Python 3.8+** (Windows onerilir).

```bat
pip install -r requirements.txt
```

- `pywin32` — Windows'ta Outlook ile gercek PST uretimi icin (COM/MAPI).
- `libpff-python` — diskteki `.ost` dosyalarini okumak icin.
- **Tkinter** Python ile birlikte gelir; ayrica kurulmaz.

Hicbiri zorunlu degildir; eksik motorlar arayuzde otomatik devre disi kalir
(alt durum cubugunda ✓ / ✗ gosterilir).

## Calistirma

```bat
python run.py
```

veya Windows'ta `OST-PST-Donusturucu-Baslat.bat` dosyasina cift tiklayin.

---

## Hangi yontem ne zaman calisir?

| Senaryo | Outlook kurulu | libpff | Sonuc |
|---|:---:|:---:|---|
| Hesap Outlook'ta tanimli, OST onbellegini PST yapmak | ✓ | – | **Tam sadakatli PST** (Sekme 1) |
| Diskte orphan `.ost`, gercek PST isteniyor | ✓ | ✓ | **PST** (Sekme 2 → PST) |
| Diskte orphan `.ost`, Outlook yok | – | ✓ | **EML / MBOX** kurtarma (Sekme 2) |

---

## Onemli teknik notlar (durust ozet)

- **OST dosyasi dogrudan Outlook'a "store" olarak eklenemez** — Microsoft buna
  izin vermez. Bu yuzden Sekme 1, yeni bos bir PST olusturup posta kutusu
  icerigini ona **kopyalar**; Sekme 2 ise dosyayi libpff ile okuyup Outlook'a
  yeniden **aktarir**.
- **PST → OST ters donusum desteklenmez.** OST dosyasi yalnizca Exchange/IMAP
  senkronu ile olusur ve profile baglidir; "tasinabilir bir OST" kavrami yoktur.
  Bu nedenle program yalnizca **OST → PST** yonune odaklanir.
- Uretilen PST'ler **Unicode** bicimindedir (eski 2 GB ANSI siniri yoktur).
- EML/MBOX ciktilari, basliklar (transport headers), govde (duz metin + HTML)
  ve ekleri korur; ancak takvim/kisi gibi MAPI'ye ozgu bazi ozellikler bu
  bicimlerde tam temsil edilemeyebilir. Tam sadakat icin Sekme 1'i kullanin.

---

## Proje yapisi

```
.
├── run.py                  # Baslatma noktasi
├── requirements.txt
├── OST-PST-Donusturucu-Baslat.bat
├── converter/
│   ├── orchestrator.py     # Hangi motorun kullanilacagina karar verir
│   ├── engine_outlook.py   # Outlook COM/MAPI motoru (gercek PST)
│   └── engine_libpff.py    # libpff/pypff okuyucu (EML/MBOX cikti)
└── gui/
    └── app.py              # Tkinter arayuzu
```
