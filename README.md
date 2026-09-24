# Brüt Ücret Okuyucu

MUHSGK PDF'lerindeki personel bazlı **Çalışmaya İlişkin Brüt Ücret Tutarı Toplamı**, **Ücret Matrahı** ve **Terkin Edilecek Tutar** değerlerini okuyup tek bir Excel dosyasına döker.

## Özellikler

- Birden fazla PDF'i tek seferde işler.
- Her PDF'in satırlarının altına **o PDF'e ait ayrı bir TOPLAM satırı** ekler.
- `Kontrol` sayfasında, okunan terkin toplamını PDF'teki `TOPLAM` ile karşılaştırır (**TAMAM** / **UYUMSUZ**).
- Basit bir masaüstü arayüzü var, Windows için tek dosyalık `.exe` olarak derlenebilir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

**Arayüz:**

```bash
python arayuz.py
```

PDF'leri ya da bir klasörü seçin, **Excel Oluştur**'a basın.

**Komut satırı:**

```bash
python brut_ucret.py                  # bulunduğu klasördeki tüm PDF'ler
python brut_ucret.py a.pdf b.pdf      # belirli PDF'ler
```

Çıktı: `brut_ucret.xlsx`

## Excel çıktısı

| Sayfa | İçerik |
|---|---|
| **Brüt Ücretler** | Personel satırları (dosya, sayfa, ad soyad, TC, işe başlama, brüt, matrah, terkin). Her PDF'in altında o PDF'in toplam satırı var. |
| **Kontrol** | PDF başına personel sayısı, brüt toplam ve terkin kontrolü |

## Windows .exe

```bat
build.bat
```

Çıktı: `dist\BrutUcretOkuyucu.exe`

> PDF ve Excel dosyaları kişisel veri (TC no, ücret) içerdiği için `.gitignore` ile repoya eklenmez.
