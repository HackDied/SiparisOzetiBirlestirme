# Siparis Ozeti Birlestirme Araci

Birden fazla siparis ozeti Excel dosyasini tek bir Excel'de birlestiren, alis fiyatlarini ve kar/zarar analizini iceren masaustu uygulamasi.

## Ozellikler

- **Coklu Dosya Birlestirme** - Birden fazla siparis ozeti `.xlsx` dosyasini tek bir Excel'de birlestirir
- **Alis & Satis Fiyatlari** - Her urun icin birim maliyet (U.COST) ve toplam maliyet (T.COST) hesaplar
- **Doviz Kuru Donusumu** - TL, EUR, USD arasi otomatik kur donusumu (Frankfurter API / ECB verileri)
- **Kar/Zarar Analizi** - Grand Summary'de toplam satis, toplam alis, indirim ve kar/zarar gosterir
- **Firma Indirimi** - Siparis ozetlerinde belirtilmeyen firma indirim oranini GUI uzerinden girebilme
- **Gemi Ismi Tespiti** - Her dosyadan gemi ismini otomatik okur, cikti dosya adinda kullanir
- **Surukleme & Birakma** - Dosya surukleyerek ekleme destegi (tkinterdnd2)
- **Profesyonel Excel Ciktisi** - Renkli banner, stilize basliklar ve formul bazli hesaplamalar

## Kurulum

### Exe (Onerilen)

`dist/SiparisOzetiBirlestirme.exe` dosyasini indirip dogrudan calistirabilirsiniz. Python kurulumu gerektirmez.

## Kullanim

1. **Dosya Sec** - Siparis ozeti Excel dosyalarini ekleyin
2. **Indirim Orani** - Firma indirim oranini girin (varsa)
3. **Doviz Kurlari** - "Guncel Kurlari Cek" ile online kurlari alin veya manuel girin
4. **Birlestir** - "Dosyalari Birlestir" butonuna tiklayin

## Girdi Excel Formati

Arac asagidaki siparis ozeti yapisini bekler:

| Hucre | Icerik |
|-------|--------|
| 15B | Gemi ismi |
| Satir 18-20, Sutun H-I | DATE, RFQ REF, QTN REF |
| Satir 21 | Baslik satiri (NO, DESCRIPTION, CODE, ...) |
| Satir 23+ | Veri satirlari |
| COST sutunu | Birim maliyet (ornek: "21500.00 TL") |

## Cikti Excel

- **Dosya adi**: `GemiIsmi_GG-AA-YYYY.xlsx`
- **Banner**: Baslik, gemi ismi, olusturma tarihi
- **Her siparis blogu**: Bilgi satiri, baslik, veri satirlari, TOTAL + COST TOTAL
- **Grand Summary**: Toplam Satis, Toplam Alis, Indirim, Final Satis Tutari, Kar/Zarar

## Ekran Goruntusu

```
+------------------------------------------+
|       MERGED ORDER SUMMARY               |
|       VESSEL: MSC NINA F                 |
|  Generated: 16.02.2026 | Files: 3       |
+------------------------------------------+
| NO | DESCRIPTION | ... | U.COST | T.COST |
|----|-------------|-----|--------|--------|
|  1 | Item A      | ... | 150.00 | 300.00 |
|  2 | Item B      | ... |  85.00 | 170.00 |
+------------------------------------------+
|          TOPLAM SATIS:    €12,500.00     |
|          TOPLAM ALIS:      €8,200.00     |
|          KAR / ZARAR:      €4,300.00     |
+------------------------------------------+
```
