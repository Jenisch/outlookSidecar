# outlookSidecar

Yerel Outlook kurulumundaki operasyon maillerinden güncel icap raporu üretmek için kullanılan masaüstü uygulaması.

## Özellikler

- Outlook uygulamasına giriş yapmadan, Windows üzerindeki mevcut oturumdan mesaj içeriklerini okur.
- Gövdesinde `25OA` ile başlayan dosya numaraları bulunan mailleri ayrıştırarak durum, başlık, detay ve uçuş bilgilerini çıkarır.
- JSON veya biçimlendirilmiş metin halinde rapor üretir.
- Geliştirme ve test için basit metin dosyalarından veri okuma imkânı sunar.
- Outlook’a bağlanmaya gerek kalmadan aynı bilgileri içeren kullanıcı dostu bir arayüz sağlar.

## Kurulum

### Hızlı Başlangıç (Windows)

1. Python 3.10+ sürümünün sisteminizde kurulu olduğundan emin olun.
2. Depoyu klonlayın ve proje klasörünü Windows gezgininde açın.
3. `launch_outlook_sidecar.bat` dosyasına çift tıklayın. Dosya otomatik olarak `requirements.txt` içindeki bağımlılıkları
   yükler ve ardından grafik arayüzünü başlatır.

> **Not:** Betik, `py` komutunu bulamazsa `python` komutunu kullanarak sistemdeki varsayılan Python kurulumunu çağırır.

### Manuel kurulum (geliştiriciler için)

1. Python 3.10+ sürümünü kullanın.
2. Depoyu klonlayın ve proje kök dizinine geçin.
3. Paket bağımlılıklarını yüklemek için:

   ```bash
   pip install -r requirements.txt
   pip install -e .
   ```

## Kullanım

### Grafik Arayüzü

`launch_outlook_sidecar.bat` betiği çalıştırıldıktan sonra açılan arayüz üzerinden:

- `Microsoft Outlook` seçeneği ile yerel Outlook kurulumuna bağlanabilir, klasör yolunu, kaç günlük veri alınacağını ve
  isteğe bağlı mesaj limitini belirleyebilirsiniz.
- `Local Text File` seçeneğiyle düz metin dosyalarını seçerek hızlıca test yapabilirsiniz.
- `Load Cases` butonuna basarak mailleri ayrıştırabilir, listeden vaka seçip detayları inceleyebilirsiniz.

### Komut Satırı

Komut satırından çalıştırmak için:

```bash
python -m outlook_sidecar.cli --mode outlook --path "Inbox" --days 2
```

#### Parametreler

- `--mode`: `outlook` (varsayılan) ya da `file`. `file` seçeneği, geliştirme amaçlı düz metin dosyalarını okur.
- `--path`: Outlook klasör yolu (`Inbox/Alt Klasör` gibi) veya metin dosyasının yolu.
- `--days`: Kaç günlük maillerin işleneceği. Boş bırakırsanız 2 gündür.
- `--limit`: En fazla kaç mail okunacağını sınırlar.
- `--json`: Çıktıyı JSON formatında üretir.

#### Örnek

Aşağıdaki komut, verilen metin dosyasından bir rapor üretir:

```bash
python -m outlook_sidecar.cli --mode file --path sample_mail.txt
```

## Geliştirme

### Testler

Birlikte gelen örnek veri için parser testlerini çalıştırmak üzere:

```bash
pytest
```

## Sınırlamalar

- Outlook API erişimi yalnızca Windows üzerinde ve Outlook uygulamasının kurulu olduğu makinelerde çalışır.
- Uçuş satırları tahmini olarak ayrıştırılır; format dışına çıkan girdiler `raw` alanında aynen saklanır.
