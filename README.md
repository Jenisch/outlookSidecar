# outlookSidecar

Yerel Outlook kurulumundaki operasyon maillerinden güncel icap raporu üretmek için kullanılan bir komut satırı aracı.

## Özellikler

- Outlook uygulamasına giriş yapmadan, Windows üzerindeki mevcut oturumdan mesaj içeriklerini okur.
- Gövdesinde `25OA` ile başlayan dosya numaraları bulunan mailleri ayrıştırarak durum, başlık, detay ve uçuş bilgilerini çıkarır.
- JSON veya biçimlendirilmiş metin halinde rapor üretir.
- Geliştirme ve test için basit metin dosyalarından veri okuma imkânı sunar.

## Kurulum

1. Python 3.10+ sürümünü kullanın.
2. Windows üzerinde `pywin32` paketi yüklü olmalıdır:

   ```bash
   pip install pywin32
   ```

3. Depoyu klonlayın ve proje kök dizinine geçin.
4. Paket bağımlılıklarını yüklemek için:

   ```bash
   pip install -e .
   ```

## Kullanım

Komut satırından çalıştırmak için:

```bash
python -m outlook_sidecar.cli --mode outlook --path "Inbox" --days 2
```

### Parametreler

- `--mode`: `outlook` (varsayılan) ya da `file`. `file` seçeneği, geliştirme amaçlı düz metin dosyalarını okur.
- `--path`: Outlook klasör yolu (`Inbox/Alt Klasör` gibi) veya metin dosyasının yolu.
- `--days`: Kaç günlük maillerin işleneceği. Boş bırakırsanız 2 gündür.
- `--limit`: En fazla kaç mail okunacağını sınırlar.
- `--json`: Çıktıyı JSON formatında üretir.

### Örnek

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
